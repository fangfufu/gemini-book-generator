import hashlib
import logging
import pathlib
import sys
import time

import requests

from book_generator.config import load_config
from book_generator.constants import DEFAULT_WRITING_TONE
from book_generator.common_generator import (
    determine_gender_from_name,
    format_character_list_for_prompt,
    generate_back_matter,
    generate_book_blurb,
    generate_book_title,
    generate_chapter_summary,
    generate_character_list,
    generate_front_matter,
    generate_overall_summary,
    generate_random_gender,
    generate_random_name,
    generate_random_topic,
    generate_setting,
    generate_writing_tone,
    is_book_non_fiction,
    save_summary_to_markdown,
)
from book_generator.fiction_generator import (
    generate_overall_story,
    generate_fiction_chapter_outline,
    generate_fiction_chapter_content,
)
from book_generator.non_fiction_generator import (
    generate_chapter_outline,
    generate_section_titles,
    generate_section_content,
)
from book_generator.docx_builder import assemble_docx, assemble_marketing_docx
from book_generator.llm_api import setup_environment
from book_generator.utils import sanitize_filename


def run_generation_process_fiction(config, output_base_dir, equation_image_dir):
    logging.info("Starting fiction book generation process...")
    start_time = time.time()
    generation_params = config.get("generation_params", {})
    book_title = generation_params.get("book_title")
    writing_tone = generation_params.get("writing_tone")

    # 1. Create a list of characters.
    logging.info("--- Generating Character List ---")
    character_list = generate_character_list(config, book_title)
    character_context_for_prompts = format_character_list_for_prompt(character_list)

    # 2. Create an overall story based on the prompt.
    logging.info("--- Generating Overall Story ---")
    overall_story = generate_overall_story(config, character_context_for_prompts)
    if not overall_story:
        logging.critical("Fatal: Failed to generate the overall story. Exiting.")
        sys.exit(1)

    # 3. Break down the story into chapters.
    logging.info("--- Generating Fiction Chapter Outline (Titles and Summaries) ---")
    fiction_chapter_count = generation_params.get("fiction_chapter_count", 20)
    chapters = generate_fiction_chapter_outline(config, overall_story, character_context_for_prompts, fiction_chapter_count)
    if not chapters:
        logging.critical("Fatal: Failed to generate chapter outline. Exiting.")
        sys.exit(1)

    # 4. Write each chapter.
    logging.info("--- Generating Fiction Chapter Content ---")
    body_matter = {}
    previous_chapter_title = ""
    previous_chapter_summary = ""

    for i, chapter in enumerate(chapters):
        chapter_title = chapter.get("title")
        chapter_summary = chapter.get("summary")
        logging.info(f"--- Generating Content for Chapter {i+1}: {chapter_title} ---")

        chapter_content = generate_fiction_chapter_content(
            config,
            overall_story,
            previous_chapter_title,
            previous_chapter_summary,
            chapter_title,
            character_context_for_prompts,
        )

        body_matter[chapter_title] = [{"title": "", "content": chapter_content}]

        previous_chapter_title = chapter_title
        previous_chapter_summary = chapter_summary

    summary_parts = []
    for i, chapter in enumerate(chapters):
        title = chapter.get("title", "")
        summary = chapter.get("summary", "").strip()
        if title and summary:
            summary_parts.append(f"Chapter {i+1} ('{title}'): {summary}")

    if summary_parts:
        summary_context = "Chapter summaries:\n" + "\n\n".join(summary_parts)
    else:
        summary_context = "[No chapter summaries available]"

    # --- Generate Overall Summary and Save to Markdown ---
    overall_summary_text = generate_overall_summary(config, book_title, summary_context)
    save_summary_to_markdown(book_title, overall_summary_text, output_base_dir)
    # --- End Overall Summary Generation ---

    logging.info("--- Generating Front Matter, Back Matter, and Marketing Content ---")
    front_matter_content = generate_front_matter(
        config,
        book_title,
        generation_params["author_name"],
        writing_tone,
        summary_context,
    )

    back_matter_content = generate_back_matter(
        config,
        book_title,
        generation_params["author_name"],
        generation_params["author_gender"],
        writing_tone,
        summary_context,
    )

    main_filename_stem = sanitize_filename(book_title)
    if main_filename_stem != "sanitized_empty":
        blurb_text = generate_book_blurb(
            config, book_title, summary_context, writing_tone
        )
        assemble_marketing_docx(
            config,
            back_matter_content,
            blurb_text,
            summary_context,
            main_filename_stem,
            equation_image_dir,
            output_base_dir,
        )
    else:
        logging.error(
            f"Could not generate valid filename stem from title '{book_title}'. Skipping marketing DOCX."
        )

    main_docx_filename = assemble_docx(
        config,
        front_matter_content,
        body_matter,
        back_matter_content,
        book_title,
        equation_image_dir,
        output_base_dir,
    )

    end_time = time.time()
    logging.info(
        f"Fiction book generation process finished in {end_time - start_time:.2f} seconds."
    )
    if main_docx_filename:
        logging.info(f"Main book saved as: {main_docx_filename}")
    else:
        logging.error("Main book DOCX file failed to save.")


def run_generation_process_non_fiction(config, output_base_dir, equation_image_dir):
    logging.info("Starting non-fiction book generation process...")
    start_time = time.time()
    generation_params = config.get("generation_params", {})
    book_title = generation_params.get("book_title")
    writing_tone = generation_params.get("writing_tone")

    # --- Generate Character List (if enabled) ---
    generate_character_list(config, book_title)  # API call inside if enabled
    # The result (or None) is stored in config['generation_params']['character_list']

    # --- Generate Chapter Outline (potentially using characters) ---
    # Prepare character context string
    character_context_for_prompts = format_character_list_for_prompt(
        config["generation_params"].get("character_list")
    )

    chapter_titles = generate_chapter_outline(config, character_context_for_prompts)

    if chapter_titles:
        formatted_outline = "\n".join(
            f"{i+1}. {title}" for i, title in enumerate(chapter_titles)
        )
        logging.info(f"Generated Chapter Outline:\n{formatted_outline}")
    else:
        logging.error("Chapter outline generation resulted in no titles. Exiting.")
        sys.exit(1)  # Cannot proceed without chapters

    # --- Pass 1: Generate Chapter Summaries ---
    logging.info("--- Starting Pass 1: Generating Chapter Summaries ---")
    chapter_summaries = {}
    previous_summaries_list = []  # Keep track of summaries generated so far

    for i, chap_title in enumerate(chapter_titles):
        logging.info(f"--- Generating Summary for Chapter {i+1}: {chap_title} ---")

        # --- Create context string from previous summaries ---
        previous_summaries_context = ""
        if previous_summaries_list:
            # Format the context clearly, e.g., as a list
            context_parts = [
                f"- Chapter {idx+1}: {s}"
                for idx, s in enumerate(previous_summaries_list)
            ]
            previous_summaries_context = "\n".join(context_parts)
            logging.debug(
                f"Providing context of {len(previous_summaries_list)} previous summaries for chapter '{chap_title}'."
            )
        # --- End context creation ---

        # Pass the context to the generation function
        summary = generate_chapter_summary(
            config,
            chap_title,
            writing_tone,
            previous_summaries_context,
            character_context_for_prompts,
        )
        chapter_summaries[chap_title] = summary

        # --- Add the generated summary to the list for the next iteration's context ---
        # Only add valid, non-placeholder summaries to the context for subsequent chapters
        cleaned_summary_for_context = summary.strip()
        if cleaned_summary_for_context and not cleaned_summary_for_context.startswith(
            "Placeholder summary"
        ):
            previous_summaries_list.append(cleaned_summary_for_context)
        # --- End adding summary to list ---

    # --- Prepare Summary Context for subsequent prompts ---
    summary_context = "[No chapter summaries available]"  # Default
    summary_parts = []
    for i, chap_title in enumerate(chapter_titles):
        summary = chapter_summaries.get(chap_title, "").strip()
        if summary and not summary.startswith("Placeholder summary"):
            summary_parts.append(f"Chapter {i+1} ('{chap_title}'): {summary}")
    if summary_parts:
        summary_context = "Chapter summaries:\n" + "\n\n".join(summary_parts)
        logging.debug(f"Generated summary context:\n{summary_context}")
    else:
        logging.error(
            "No valid chapter summaries generated to create summary context."
        )

    # --- Generate Overall Summary and Save to Markdown ---
    overall_summary_text = generate_overall_summary(config, book_title, summary_context)
    save_summary_to_markdown(book_title, overall_summary_text, output_base_dir)
    # --- End Overall Summary Generation ---

    # --- Generate Front/Back Matter & Marketing Content (using summaries) ---
    logging.info("--- Generating Front Matter, Back Matter, and Marketing Content ---")
    front_matter_content = generate_front_matter(
        config,
        book_title,
        generation_params["author_name"],
        writing_tone,
        summary_context,
    )  # Multiple API calls inside

    back_matter_content = generate_back_matter(
        config,
        book_title,
        generation_params["author_name"],
        generation_params["author_gender"],
        writing_tone,
        summary_context,
    )  # Multiple API calls inside

    main_filename_stem = sanitize_filename(book_title)
    if main_filename_stem != "sanitized_empty":
        blurb_text = generate_book_blurb(
            config, book_title, summary_context, writing_tone
        )  # API call
        assemble_marketing_docx(
            config,
            back_matter_content,
            blurb_text,
            summary_context,
            main_filename_stem,
            equation_image_dir,
            output_base_dir,
        )
    else:
        logging.error(
            f"Could not generate valid filename stem from title '{book_title}'. Skipping marketing DOCX."
        )

    # --- Pass 2: Generate Section Content (using summaries) ---
    logging.info("--- Starting Pass 2: Generating Section Content ---")
    body_matter = {}
    is_fiction_book = generation_params.get("is_fiction", False)

    for i, chap_title in enumerate(chapter_titles):
        logging.info(f"--- Processing Chapter {i+1}: {chap_title} ---")
        chapter_summary = chapter_summaries.get(
            chap_title, f"Placeholder summary for chapter '{chap_title}'."
        )
        body_matter[chap_title] = []

        if is_fiction_book:
            logging.info(
                f"Book is fiction. Generating content for chapter '{chap_title}' as a single block."
            )
            fiction_section_title = "Narrative"
            chapter_content_as_single_block = generate_section_content(
                config,
                chap_title,
                fiction_section_title,
                1,
                1,
                chapter_summary,
                writing_tone,
                character_context_for_prompts,
            )
            body_matter[chap_title].append(
                {
                    "title": "",
                    "content": chapter_content_as_single_block,
                }
            )
            if (
                not chapter_content_as_single_block
                or "Content generation failed" in chapter_content_as_single_block
            ):
                logging.error(
                    f"Content generation potentially failed for fiction chapter '{chap_title}' (treated as single section)."
                )
        else:
            section_titles = generate_section_titles(
                config,
                chap_title,
                chapter_summary,
                chapter_titles,
                chapter_summaries,
                character_context_for_prompts,
            )

            if not section_titles:
                logging.error(
                    f"No section titles generated for chapter '{chap_title}'. Adding chapter summary as placeholder content."
                )
                body_matter[chap_title].append(
                    {"title": "Chapter Overview", "content": chapter_summary}
                )
                continue
            else:
                formatted_sections = "\n".join(
                    f"  {j+1}. {title}" for j, title in enumerate(section_titles)
                )
                logging.info(
                    f"Generated {len(section_titles)} section titles for chapter '{chap_title}':\n{formatted_sections}"
                )

            for j, sec_title in enumerate(section_titles):
                section_content = generate_section_content(
                    config,
                    chap_title,
                    sec_title,
                    j + 1,
                    len(section_titles),
                    chapter_summary,
                    writing_tone,
                    character_context_for_prompts,
                )
                body_matter[chap_title].append(
                    {"title": sec_title, "content": section_content}
                )
                if (
                    not section_content
                    or "Content generation failed" in section_content
                ):
                    logging.error(
                        f"Content generation potentially failed for Chapter '{chap_title}', Section '{sec_title}'."
                    )

    logging.info("Finished generating all body matter.")

    main_docx_filename = assemble_docx(
        config,
        front_matter_content,
        body_matter,
        back_matter_content,
        book_title,
        equation_image_dir,
        output_base_dir,
    )

    end_time = time.time()
    logging.info(
        f"Book generation process finished in {end_time - start_time:.2f} seconds."
    )
    if main_docx_filename:
        logging.info(f"Main book saved as: {main_docx_filename}")
    else:
        logging.error("Main book DOCX file failed to save.")


def run_generation_process():
    logging.info("Starting book generation process...")
    start_time = time.time()
    config = load_config()

    api_settings = config.get("api_settings", {})

    api_provider = api_settings.get("provider", "gemini")
    if api_provider == "gemini":
        setup_environment()
        logging.info("Gemini API provider configured.")
    elif api_provider == "ollama":
        ollama_settings = config.get("ollama_settings", {})
        ollama_base_url = ollama_settings.get("base_url", "http://localhost:11434")
        try:
            logging.info(
                f"Ollama API provider selected. Attempting to connect to: {ollama_base_url}"
            )
            requests.get(
                f"{ollama_base_url.rstrip('/')}/api/tags", timeout=5
            ).raise_for_status()
            logging.info(
                f"Successfully connected to Ollama server at {ollama_base_url}"
            )
        except requests.exceptions.RequestException as e:
            logging.error(
                f"Error: Could not connect or communicate with Ollama server at {ollama_base_url}. Error: {e}. Please ensure Ollama is running and accessible."
            )
            sys.exit(1)
    else:
        logging.error(
            f"Unsupported API provider '{api_provider}' specified in config. Supported: 'gemini', 'ollama'. Exiting."
        )
        sys.exit(1)

    output_base_dir_str = config.get("output_dir", "output")
    output_base_dir = pathlib.Path(output_base_dir_str)
    try:
        output_base_dir.mkdir(parents=True, exist_ok=True)
        logging.info(f"Using output directory: {output_base_dir.resolve()}")
    except Exception as e:
        logging.error(
            f"Failed to create output directory '{output_base_dir}': {e}. Exiting."
        )
        sys.exit(1)
    config["output_dir_resolved"] = output_base_dir

    base_cache_dir_from_config = api_settings.get("base_cache_dir", "api_cache")
    pathlib.Path(base_cache_dir_from_config).mkdir(parents=True, exist_ok=True)
    logging.info(f"Using base cache directory: {base_cache_dir_from_config}")

    generation_params = config.setdefault("generation_params", {})

    if not generation_params.get("main_topic"):
        logging.info(
            "No 'main_topic' found in config. Attempting to auto-generate one."
        )
        random_topic = generate_random_topic(config)
        if random_topic:
            generation_params["main_topic"] = random_topic
            logging.info(f"Auto-generated main_topic: '{random_topic}'")
        else:
            logging.critical(
                "Fatal: Failed to auto-generate random topic and none provided in config. Exiting."
            )
            sys.exit(1)
    else:
        logging.info(
            f"Using main_topic from config: '{generation_params['main_topic']}'"
        )

    main_topic = generation_params["main_topic"]
    sanitized_topic = sanitize_filename(main_topic, 64)
    topic_hash = hashlib.sha1(main_topic.encode("utf-8")).hexdigest()[:8]
    topic_dir_name = f"{sanitized_topic}_{topic_hash}"

    if api_provider == "gemini":
        gemini_conf = api_settings.get("gemini", {})
        model_name_for_cache = gemini_conf.get("model", "gemini-2.0-flash-latest")
    elif api_provider == "ollama":
        ollama_conf = api_settings.get("ollama", {})
        model_name_for_cache = ollama_conf.get("model", "ollama_default_model")
    else:
        model_name_for_cache = "unknown_api_provider_model"
    sanitized_model_name = sanitize_filename(model_name_for_cache, 30)
    logging.info(f"Using model name for cache path: '{sanitized_model_name}'")

    topic_specific_cache_dir = (
        pathlib.Path(base_cache_dir_from_config) / sanitized_model_name / topic_dir_name
    )
    logging.info(f"Topic-specific cache directory set to: {topic_specific_cache_dir}")
    config["cache_dir"] = str(
        topic_specific_cache_dir
    )
    topic_specific_cache_dir.mkdir(parents=True, exist_ok=True)

    equation_image_dir = topic_specific_cache_dir / "equation_images"
    equation_image_dir.mkdir(parents=True, exist_ok=True)
    logging.info(f"Equation image directory set to: {equation_image_dir}")

    if not generation_params.get("setting"):
        logging.info("No 'setting' found in config. Attempting to auto-generate one.")
        generated_setting = generate_setting(config)
        if generated_setting:
            generation_params["setting"] = generated_setting
            logging.info(f"Auto-generated setting: '{generated_setting[:100]}...'")
        else:
            logging.error("Failed to auto-generate setting. Using placeholder.")
            generation_params["setting"] = "[Setting Generation Failed]"
    else:
        logging.info(
            f"Using setting from config: '{generation_params['setting'][:100]}...'"
        )

    generation_params.setdefault(
        "setting", "[No Setting Provided or Generation Failed]"
    )
    logging.info(f"Final setting being used: '{generation_params['setting'][:100]}...'")

    author_name = generation_params.get("author_name", "").strip()
    author_gender = generation_params.get("author_gender", "").strip().lower()
    valid_genders = ["male", "female", "other"]

    if not author_name and not author_gender:
        logging.info("Author name and gender missing. Generating both...")
        generated_gender = generate_random_gender(config)
        logging.info(f"Randomly selected gender: {generated_gender}")
        generated_name = generate_random_name(config, generated_gender)
        if generated_name:
            author_name = generated_name
            author_gender = generated_gender
            generation_params["author_name"] = author_name
            generation_params["author_gender"] = author_gender
            logging.info(
                f"Generated Author: Name='{author_name}', Gender='{author_gender}'"
            )
        else:
            logging.critical(
                "Fatal: Failed to generate author name when both name and gender were missing. Exiting."
            )
            sys.exit(1)

    elif not author_name:
        logging.info(
            f"Author name missing. Generating name for specified gender: '{author_gender}'..."
        )
        if author_gender not in valid_genders:
            logging.error(
                f"Provided gender '{author_gender}' is not standard ({valid_genders}). Attempting name generation anyway."
            )

        generated_name = generate_random_name(config, author_gender)
        if generated_name:
            author_name = generated_name
            generation_params["author_name"] = author_name
            logging.info(
                f"Generated Author Name: '{author_name}' (Gender was '{author_gender}')"
            )
        else:
            logging.critical(
                f"Fatal: Failed to generate author name for gender '{author_gender}'. Exiting."
            )
            sys.exit(1)

    elif not author_gender:
        logging.info(
            f"Author gender missing. Attempting to determine gender from name: '{author_name}'..."
        )
        determined_gender = determine_gender_from_name(config, author_name)
        if determined_gender:
            author_gender = determined_gender
            generation_params["author_gender"] = author_gender
            logging.info(
                f"Determined Author Gender: '{author_gender}' (Name was '{author_name}')"
            )
        else:
            logging.error(
                f"Could not determine gender for '{author_name}'. Falling back to 'other'."
            )
            author_gender = "other"
            generation_params["author_gender"] = author_gender

    else:
        logging.info(f"Using Author Name from config: '{author_name}'")
        if author_gender not in valid_genders:
            logging.error(
                f"Author gender ('{author_gender}') from config is not standard ({valid_genders}). Using it anyway."
            )
        logging.info(f"Using Author Gender from config: '{author_gender}'")

    if not generation_params.get("author_name"):
        logging.critical(
            "Fatal: Author name is still missing after processing. Exiting."
        )
        sys.exit(1)
    if not generation_params.get("author_gender"):
        logging.critical(
            "Fatal: Author gender is still missing after processing. Exiting."
        )
        sys.exit(1)

    logging.info(
        f"Final Author Details: Name='{generation_params['author_name']}', Gender='{generation_params['author_gender']}'"
    )
    writing_tone = generation_params.get("writing_tone", "").strip()

    if not writing_tone:
        logging.info(
            "No 'writing_tone' found or it was empty in config. Attempting to auto-generate one."
        )
        generated_tone = generate_writing_tone(config)
        if generated_tone:
            writing_tone = generated_tone
            generation_params["writing_tone"] = (
                generated_tone
            )
            logging.info(f"Auto-generated writing tone: '{writing_tone}'")
        else:
            writing_tone = DEFAULT_WRITING_TONE
            generation_params["writing_tone"] = (
                writing_tone
            )
            logging.error(
                f"Failed to auto-generate writing tone. Using default: '{writing_tone}'"
            )
    else:
        logging.info(f"Using writing_tone from config: '{writing_tone}'")

    if (
        not writing_tone
    ):
        writing_tone = DEFAULT_WRITING_TONE
        generation_params["writing_tone"] = writing_tone
        logging.error(
            f"Writing tone was still empty after checks. Using default: '{writing_tone}'"
        )

    logging.info(f"Final writing tone being used: '{writing_tone}'")
    book_title_from_config = generation_params.get("book_title", "").strip()

    if book_title_from_config:
        book_title = book_title_from_config
        logging.info(f"Using Book Title from config: '{book_title}'")
    else:
        logging.info(
            "No 'book_title' found in config or it was empty. Attempting to auto-generate one."
        )
        book_title = generate_book_title(config)
        logging.info(f"Successfully auto-generated Book Title: {book_title}")
        generation_params["book_title"] = book_title

    is_book_non_fiction(
        config, book_title
    )

    if generation_params.get("is_fiction", False):
        logging.info(
            f"Book '{book_title}' has been identified as fiction. Starting fiction generation pipeline."
        )
        run_generation_process_fiction(config, output_base_dir, equation_image_dir)
    else:
        logging.info(
            f"Book '{book_title}' has been identified as non-fiction. Starting non-fiction generation pipeline."
        )
        run_generation_process_non_fiction(config, output_base_dir, equation_image_dir)

    end_time = time.time()
    logging.info(
        f"Book generation process finished in {end_time - start_time:.2f} seconds."
    )
