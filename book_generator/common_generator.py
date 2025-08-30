import logging
import random
import re
import sys
import time
import uuid

from random_words import RandomWords

from book_generator.llm_api import call_llm_api
from book_generator.utils import sanitize_filename


# --- Book Generation Functions ---
def generate_random_gender(config):
    """Randomly returns either "male" or "female"."""
    return random.choice(["male", "female"])


def determine_gender_from_name(config, author_name):
    """Determines the likely gender (male/female) based on the author's name using the Gemini API."""
    logging.info(f"Attempting to determine gender for author name: '{author_name}'...")

    # Check if the name seems valid (basic check)
    if (
        not author_name
        or not isinstance(author_name, str)
        or " " not in author_name.strip()
    ):
        logging.error(
            f"Invalid or potentially incomplete name provided ('{author_name}'). Cannot reliably determine gender."
        )
        return None  # Indicate failure

    prompt = f"""
Based *only* on the full name '{author_name}', what is the most likely gender associated with the first name?

Consider common associations in Western cultures, primarily English-speaking contexts, as the book is in British English.

Output *only* one of the following words:
- male
- female
- other (use this only if the name is strongly ambiguous, unisex, or clearly not a typical given name)

Do not add any introductory text, explanations, or quotation marks. Just the single word.
"""

    gender_text = call_llm_api(
        prompt,
        config,
        cache_prefix=f"determine_gender_{sanitize_filename(author_name, 30)}",
    )

    if gender_text:
        cleaned_gender = gender_text.strip().lower()
        valid_genders = ["male", "female", "other"]
        if cleaned_gender in valid_genders:
            logging.info(
                f"Successfully determined likely gender for '{author_name}': '{cleaned_gender}'"
            )
            return cleaned_gender
        else:
            logging.error(
                f"API returned an unexpected value for gender ('{gender_text}'). Treating as undetermined."
            )
            return None  # Indicate failure or ambiguity
    else:
        logging.error(f"Failed to determine gender for '{author_name}' via API.")
        return None  # Indicate failure


def generate_random_name(config, gender):
    """Generates a random author name based on the specified gender using the Gemini API."""
    logging.info(f"Auto-generating random author name for gender: {gender}...")
    prompt = f"""Generate a single, plausible-sounding full name (first and last name)
for a fictional author. The author's gender is {gender}.
The author is notionally writing a book about
'{config['generation_params']['main_topic']}'. The setting of the book is
described as: {config['generation_params']['setting']}.
Consider a name that might appear on a book. Output *only* the full name.
Do not add introductory text, explanations, or quotation marks."""

    name_text = call_llm_api(prompt, config, cache_prefix="random_name")

    if name_text:
        cleaned_name = name_text.strip().strip("\"'").strip()
        if cleaned_name and " " in cleaned_name:
            logging.info(f"Successfully auto-generated random name: '{cleaned_name}'")
            return cleaned_name
        else:
            logging.error(
                f"Generated name ('{name_text}') might be invalid (e.g., missing space). Using it anyway."
            )
            return cleaned_name if cleaned_name else None
    else:
        logging.error("Failed to generate random name via API.")
        return None


def generate_random_topic(config):
    """
    Generates a book topic using the Gemini API.
    Uses a seed from config if provided, otherwise generates a random one.
    """
    logging.info("Determining topic seed...")
    if "generation_params" not in config:
        config["generation_params"] = {}

    # Check if a seed is provided in the config
    provided_seed = config["generation_params"].get("random_topic_seed", "").strip()

    if provided_seed:
        random_seed = provided_seed
        logging.info(f"Using random topic seed from config: '{random_seed}'")
    else:
        logging.info("No seed in config, generating a new random topic seed...")
        # Generate a new random seed if none was provided
        try:
            random_seed = " ".join(RandomWords().random_words(count=15))
            # Store the newly generated seed back into the config dictionary (in memory)
            config["generation_params"]["random_topic_seed"] = random_seed
            logging.info(f"Generated random topic seed: '{random_seed}'")
        except Exception as e:
            logging.error(f"Failed to generate random words for seed: {e}")
            # Fallback seed in case RandomWords fails
            random_seed = f"fallback_seed_{uuid.uuid4().hex[:8]}"
            config["generation_params"]["random_topic_seed"] = random_seed
            logging.error(f"Using fallback seed: '{random_seed}'")

    prompt = f"""Generate a topic for a book.
Random seed: {random_seed}
Output *only* the topic text itself.
Do not add introductory text, explanations, or quotation marks.
Output in British English."""

    logging.info(f"Generating random topic using seed: '{random_seed}'...")
    topic_text = call_llm_api(prompt, config, cache_prefix="random_topic")

    if topic_text:
        cleaned_topic = topic_text.strip().strip("\"'").rstrip(".").strip()
        if cleaned_topic:
            logging.info(f"Successfully auto-generated random topic: '{cleaned_topic}'")
            return cleaned_topic
        else:
            logging.error("Generated random topic was empty after cleaning.")
            return None
    else:
        logging.error("Failed to generate random topic via API.")
        return None


def generate_setting(config):
    """Generates the setting using the Gemini API based on the main topic."""
    logging.info("Auto-generating setting...")
    main_topic = config.get("generation_params", {}).get(
        "main_topic", "[No Main Topic Provided]"
    )

    if main_topic == "[No Main Topic Provided]":
        logging.error("Cannot generate setting without 'main_topic' in config.")
        return None

    prompt = f"""Based on the main topic '{main_topic}', generate an
short description of the setting where this topic could be explored.
Output only the setting description text. Do not add introductory text.
Output in British English."""

    setting_text = call_llm_api(prompt, config, cache_prefix="setting")

    if setting_text:
        cleaned_setting = setting_text.strip().strip("\"'")
        if cleaned_setting:
            logging.info(f"Successfully auto-generated setting:\n'{cleaned_setting}'")
            return cleaned_setting
        else:
            logging.error("Generated setting was empty after cleaning.")
            return None
    else:
        logging.error("Failed to generate setting via API.")
        return None


def generate_writing_tone(config):
    """Generates a suitable writing tone using the Gemini API."""
    logging.info("Auto-generating writing tone...")
    main_topic = config.get("generation_params", {}).get(
        "main_topic", "[No Main Topic Provided]"
    )
    setting = config.get("generation_params", {}).get(
        "setting", "[No  Setting Provided]"
    )

    if main_topic == "[No Main Topic Provided]" or setting == "[No  Setting Provided]":
        logging.error(
            "Cannot generate specific writing tone without 'main_topic' and 'setting'. Using a generic prompt."
        )
        prompt = f"""Generate some words describing a suitable writing tone
for a book. Output only the phrase describing the tone. Do not add introductory
text. Output in British English."""
    else:
        prompt = f"""Based on the main topic '{main_topic}', a setting described as:
"{setting}", generate some words describing the most suitable writing tone for a book exploring this
topic. Output *only* the phrase describing the tone. Do not add introductory
text. Output in British English."""
    tone_text = call_llm_api(prompt, config, cache_prefix="writing_tone")

    if tone_text:
        cleaned_tone = tone_text.strip().strip("\"'").rstrip(".")
        if cleaned_tone:
            logging.info(f"Successfully auto-generated writing tone: '{cleaned_tone}'")
            return cleaned_tone
        else:
            logging.error("Generated writing tone was empty after cleaning.")
            return None
    else:
        logging.error("Failed to generate writing tone via API.")
        return None


def generate_book_title(config):
    """Generates the book title using the Gemini API. Exits script on failure."""
    logging.info("Generating book title...")
    prompt = f"""Generate the book title for a book about
'{config['generation_params']['main_topic']}'. The setting of the book is
described as: {config['generation_params']['setting']}.
Do not generate a two-part title. The generated title must not contain a subtitle.
Provide only the title text. Do not add introductory text. The title must not
contain these punctuations: '-' or ':'. Output one title only. Output in
British English."""
    title = call_llm_api(prompt, config, cache_prefix="book_title")
    if title is None:
        logging.error("Fatal: Failed to generate book title after retries. Exiting.")
        sys.exit(1)
    cleaned_title = title.strip().strip("\"'")
    if not cleaned_title:
        logging.error("Fatal: Generated book title is empty. Exiting.")
        sys.exit(1)
    return cleaned_title


def generate_book_subtitle(config, book_title, summary_context):
    """Generates a book subtitle based on the title, topic, and chapter summaries."""
    logging.info(f"Generating subtitle for book: '{book_title}' (using summaries)...")
    gen_params = config.get("generation_params", {})

    # Check if subtitle is provided in config
    subtitle_from_config = gen_params.get("book_subtitle", "").strip()
    if subtitle_from_config:
        logging.info(f"Using book_subtitle from config: '{subtitle_from_config}'")
        return subtitle_from_config

    logging.info(
        "No 'book_subtitle' found in config or it was empty. Attempting to auto-generate one."
    )

    prompt = f"""Generate a subtitle for the book titled '{book_title}'.
The main topic of the book is '{config['generation_params']['main_topic']}'.
The setting of the book described as:
{config['generation_params']['setting']}.
{summary_context}
The subtitle should complement the main title.
Do not generate a two-part subtitle.
The subtitle must not contain these punctuations: '-' or ':'.
Provide only the subtitle text. Output one subtitle only.
Do not add introductory text. Output in British English."""

    subtitle = call_llm_api(prompt, config, cache_prefix="book_subtitle")
    if subtitle:
        cleaned_subtitle = subtitle.strip().strip("\"'")
        if cleaned_subtitle:
            logging.info(f"Successfully generated subtitle: {cleaned_subtitle}")
            return cleaned_subtitle
        else:
            logging.error(
                "Generated subtitle was empty after cleaning. No subtitle will be used."
            )
            return None
    else:
        logging.error("Failed to generate subtitle. No subtitle will be used.")
        return None


def is_book_non_fiction(config, book_title):
    """
    Determines if the book is likely non-fiction using the Gemini API.
    Caches the result in config['generation_params']['is_fiction'].
    Returns True if likely non-fiction, False otherwise (fiction or indeterminate).
    """
    gen_params = config.setdefault(
        "generation_params", {}
    )  # Ensure gen_params exists and can be modified

    # Check if 'is_fiction' is already determined and cached
    if "is_fiction" in gen_params and isinstance(gen_params["is_fiction"], bool):
        is_fiction_cached = gen_params["is_fiction"]
        logging.info(
            f"Using cached 'is_fiction' status for book '{book_title}': {is_fiction_cached}. "
            f"Returning non-fiction status: {not is_fiction_cached}."
        )
        return (
            not is_fiction_cached
        )  # True if non-fiction (is_fiction is False), False if fiction (is_fiction is True)

    # If not cached, proceed with API call
    main_topic = gen_params.get("main_topic", "[No Main Topic Provided]")
    setting = gen_params.get(
        "setting", "[No Setting Provided]"
    )  # Corrected variable name

    if main_topic == "[No Main Topic Provided]":
        logging.error(
            "Cannot determine book type for character list decision: 'main_topic' is missing. "
            "Assuming fiction to be safe and allow character generation if not explicitly disabled."
        )
        # Store this assumption and return
        gen_params["is_fiction"] = True  # Assume fiction
        return False  # is_book_non_fiction returns False for fiction

    prompt = f"""
Based on the following details of a book:
- Title: '{book_title}'
- Main Topic: '{main_topic}'
- Setting: "{setting}"

Is this book most likely non-fiction?
Answer with only 'yes' or 'no'. Do not add any explanations, quotation marks, or other text. Just the single word.
"""
    logging.info(
        f"Asking Gemini if book '{book_title}' is non-fiction for character list decision..."
    )

    # Create a cache prefix for this specific query
    safe_title_prefix = sanitize_filename(book_title, 30)
    cache_prefix = f"is_non_fiction_{safe_title_prefix}"

    response_text = call_llm_api(prompt, config, cache_prefix=cache_prefix)

    determined_is_non_fiction = False  # Default to fiction/indeterminate if API fails or gives unexpected response
    if response_text:
        answer = response_text.strip().lower()
        if answer == "yes":
            logging.info(f"Gemini indicates book '{book_title}' is likely non-fiction.")
            determined_is_non_fiction = True
        elif answer == "no":
            logging.info(
                f"Gemini indicates book '{book_title}' is likely fiction or its type is indeterminate."
            )
            determined_is_non_fiction = False
        else:
            logging.error(
                f"Unexpected response from Gemini for non-fiction check ('{response_text}'). "
                "Treating as fiction/indeterminate."
            )
            # determined_is_non_fiction remains False
    else:
        logging.error(
            f"Failed to get response from Gemini for non-fiction check for '{book_title}'. "
            "Treating as fiction/indeterminate."
        )
        # determined_is_non_fiction remains False

    # Cache the 'is_fiction' status
    # If determined_is_non_fiction is True (book is non-fiction), then gen_params['is_fiction'] should be False.
    # If determined_is_non_fiction is False (book is fiction/indeterminate), then gen_params['is_fiction'] should be True.
    gen_params["is_fiction"] = not determined_is_non_fiction
    logging.info(
        f"Stored 'is_fiction': {gen_params['is_fiction']} in generation_params for '{book_title}'."
    )

    return determined_is_non_fiction


def generate_character_list(config, book_title):
    """
    Generates a list of character names and descriptions based on book details,
    if enabled in the config or if the book is determined to be fiction.
    """
    gen_params = config.get("generation_params", {})

    # Check for explicit override in config
    config_override = gen_params.get(
        "generate_character_list"
    )  # Can be True, False, or None

    should_generate = False  # Default assumption is no generation

    if config_override is False:
        # Config explicitly says NOT to generate
        logging.info("Character list generation explicitly disabled by config.")
        should_generate = False
    elif config_override is True:
        # Config explicitly says to generate (override fiction check)
        logging.info(
            "Character list generation explicitly enabled by config (overriding fiction check)."
        )
        should_generate = True
    else:  # config_override is None (key not present)
        # Fall back to the fiction/non-fiction check
        logging.info(
            "Character list generation not explicitly set in config. Using fiction/non-fiction check."
        )
        should_generate = not is_book_non_fiction(config, book_title)

    # If we decided not to generate, set the config key to None and return
    if not should_generate:
        gen_params["character_list"] = None  # Ensure key exists
        return None

    logging.info("Attempting to generate character list...")

    main_topic = gen_params.get("main_topic", "[No Main Topic Provided]")
    setting = gen_params.get("setting", "[No Setting Provided]")

    # Critical prerequisites for character generation prompt
    if (
        main_topic == "[No Main Topic Provided]"
        or setting == "[No Setting Provided]"
        or not book_title  # book_title is a direct argument
    ):
        logging.error(
            "Cannot generate character list: 'main_topic', 'setting', or 'book_title' is missing/invalid. Skipping."
        )
        gen_params["character_list"] = None
        return None

    prompt = f"""
Based on the book titled '{book_title}', which has the main topic '{main_topic}',
a setting described as: "{setting}".

Generate a long list of characters who appear in this book.
For each character, provide their full name and a brief description
of their role, personality, or significance within the context of the topic and setting.
Do not use words that express uncertainty in the description. Do not use words such as "likely" or "potentially"

Format the output as a Markdown bulleted list. Each character should be an item.
Start the item with the character's name in bold, followed by a colon, and then the description.

Example:
*   **Character Name One:** A brief description of this character's role or significance.
*   **Another Character:** Their description and connection to the concepts.

Provide *only* the Markdown list of characters. Do not add introductory text like "Here is the character list:".
Output in British English.
"""

    character_list_text = call_llm_api(prompt, config, cache_prefix="character_list")

    if character_list_text:
        cleaned_text = character_list_text.strip()
        # Basic parsing: Split into lines and try to extract name/description
        characters = []
        for line in cleaned_text.split("\n"):
            line = line.strip()
            # Regex to capture bold name and the rest of the description
            match = re.match(r"^\*\s*\*\*(.*?)\*\*:\s*(.*)", line)
            if match:
                name = match.group(1).strip()
                description = match.group(2).strip()
                if name and description:
                    characters.append({"name": name, "description": description})
            elif line.startswith("* "):  # Handle cases where bolding might fail
                # Try a simpler split if regex fails but it looks like a list item
                parts = line[2:].split(":", 1)
                if len(parts) == 2 and parts[0].strip():
                    name = parts[0].strip()
                    description = parts[1].strip()
                    characters.append({"name": name, "description": description})

        if characters:
            logging.info(
                f"Successfully generated and parsed {len(characters)} characters."
            )
            # Store the parsed list in config
            gen_params["character_list"] = characters

            if characters:  # Check if the list is not empty
                character_log_details = "\n".join(
                    f"- {char.get('name', 'Unnamed')}: {char.get('description', 'No description')}"
                    for char in characters
                )
                logging.info(f"Generated Characters:\n{character_log_details}")

            return characters
        else:
            logging.error(
                f"Could not parse character list from API response. Response:\n{cleaned_text}"
            )
            gen_params["character_list"] = None
            return None
    else:
        logging.error("Failed to generate character list via API.")
        gen_params["character_list"] = None
        return None


# --- Helper function to format character list for prompts ---
def format_character_list_for_prompt(character_list):
    """Formats the character list into a string suitable for API prompts."""
    if not character_list or not isinstance(character_list, list):
        return ""  # Return empty string if no characters

    formatted_items = []
    for char in character_list:
        if isinstance(char, dict) and "name" in char and "description" in char:
            formatted_items.append(f"- {char['name']}: {char['description']}")
        # Add handling for other potential formats if needed

    if not formatted_items:
        return ""

    return (
        "Potential characters:\n"
        + "\n".join(formatted_items)
        + "\n"
        + "Not all potential characters have to be used."
    )










def generate_chapter_summary(
    config,
    chapter_title,
    writing_tone,
    previous_summaries_context="",
    character_context="",
):
    """
    Generates a brief summary for a given chapter title, considering previous summaries.

    Args:
        config (dict): The application configuration.
        chapter_title (str): The title of the chapter to summarize.
        writing_tone (str): The desired writing tone.
        previous_summaries_context (str, optional): A string containing summaries
                                                    of preceding chapters. Defaults to "".
    """
    logging.info(f"Generating summary for chapter: '{chapter_title}'...")

    # --- Build the prompt ---
    length_modifier = (
        config.get("generation_params", "").get("length_modifier", "").strip()
    )
    prompt_parts = [
        f"Write a one-paragraph {length_modifier} short and concise summary for the chapter titled '{chapter_title}'.",
        f"This chapter is part of a book about '{config['generation_params']['main_topic']}'.",
        f"The setting of the book is described as: {config['generation_params']['setting']}.",
        f"{character_context}",
    ]

    # Conditionally add context about previous chapters
    if previous_summaries_context:
        # Add the context clearly separated
        prompt_parts.append(
            f"\nContext: The summaries of the preceding chapters are:\n{previous_summaries_context}\n"
        )
        # Add instruction to avoid repetition
        prompt_parts.append(
            """Based on the preceding chapter summaries provided above, ensure
this new summary is distinct and logically follows or contrasts with them,
avoiding unnecessary repetition of themes or information already covered."""
        )

    # Prepare character-specific instruction
    character_consideration_text = ""
    if character_context:
        character_consideration_text = "Consider how the characters might be involved or relevant to this chapter's summary.\n"

    # Add remaining instructions
    prompt_parts.extend(
        [
            f"""\nMaintain a tone that is {writing_tone}.
{character_consideration_text}Output only the summary text for the current chapter ('{chapter_title}').
Do not add introductory text like 'This chapter summary is:'.
Output in British English."""
        ]
    )
    prompt = "\n".join(prompt_parts)
    logging.debug(
        f"Prompt for chapter summary '{chapter_title}':\n{prompt}"
    )  # Log the full prompt for debugging if needed
    # --- End prompt building ---

    summary_cache_prefix = f"summary_{sanitize_filename(chapter_title, max_length=40)}"
    summary = call_llm_api(prompt, config, cache_prefix=summary_cache_prefix)

    if summary:
        cleaned_summary = summary.strip()
        if cleaned_summary:
            logging.info(
                f"Successfully generated summary for '{chapter_title}':\n'{cleaned_summary}'"
            )
            return cleaned_summary
        else:
            logging.error(
                f"Generated summary for '{chapter_title}' was empty after cleaning."
            )
            # Return placeholder if summary is empty after cleaning
            return f"Placeholder summary for chapter '{chapter_title}' focusing on {config['generation_params']['main_topic']}."
    else:
        logging.error(
            f"Failed to generate summary for chapter '{chapter_title}'. Using placeholder."
        )
        return f"Placeholder summary for chapter '{chapter_title}' focusing on {config['generation_params']['main_topic']}."






def generate_front_matter(
    config, book_title, author_name, writing_tone, summary_context
):
    """Generates front matter elements, including the subtitle."""
    logging.info("Generating front matter...")
    front_matter = {}

    # Subtitle generation already calls call_llm_api with its own prefix
    book_subtitle = generate_book_subtitle(config, book_title, summary_context)
    config.get("generation_params", {})["book_subtitle"] = book_subtitle

    front_matter["title_page"] = {
        "title": book_title,
        "subtitle": book_subtitle,
        "author": author_name,
    }

    current_year = time.strftime("%Y")
    front_matter["copyright_page"] = (
        f"""
Copyright © {current_year} by {author_name}



All rights reserved.



No part of this publication may be reproduced, distributed, or transmitted in
any form or by any means, including photocopying, recording, or other electronic
or mechanical methods, without the prior written permission of the publisher,
except in the case of brief quotations embodied in critical reviews and certain
other non-commercial uses permitted by copyright law. For permission requests,
contact {author_name}.



The story, all names, characters, and incidents portrayed in this production
are fictitious. No identification with actual persons (living or deceased),
places, buildings, and products are intended or should be inferred.



This publication is designed to provide accurate and authoritative information
in regard to the subject matter covered. It is sold with the understanding that
neither the author nor the publisher is engaged in rendering legal, investment,
accounting or other professional services. While the publisher and author have
used their best efforts in preparing this book, they make no representations or
warranties with respect to the accuracy or completeness of the contents of this
book and specifically disclaim any implied warranties of merchantability or
fitness for a particular purpose. No warranty may be created or extended by
sales representatives or written sales materials. The advice and strategies
contained herein may not be suitable for your situation. You should consult with
a professional when appropriate. Neither the publisher nor the author shall be
liable for any loss of profit or any other commercial damages, including but not
limited to special, incidental, consequential, personal, or other damages.
""".strip()
    )

    common_prompt_base = f"""
for the book '{book_title}' about {config['generation_params']['main_topic']}, with the setting:
{config['generation_params']['setting']}. {summary_context}
Maintain a tone that is {writing_tone}. The author of this book is {author_name}.
Output *only* the text content for this section. Do not add introductory text.
Output in British English.
"""
    fm_elements_prompts = {
        "Dedication": f"Write an inspiring dedication {common_prompt_base}",  # Dedication is usually always present
        # Foreword generation is now conditional
        "Preface": f"Write a Preface {common_prompt_base}. {author_name} explains their motivation or the book's scope.",
        # Acknowledgements generation will also be conditional
    }

    # Conditionally add Foreword prompt
    gen_params = config.get("generation_params", {})
    should_generate_foreword = gen_params.get(
        "generate_foreword", True
    )  # Default to True if not specified
    if should_generate_foreword:
        fm_elements_prompts["Foreword"] = (
            f"Write a Foreword by a fictional expert relevant {common_prompt_base}. Make sure this fictional expert provides their name and credential at the end. Discuss the book's significance or context."
        )

    # Conditionally add Acknowledgements prompt
    should_generate_acknowledgements = gen_params.get(
        "generate_acknowledgements", True
    )  # Default to True
    if should_generate_acknowledgements:
        fm_elements_prompts["Acknowledgements"] = (
            f"Write an Acknowledgements {common_prompt_base}. {author_name} thanks individuals and groups who contributed."
        )

    for element, prompt in fm_elements_prompts.items():
        logging.info(f"Generating {element}...")
        # Use the element name (lowercase) as the prefix
        cache_prefix_str = element.lower()
        content = call_llm_api(prompt, config, cache_prefix=cache_prefix_str)

        processed_content = f"[{element} content generation failed.]"
        if content:
            content = content.strip()
            if content:
                expected_title = element
                lines = content.splitlines()
                processed_lines = list(lines)

                first_line_index = -1
                for i, line in enumerate(lines):
                    if line.strip():
                        first_line_index = i
                        break

                if first_line_index != -1:
                    first_line_content = lines[first_line_index].strip()
                    cleaned_first_line = re.sub(
                        r"^\s*#+\s*", "", first_line_content
                    ).strip()

                    if cleaned_first_line.lower() == expected_title.lower():
                        logging.debug(
                            f"Found and removing title '{lines[first_line_index]}' from {element} content."
                        )
                        del processed_lines[first_line_index]
                        while processed_lines and not processed_lines[0].strip():
                            logging.debug(
                                f"Removing blank line after title in {element}."
                            )
                            del processed_lines[0]
                        processed_content = "\n".join(processed_lines).strip()
                    else:
                        logging.debug(
                            f"First line of {element} ('{cleaned_first_line}') did not match expected title '{expected_title}'. Keeping original."
                        )
                        processed_content = content
                else:
                    processed_content = ""
            else:
                processed_content = ""
        else:
            logging.error(f"Failed to generate content for {element}.")
            # processed_content already holds the error message

        # Ensure Foreword key exists even if not generated, to avoid key errors later
        if element == "Foreword" and not should_generate_foreword:
            front_matter[element.lower()] = (
                None  # Or an empty string, depending on how you want to handle it
            )
            logging.info("Foreword generation skipped as per 'generate_foreword' flag.")
        # Ensure Acknowledgements key exists even if not generated
        elif element == "Acknowledgements" and not should_generate_acknowledgements:
            front_matter[element.lower()] = None
            logging.info(
                "Acknowledgements generation skipped as per 'generate_acknowledgements' flag."
            )

        front_matter[element.lower()] = processed_content

    return front_matter


def generate_appendix_subsection_titles(
    config, book_title, writing_tone, summary_context
):
    """Generates a list of potential subsection titles for the Appendix."""
    logging.info(f"Generating subsection titles for Appendix of book: '{book_title}'")
    main_topic = config["generation_params"]["main_topic"]
    setting = config["generation_params"]["setting"]

    prompt = f"""
Based on the book '{book_title}' about '{main_topic}',
with the setting: "{setting}",
and considering the following chapter summaries (if available):
{summary_context if summary_context.strip() and summary_context != "[No chapter summaries available]" else "No chapter summaries were provided for context."}

Generate a list of potential subsection titles for an Appendix section.
The appendix should contain supplementary material, detailed explanations,
data, or further readings relevant to the book's topic.

Format the output as a numbered list, with each title on a new line.
Example:
1. Subsection Title One
2. Another Relevant Subsection
3. Detailed Data Tables

Provide *only* the numbered list of subsection titles.
Do not add introductory text like "Here are the subsection titles:".
Output in British English.
"""
    cache_prefix = "appendix_subsection_titles"
    titles_text = call_llm_api(prompt, config, cache_prefix=cache_prefix)

    if titles_text:
        subsection_titles = []
        for line in titles_text.strip().split("\n"):
            line = line.strip()
            match = re.match(r"^\d+\.\s*(.*)", line)
            if match:
                title = match.group(1).strip()
                if title:
                    subsection_titles.append(title)
        if subsection_titles:
            logging.info(
                f"Successfully generated {len(subsection_titles)} appendix subsection titles."
            )
            return subsection_titles
        else:
            logging.error(
                f"Could not parse appendix subsection titles from API response: {titles_text}"
            )
            return []
    else:
        logging.error("Failed to generate appendix subsection titles via API.")
        return []


def generate_appendix_subsection_content(
    config,
    book_title,
    writing_tone,
    summary_context,
    subsection_title,
    all_subsection_titles,
):
    """Generates content for a specific appendix subsection."""
    logging.info(f"Generating content for Appendix subsection: '{subsection_title}'")
    main_topic = config["generation_params"]["main_topic"]
    setting = config["generation_params"]["setting"]

    prompt = f"""
Context for the book:
- Title: '{book_title}'
- Main Topic: '{main_topic}'
- Setting: "{setting}"
- Chapter Summaries (if available): {summary_context if summary_context.strip() and summary_context != "[No chapter summaries available]" else "No chapter summaries were provided for context."}
- Desired Writing Tone: {writing_tone}

Appendix Context:
- All planned subsection titles for this appendix: {', '.join(all_subsection_titles)}
- Current Subsection to write content for: '{subsection_title}'

Task:
Write the detailed content for the appendix subsection titled '{subsection_title}'.
Ensure the content is supplementary, informative, and directly relevant to this subsection's title and the overall book topic.
Avoid repeating content extensively covered in the main chapters (as suggested by the chapter summaries).

Instructions:
- Write substantial content for this subsection.
- Output *only* the text content for this subsection.
- Do *not* include the subsection title (e.g., "## {subsection_title}") in the output itself; start directly with the content.
- Format the output using standard Markdown (paragraphs, lists, bold, italics, tables).
- CRITICAL: Ensure all bulleted or numbered lists are preceded by a blank line in the Markdown output.
- Ensure paragraphs are separated by double line breaks in the Markdown source.
- If mathematical equations are necessary, format them using standard LaTeX syntax: use $...$ for inline math and $$...$$ for display math.
- Write the entire output in British English.
"""
    cache_prefix = f"appendix_content_{sanitize_filename(subsection_title, 60)}"
    content = call_llm_api(prompt, config, cache_prefix=cache_prefix)
    return content.strip() if content else None


def generate_back_matter(
    config, book_title, author_name, author_gender, writing_tone, summary_context
):
    """Generates back matter elements."""
    logging.info("Generating back matter...")
    # Get generation_params, defaulting to an empty dict if not present
    gen_params = config.get("generation_params", {})
    # Get the generate_appendix flag, defaulting to True if not specified
    should_generate_appendix = gen_params.get("generate_appendix", True)

    back_matter = {}
    common_prompt_base = f"""
of the book '{book_title}' about {config['generation_params']['main_topic']}, with the setting:
{config['generation_params']['setting']} {summary_context}
Maintain a tone that is {writing_tone}.
Do not add introductory text. Output *only* the text content for this section.
Output in British English."""

    # --- Handle Appendix Separately with Subsections ---
    if should_generate_appendix:
        logging.info("Attempting to generate Appendix content with subsections...")
        appendix_subsection_titles = generate_appendix_subsection_titles(
            config, book_title, writing_tone, summary_context
        )
        appendix_items = []  # Changed from appendix_content_parts

        if appendix_subsection_titles:
            logging.info(
                f"Generated {len(appendix_subsection_titles)} appendix subsection titles: {', '.join(appendix_subsection_titles)}"
            )
            for sub_title in appendix_subsection_titles:
                sub_content_md = (
                    generate_appendix_subsection_content(  # Renamed for clarity
                        config,
                        book_title,
                        writing_tone,
                        summary_context,
                        sub_title,
                        appendix_subsection_titles,
                    )
                )
                current_item = {"title": sub_title}
                if sub_content_md:
                    current_item["content"] = sub_content_md
                else:
                    logging.error(
                        f"Content generation failed for Appendix subsection: '{sub_title}'. Adding placeholder."
                    )
                    current_item["content"] = (
                        "[Content generation failed for this subsection.]"
                    )
                appendix_items.append(current_item)

            if appendix_items:  # If we have items (even with failed content)
                back_matter["appendix"] = appendix_items
            else:  # Should not happen if appendix_subsection_titles was non-empty, but as a safeguard
                logging.error(
                    "Appendix subsection titles were present, but no items were created. Appendix will be placeholder."
                )
                back_matter["appendix"] = (
                    "[Appendix generation failed: Could not process subsections.]"  # String placeholder
                )

        else:
            logging.error(
                "Failed to generate appendix subsection titles. Appendix will be minimal or placeholder."
            )
            back_matter["appendix"] = (
                "[Appendix generation failed: Could not determine subsections.]"  # String placeholder
            )
    else:
        logging.info(
            "Appendix generation skipped as per 'generate_appendix' flag in config."
        )
        # Set to None so it's skipped by assemble_docx's valid_bm_keys logic
        back_matter["appendix"] = None
    # --- End Appendix Handling ---

    bm_elements_prompts = {
        # Appendix is handled above
        "About the Author": f"Write an 'About the Author' section for {author_gender} author {author_name} {common_prompt_base}. Ensure the output is suitable for a book.",
    }
    for element, prompt in bm_elements_prompts.items():
        key = element.lower().replace(" ", "_")
        content = call_llm_api(prompt, config, cache_prefix=key)
        back_matter[key] = (
            content.strip() if content else f"[{element} content generation failed.]"
        )

    return back_matter


def generate_book_blurb(config, book_title, summary_context, writing_tone):
    """Generates a marketing blurb for the book."""
    logging.info(f"Generating marketing blurb for book: '{book_title}'...")

    main_topic = config["generation_params"]["main_topic"]
    setting = config["generation_params"]["setting"]

    prompt = f"""Write a compelling marketing blurb for a book titled '{book_title}'.
The main topic of the book is '{main_topic}'.
The setting of the book is described as: {setting}.
{summary_context}
The blurb should entice readers while accurately reflecting the book's content.
Maintain a tone that is {writing_tone}, but adapted for marketing purposes (e.g., more engaging, intriguing).
Output only the blurb text. Do not add introductory text.
Output in British English."""

    blurb = call_llm_api(prompt, config, cache_prefix="book_blurb")

    if blurb:
        cleaned_blurb = blurb.strip()
        if cleaned_blurb:
            logging.info("Successfully generated book blurb.")
            return cleaned_blurb
        else:
            logging.error(
                "Generated book blurb was empty after cleaning. Using placeholder."
            )
            return f"Placeholder blurb for book '{book_title}'."
    else:
        logging.error("Failed to generate book blurb. Using placeholder.")
        return f"Placeholder blurb for book '{book_title}'."


def generate_overall_summary(config, book_title, summary_context):
    """Generates a single overall book summary using chapter summaries."""
    logging.info(f"Generating overall book summary for: '{book_title}'...")

    if not summary_context or summary_context == "[No chapter summaries available]":
        logging.error(
            "Cannot generate overall summary: No chapter summaries available."
        )
        return f"Placeholder overall summary for the book '{book_title}'."

    prompt = f"""Based *only* on the following chapter summaries for the book
titled '{book_title}', write a short overall summary or abstract of
the entire book.


Chapter Summaries:
{summary_context}

Output *only* the overall summary text. Do not add introductory text.
Output in British English."""

    overall_summary = call_llm_api(prompt, config, cache_prefix="overall_book_summary")

    if overall_summary:
        cleaned_summary = overall_summary.strip()
        if cleaned_summary:
            logging.info("Successfully generated overall book summary.")
            return cleaned_summary
        else:
            logging.error(
                "Generated overall summary was empty after cleaning. Using placeholder."
            )
            return f"Placeholder overall summary for the book '{book_title}'."
    else:
        logging.error(
            "Failed to generate overall book summary via API. Using placeholder."
        )
        return f"Placeholder overall summary for the book '{book_title}'."


def save_summary_to_markdown(book_title, overall_summary, output_dir):
    """Saves the overall book summary to a Markdown file."""
    if not overall_summary:
        logging.error("No overall summary provided to save.")
        return

    filename_stem = sanitize_filename(book_title)
    if not filename_stem or filename_stem.startswith("sanitized_empty"):
        logging.error(
            f"Could not create a valid filename from title '{book_title}'. Skipping summary Markdown save."
        )
        return

    output_filename = output_dir / f"{filename_stem}_Summary.md"
    markdown_content = f"# {book_title}\n\n{overall_summary}\n"

    try:
        with open(output_filename, "w", encoding="utf-8") as f:
            f.write(markdown_content)
        logging.info(f"Overall book summary saved to Markdown: '{output_filename}'")
    except Exception as e:
        logging.error(
            f"Error saving overall summary to Markdown file '{output_filename}': {e}"
        )
