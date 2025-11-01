"""
This file contains the functions for generating non-fiction content.
"""
import logging
import re
import sys

from book_generator.llm_api import call_llm_api
from book_generator.utils import sanitize_filename


def generate_chapter_outline(config, character_context=""):
    """Generates a list of chapter titles."""
    logging.info("Generating chapter outline...")
    is_fiction = config.get("generation_params", {}).get("is_fiction", False)
    if is_fiction:
        logging.info("Book is fiction, skipping non-fiction chapter outline generation.")
        return []

    length_modifier = (
        config.get("generation_params", "").get("length_modifier", "").strip()
    )
    is_fiction = config.get("generation_params", {}).get("is_fiction", False)

    if is_fiction:
        list_description_for_prompt = f"list"
    else:
        list_description_for_prompt = f"{length_modifier} short list"

    prompt = f"""Generate a {list_description_for_prompt} of chapter titles for a
book about '{config['generation_params']['main_topic']}'. The setting of the book
is described as: {config['generation_params']['setting']}.
{character_context}
The chapters should logically progress through the topic, potentially involving the characters.
Ensure the list of chapter titles is appropriate for the type of book (fiction/non-fiction).
Format the output as a numbered list, with each title on a new line.
Start numbering from 1. Example:
1. Chapter Title One
2. Chapter Title Two
Do not use font formatting (e.g. bold, italics and etc) in the chapter title.
Do not generate two-part titles. The generated titles must not contain subtitles.
The chapter titles must not contain these punctuations: '-' or ':'.
Do not add introductory text.
Output in British English."""
    outline_text = call_llm_api(prompt, config, cache_prefix="chapter_outline")
    if outline_text:
        chapter_titles = []
        # Improved parsing to handle potential variations
        for line in outline_text.strip().split("\n"):
            line = line.strip()
            match = re.match(r"^\d+\.\s*(.*)", line)
            if match:
                title = match.group(1).strip()
                if title:  # Ensure title is not empty
                    chapter_titles.append(title)
        return chapter_titles
    else:
        logging.error("Failed to generate chapter outline via API. Using fallback.")
        return [
            f"Chapter {i+1}: Placeholder Title"
            for i in range(config["generation_params"]["num_chapters_fallback"])
        ]


def generate_section_titles(
    config,
    chapter_title,
    chapter_summary,
    all_chapter_titles,
    all_chapter_summaries,
    character_context="",
):
    """
    Generates a list of section titles for a given chapter, using its summary
    and considering the context of other chapters to avoid repetition.
    """
    logging.info(
        f"Generating section titles for chapter: '{chapter_title}' (using summary and full book context)..."
    )

    # --- Prepare Context for Prompt ---
    # Format all chapter titles
    all_titles_context = "\n".join(
        f"- {idx+1}. {title}" for idx, title in enumerate(all_chapter_titles)
    )

    # Format all chapter summaries (excluding the current one for brevity,
    # as it's provided separately)
    all_summaries_context_parts = []
    for idx, title in enumerate(all_chapter_titles):
        if title != chapter_title:  # Exclude current chapter's summary from this list
            summary = all_chapter_summaries.get(title, "[Summary not available]")
            all_summaries_context_parts.append(f"- Chapter '{title}': {summary}")
    all_summaries_context = "\n".join(all_summaries_context_parts)
    # --- End Context Preparation ---

    length_modifier = (
        config.get("generation_params", {}).get("length_modifier", "").strip()
    )

    prompt = f"""
Context for the entire book:
Main Topic: '{config['generation_params']['main_topic']}'
Setting: {config['generation_params']['setting']}
{character_context}

Full Chapter Outline:
{all_titles_context}

Summaries of OTHER chapters (for context on what's covered elsewhere):
{all_summaries_context}

---
Task:
Generate a {length_modifier} short list of relevant section titles specifically for the
chapter titled '{chapter_title}'.This chapter's specific summary is:
"{chapter_summary}"

Instructions:
- The section titles should logically break down the chapter's topic as
described in *its specific summary*.
{"- Consider how the characters might relate to these sections." if character_context else ""}
- Ensure the generated section titles are distinct and avoid significant
overlap with topics clearly covered in the *summaries of other chapters*
provided above or topics strongly implied by the *titles of other chapters*.
- Format the output as a numbered list, with each title on a new line
(e.g., 1. Section Title One).
- Do not generate two-part section titles. The generated section titles must
not contain subtitles.
- The section titles must not contain these punctuations: '-' or ':'.
- Do not use font formatting (e.g., bold, italics) in the section titles.
- Output *only* the numbered list of section titles. Do not add introductory text.
- Output in British English.
"""

    section_titles_cache_prefix = (
        f"section_titles_{sanitize_filename(chapter_title, max_length=40)}"
    )
    titles_text = call_llm_api(prompt, config, cache_prefix=section_titles_cache_prefix)
    num_chapter_fallback = config["generation_params"]["num_chapter_fallback"]
    if titles_text:
        section_titles = []
        # Improved parsing
        for line in titles_text.strip().split("\n"):
            line = line.strip()
            match = re.match(r"^\d+\.\s*(.*)", line)
            if match:
                title = match.group(1).strip()
                if title:
                    section_titles.append(title)

        # Use generated titles if we got *any*, otherwise fallback
        if section_titles:
            logging.info(
                f"Successfully generated {len(section_titles)} section titles for '{chapter_title}' considering context."
            )
            return section_titles
        else:
            logging.error(
                f"Could not parse section titles for '{chapter_title}' from API response or response was empty. Using placeholders. Response:\n{titles_text}"
            )
            return [
                f"Section {i+1}: Placeholder Title" for i in range(num_chapter_fallback)
            ]
    else:
        logging.error(
            f"Failed to generate section titles for '{chapter_title}' via API. Using placeholders."
        )
        return [
            f"Section {i+1}: Placeholder Title" for i in range(num_chapter_fallback)
        ]


def generate_section_content(
    config,
    chapter_title,
    section_title,
    section_num,
    total_sections,
    chapter_summary,
    writing_tone,
    character_context="",
):
    """Generates content for a single section using Markdown, asking AI to use
    LaTeX math and avoid sub-headings."""
    logging.info(
        f"Generating content for: Chapter '{chapter_title}' -> Section {section_num}/{total_sections}: '{section_title}' (using summary)"
    )

    prompt = f"""
Context:
- Book Main Topic: '{config['generation_params']['main_topic']}'
- Setting: {config['generation_params']['setting']}
{"-" if character_context else ""}{character_context}

---
Current Task Context:
- Current Chapter Title: '{chapter_title}'
- Current Chapter Summary: "{chapter_summary}"
- Current Section Title: '{section_title}'
- Current Section Number: {section_num} of {total_sections}
- Desired Writing Tone: {writing_tone}

Task:
Write a detailed section for the book described above, focusing specifically on
the topic defined by the section title ('{section_title}'). Ensure the content
fits logically within the context provided by the current chapter summary.

Instructions:
- Output *only* the text content for this section.
- Do *not* include the main chapter title or the section title in the output
itself. Start directly with the section's content.
- Format the output using standard Markdown (paragraphs, lists, bold, italics,
tables).
- CRITICAL: Ensure all bulleted or numbered lists are preceded by a blank line in the
Markdown output.
- Ensure paragraphs are separated by double line breaks in the Markdown source.
- Do *not* include any Markdown sub-headings (like ## Heading Level 2 or
### Heading Level 3).
- If mathematical equations are necessary, format them using standard LaTeX
syntax: use $...$ for inline math and $$...$$ for display math.
- Write the entire output in British English.
"""
    # --- Create the cache prefix ---
    # Sanitize chapter and section titles and combine them, limit length
    safe_chapter_title = sanitize_filename(chapter_title, max_length=30)
    safe_section_title = sanitize_filename(section_title, max_length=30)
    # Use a clear separator like '__' which is less common in titles
    cache_prefix_str = f"content_{safe_chapter_title}__{safe_section_title}"
    # --- End cache prefix creation ---

    # --- Call API with the prefix ---
    content = call_llm_api(prompt, config, cache_prefix=cache_prefix_str)
    # --- End API call ---

    return (
        content
        if content
        else f"**Content generation failed for Chapter '{chapter_title}', Section '{section_title}'.**"
    )
