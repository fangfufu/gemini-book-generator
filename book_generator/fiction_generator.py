"""
This file contains the functions for generating fiction content.
"""
import logging
import re
import sys

from book_generator.llm_api import call_llm_api
from book_generator.utils import sanitize_filename


def generate_overall_story(config, writing_tone=""):
    """Generates the overall story for a fiction book."""
    logging.info("Generating overall story...")
    main_topic = config.get("generation_params", {}).get("main_topic", "[No Main Topic Provided]")
    setting = config.get("generation_params", {}).get("setting", "[No Setting Provided]")

    prompt = f"""
Based on the main topic '{main_topic}', a setting described as:
"{setting}".

Write a detailed overall story of about 2000 words.
This story will serve as the master plot for the entire book.
Do not include chapter headings in the output.
The writing tone should be: {writing_tone}.
Output only the story text. Do not add introductory text.
Output in British English.
"""
    story_text = call_llm_api(prompt, config, cache_prefix="overall_story")

    if story_text:
        cleaned_story = story_text.strip()
        if cleaned_story:
            logging.info("Successfully generated overall story.")
            return cleaned_story
        else:
            logging.error("Generated overall story was empty after cleaning.")
            return None
    else:
        logging.error("Failed to generate overall story via API.")
        return None


def generate_fiction_chapter_outline(config, overall_story, character_context="", location_context="", fiction_chapter_count=20, writing_tone=""):
    """Generates a list of chapter titles and summaries for a fiction book."""
    logging.info("Generating fiction chapter outline (titles and summaries)...")
    prompt = f"""
Based on the following overall story:
--- STORY START ---
{overall_story}
--- STORY END ---

And the following characters:
{character_context}

And the following locations:
{location_context}

Break down the story into {fiction_chapter_count} chapters. For each chapter, provide a title and a one-paragraph summary.
The chapters should logically follow the progression of the story.
The writing tone for the summaries should be: {writing_tone}.

Format the output as a numbered list of chapters. For each chapter, provide the title and then the summary.
Example:
1.  **Chapter Title One:** A one-paragraph summary of the first chapter.
2.  **Chapter Title Two:** A one-paragraph summary of the second chapter.

Output only the list of chapters. Do not add any introductory text.
Output in British English.
"""
    outline_text = call_llm_api(prompt, config, cache_prefix="fiction_chapter_outline")

    if outline_text:
        chapters = []
        # Regex to capture chapter number, title, and summary
        pattern = re.compile(r"^\d+\.\s+\*\*(.*?):\*\*\s*(.*)", re.MULTILINE)
        matches = pattern.finditer(outline_text)
        for match in matches:
            title = match.group(1).strip()
            summary = match.group(2).strip()
            if title and summary:
                chapters.append({"title": title, "summary": summary})

        if chapters:
            logging.info(f"Successfully generated {len(chapters)} chapter outlines.")
            return chapters
        else:
            logging.error(f"Could not parse chapter outline from API response: {outline_text}")
            return None
    else:
        logging.error("Failed to generate fiction chapter outline via API.")
        return None


def generate_fiction_chapter_content(
    config,
    previous_chapters_summary,
    chapter_title,
    chapter_summary,
    character_context="",
    location_context="",
    writing_tone="",
):
    """Generates the content for a single fiction chapter."""
    logging.info(f"Generating content for fiction chapter: '{chapter_title}'...")

    prompt = f"""
You are a fiction writer. Your task is to write a single chapter of a book.

Here is the context for your writing:

--- Summary of Previous Chapters ---
{previous_chapters_summary}
--- End Summary of Previous Chapters ---

--- Current Chapter ---
Title: {chapter_title}
Summary: {chapter_summary}
--- End Current Chapter ---

{character_context}

{location_context}

Writing Tone: {writing_tone}

Based on all the context provided, write the full content of the current chapter.

Instructions:
- The chapter must logically and narratively continue from the previous chapter.
- Write about 2000 words.
- Avoid repeating the structure and content of the previous chapter.
- Output *only* the text content for this chapter.
- Do *not* include the main chapter title in the output itself. Start directly with the chapter's content.
- Format the output using standard Markdown (paragraphs, lists, bold, italics).
- Write the entire output in British English.
"""
    cache_prefix = f"fiction_content_{sanitize_filename(chapter_title, max_length=40)}"
    content = call_llm_api(prompt, config, cache_prefix=cache_prefix)

    return (
        content
        if content
        else f"**Content generation failed for Chapter '{chapter_title}'.**"
    )


def generate_location_list(config, overall_story):
    """
    Generates a list of physical locations based on the overall story.
    """
    logging.info("Attempting to generate location list...")

    if not overall_story:
        logging.error("Cannot generate location list: 'overall_story' is missing. Skipping.")
        return None

    prompt = f"""
Based on the following overall story:
--- STORY START ---
{overall_story}
--- STORY END ---

Generate a list of physical locations that appear in this story.
For each location, provide its name and a brief description of its significance within the story.

Format the output as a Markdown bulleted list. Each location should be an item.
Start the item with the location's name in bold, followed by a colon, and then the description.

Example:
*   **Location Name One:** A brief description of this location's role or significance.
*   **Another Location:** Its description and connection to the story.

Provide *only* the Markdown list of physical locations. Do not add introductory text.
Output in British English.
"""

    location_list_text = call_llm_api(prompt, config, cache_prefix="location_list")

    if location_list_text:
        cleaned_text = location_list_text.strip()
        locations = []
        for line in cleaned_text.split("\n"):
            line = line.strip()
            match = re.match(r"^\*\s*\*\*(.*?)\*\*:\s*(.*)", line)
            if match:
                name = match.group(1).strip()
                description = match.group(2).strip()
                if name and description:
                    locations.append({"name": name, "description": description})
            elif line.startswith("* "):
                parts = line[2:].split(":", 1)
                if len(parts) == 2 and parts[0].strip():
                    name = parts[0].strip()
                    description = parts[1].strip()
                    locations.append({"name": name, "description": description})

        if locations:
            logging.info(
                f"Successfully generated and parsed {len(locations)} locations."
            )
            return locations
        else:
            logging.error(
                f"Could not parse location list from API response. Response:\n{cleaned_text}"
            )
            return None
    else:
        logging.error("Failed to generate location list via API.")
        return None


def update_location_list(config, location_list, chapter_content):
    """
    Updates the physical location list based on the content of the latest chapter.
    """
    logging.info("Updating location list...")

    location_context = ""
    if location_list:
        location_context = "\n".join(
            [f"- {loc['name']}: {loc['description']}" for loc in location_list]
        )

    prompt = f"""
Given the following existing list of locations:
{location_context}

And the following chapter content:
--- CHAPTER CONTENT START ---
{chapter_content}
--- CHAPTER CONTENT END ---

Update the physical location list based on the chapter content.
- If a new physical location is introduced, add it to the list with a description.
- If an existing physical location's description needs to be updated, modify it.
- If a physical location is not mentioned, keep it in the list as is.

Format the output as a Markdown bulleted list. Each location should be an item.
Start the item with the location's name in bold, followed by a colon, and then the description.

Example:
*   **Location Name One:** An updated or new description.
*   **New Location:** A description of this newly introduced location.

Provide *only* the Markdown list of locations. Do not add introductory text.
Output in British English.
"""

    updated_location_list_text = call_llm_api(
        prompt, config, cache_prefix="update_location_list"
    )

    if updated_location_list_text:
        cleaned_text = updated_location_list_text.strip()
        locations = []
        for line in cleaned_text.split("\n"):
            line = line.strip()
            match = re.match(r"^\*\s*\*\*(.*?)\*\*:\s*(.*)", line)
            if match:
                name = match.group(1).strip()
                description = match.group(2).strip()
                if name and description:
                    locations.append({"name": name, "description": description})
            elif line.startswith("* "):
                parts = line[2:].split(":", 1)
                if len(parts) == 2 and parts[0].strip():
                    name = parts[0].strip()
                    description = parts[1].strip()
                    locations.append({"name": name, "description": description})

        if locations:
            logging.info(
                f"Successfully updated and parsed {len(locations)} locations."
            )
            return locations
        else:
            logging.error(
                f"Could not parse updated location list from API response. Response:\n{cleaned_text}"
            )
            return location_list
    else:
        logging.error("Failed to update location list via API.")
        return location_list


def summarize_fiction_chapter(config, chapter_title, chapter_content, writing_tone=""):
    """Summarizes the content of a fiction chapter."""
    logging.info(f"Summarizing content for chapter '{chapter_title}'...")

    prompt = f"""
The following is the full text of a chapter titled '{chapter_title}'.
--- CHAPTER CONTENT START ---
{chapter_content}
--- CHAPTER CONTENT END ---

Your task is to summarize this chapter in a short paragraph.
The summary should capture the key events, character developments, and plot advancements.

Output only the summary paragraph. Do not add any introductory text.
Output in British English.
"""
    cache_prefix = (
        f"fiction_summary_{sanitize_filename(chapter_title, max_length=40)}"
    )
    summary = call_llm_api(prompt, config, cache_prefix=cache_prefix)

    if summary:
        cleaned_summary = summary.strip()
        if cleaned_summary:
            logging.info(
                f"Successfully generated summary for chapter '{chapter_title}'."
            )
            return cleaned_summary
    logging.error(f"Failed to generate summary for chapter '{chapter_title}'.")
    return None
