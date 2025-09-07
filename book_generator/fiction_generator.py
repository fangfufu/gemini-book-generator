"""
This file contains the functions for generating fiction content.
"""
import logging
import re
import sys

from book_generator.llm_api import call_llm_api
from book_generator.utils import sanitize_filename


def generate_overall_story(config, character_context="", writing_tone=""):
    """Generates the overall story for a fiction book."""
    logging.info("Generating overall story...")
    main_topic = config.get("generation_params", {}).get("main_topic", "[No Main Topic Provided]")
    setting = config.get("generation_params", {}).get("setting", "[No Setting Provided]")

    prompt = f"""
Based on the main topic '{main_topic}', a setting described as:
"{setting}", and the following characters:
{character_context}

Write a  detailed overall story of about 2000 words. This story will serve as the master plot for the entire book.
The story should have a clear beginning, middle, and end.
It should introduce the main conflict, develop the plot, and provide a resolution.
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


def generate_fiction_chapter_outline(config, overall_story, character_context="", fiction_chapter_count=20, writing_tone=""):
    """Generates a list of chapter titles and summaries for a fiction book."""
    logging.info("Generating fiction chapter outline (titles and summaries)...")
    prompt = f"""
Based on the following overall story:
--- STORY START ---
{overall_story}
--- STORY END ---

And the following characters:
{character_context}

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
    previous_chapter_title,
    previous_chapter_summary,
    chapter_title,
    chapter_summary,
    character_context="",
    writing_tone="",
):
    """Generates the content for a single fiction chapter."""
    logging.info(f"Generating content for fiction chapter: '{chapter_title}'...")

    prompt = f"""
Context:
- Characters: {character_context}
- Previous Chapter Title: '{previous_chapter_title}'
- Previous Chapter Summary: '{previous_chapter_summary}'
- Current Chapter Title: '{chapter_title}'
- Current Chapter Summary: '{chapter_summary}'
- Writing Tone: {writing_tone}

Task:
Write a detailed chapter for the book, based on its summary, focusing on the story elements relevant to the current chapter title.
The chapter should be approximately 2000 words long.
Ensure the chapter is distinct from the previous chapter.
Adhere strictly to the specified writing tone.
Avoid repetition.

Instructions:
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
