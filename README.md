# Gemini Book Generator

  A Python script that leverages the Google Gemini API to automatically generate complete books, including front matter, body content, back matter, and marketing materials, based on a configuration file. The script formats the output into DOCX files with support for LaTeX math rendering.

  ## Features

  *   **AI-Powered Content Generation:** Uses the Google Gemini API to generate:
      *   Book Title & Subtitle
      *   Chapter Outlines & Summaries
      *   Section Titles & Detailed Content (Markdown format)
      *   Front Matter (Dedication, Foreword, Preface, Acknowledgements)
      *   Back Matter (Appendix, Glossary, Bibliography, About the Author)
      *   Marketing Blurb
  *   **Configurable Generation:** Control various aspects via `config.yaml`:
      *   Main Topic, Universe Setting, Key Concepts
      *   Author Details (Name, Gender) - Can be auto-generated
      *   Writing Tone - Can be auto-generated
      *   Gemini Model & API Parameters (Temperature, Retries)
      *   DOCX Styling (Font, Size, Page Size, Margins)
  *   **Automatic Topic/Author Generation:** Option to automatically generate a random topic, universe setting, key concepts, author name/gender, and writing tone if not specified.
  *   **DOCX Output:** Assembles the generated content into a structured `.docx` file using `python-docx`.
  *   **Complex DOCX Formatting:**
      *   Handles Markdown conversion (including tables, lists, emphasis).
      *   Applies custom styles (Normal, Headings, Code Blocks, Lists).
      *   Configurable page size (6x9, A4) and mirrored margins with gutter.
      *   Section breaks with distinct page numbering (Roman numerals for front matter, decimals for body/back matter).
  *   **LaTeX Math Support:**
      *   Parses LaTeX math delimited by `$...$` (inline) and `$$...$$` (display) using `pymdownx.arithmatex`.
      *   Renders LaTeX equations to PNG images using Matplotlib.
      *   Inserts math images into the DOCX, attempting to scale appropriately.
  *   **API Caching:** Caches Gemini API responses locally to save costs and speed up subsequent runs with the same prompts. Cache files are organized into topic-specific directories.
  *   **Marketing Document:** Generates a separate `_Marketing.docx` file containing the book blurb, author bio, and key details.
  *   **Error Handling & Logging:** Includes retries for API calls and detailed logging throughout the process.
  *   **Debugging:** Option to save intermediate HTML generated from Markdown.

  ## Requirements

  *   Python 3.x
  *   Google Gemini API Key
  *   Required Python packages (install via `requirements.txt`):
      *   `google-generativeai`
      *   `python-docx`
      *   `PyYAML`
      *   `python-dotenv`
      *   `matplotlib`
      *   `lxml`
      *   `Pillow`
      *   `random_words`
      *   `markdown`
      *   `pymdown-extensions`

  ## Setup

  1.  **Clone the repository:**
      ```bash
      git clone <your-repo-url>
      cd <your-repo-directory>
      ```

  2.  **Create a virtual environment (recommended):**
      ```bash
      python -m venv venv
      source venv/bin/activate  # On Windows use `venv\Scripts\activate`
      ```

  3.  **Install dependencies:**
      ```bash
      pip install -r requirements.txt
      ```
      *(Note: You may need to install LaTeX system-wide if you don't have it, for Matplotlib's `usetex` feature).*

  4.  **Set up API Key:**
      *   Create a file named `.env` in the project root.
      *   Add your Google Gemini API key to the `.env` file:
          ```env
          GEMINI_API_KEY='YOUR_API_KEY_HERE'
          ```

  ## Configuration

  *   The core behaviour of the script is controlled by the `config.yaml` file.
  *   This file defines generation parameters (topic, author, tone, etc.), API settings (model, temperature), styling options (fonts, margins), and debugging flags.
  *   You need to create and populate `config.yaml` according to the script's expected structure. Refer to the `load_config` function and parameter usage within `generate_book.py` for details on required and optional keys.

  ## Usage

  1.  Ensure your `config.yaml` file is correctly configured.
  2.  Make sure your `.env` file contains the `GEMINI_API_KEY`.
  3.  Run the script from the project's root directory:
      ```bash
      python generate_book.py
      ```
  4.  The script will log its progress to the console.

  ## Output

  The script generates the following outputs in the project directory (or as configured):

  *   **`<sanitized_book_title>.docx`:** The main generated book file.
  *   **`<sanitized_book_title>_Marketing.docx`:** A separate document with marketing information (blurb, author bio, etc.).
  *   **`api_cache/` (or configured cache directory):**
      *   Contains topic-specific subdirectories (e.g., `My_Awesome_Topic_abc123def/`).
      *   Each topic directory holds:
          *   `.json` files caching API prompts and responses.
          *   `equation_images/` subdirectory containing rendered LaTeX math PNGs.
          *   `debug_html/` subdirectory (if enabled) containing intermediate HTML files.

  ## Caching

  *   API calls are cached based on a hash of the prompt text.
  *   Cache files are stored in a base directory (default: `api_cache`), further organized into subdirectories named after the sanitized main topic and a short hash. This keeps caches for different book projects separate.
  *   Cache filenames include an optional, sanitized prefix (e.g., `chapter_outline`, `section_content_...`) for easier identification.
  *   To force regeneration, delete the relevant `.json` files from the topic-specific cache directory or the entire directory itself.

  ## LaTeX / Math Support

  *   The script uses Matplotlib with `text.usetex = True` to render LaTeX math. Ensure you have a working LaTeX distribution installed for this feature.
  *   Markdown content generated by the AI should use standard LaTeX delimiters:
      *   `$...$` for inline math.
      *   `$$...$$` for display math.
  *   These are parsed by `pymdownx.arithmatex` during Markdown conversion.
  *   The `render_latex_to_image` function extracts the LaTeX code, renders it to a PNG image using Matplotlib, and saves it in the `equation_images` cache subdirectory.
  *   These images are then embedded into the DOCX file. Inline math height is adjusted based on font size, and display math can be scaled to fit page width.

  ## Debugging

  *   Set `save_intermediate_html: true` under `debug_options` in `config.yaml` to save the HTML generated from the Markdown content of each section before it's processed into the DOCX. These files are saved in the `debug_html` subdirectory within the topic-specific cache folder.

## License

    Gemini Book Generator - Automatically generate books
    Copyright (C) <2025>  Fufu Fang

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
