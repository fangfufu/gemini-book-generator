import logging
import re
import unicodedata
import uuid

import matplotlib.pyplot as plt


def sanitize_filename(filename_base, max_length=200):
    """Converts a string into a safe filename stem (without extension)."""
    filename_base = str(filename_base)  # Ensure it's a string
    filename_base = (
        unicodedata.normalize("NFKD", filename_base)
        .encode("ascii", "ignore")
        .decode("ascii")
    )
    # Replace whitespace and common invalid chars with underscore
    filename_base = re.sub(r"[\s/\\:?*\"<>|]+", "_", filename_base)
    # Remove any remaining non-alphanumeric characters (except underscore, hyphen, period)
    filename_base = re.sub(r"[^\w\-_\.]", "", filename_base)
    # Remove leading/trailing underscores, periods, hyphens
    filename_base = filename_base.strip("._-")
    # Limit length
    filename_base = filename_base[:max_length]
    # Handle empty string after sanitization
    if not filename_base:
        # Use a generic placeholder if sanitization results in empty string
        return "sanitized_empty"
    # Remove trailing periods (can cause issues on Windows)
    filename_base = filename_base.rstrip(".")
    # Ensure it's not empty again after stripping trailing period
    if not filename_base:
        return "sanitized_empty_final"
    return filename_base


def render_latex_to_image(
    latex_string_with_delimiters, equation_image_dir, is_display_style=False
):
    """
    Renders a LaTeX math string (potentially including delimiters like \(...\) or $$...$$)
    to a PNG image using matplotlib's mathtext.
    Adjusts padding based on whether it's display or inline style.
    """
    # 1. Clean the input string: Remove common LaTeX math delimiters robustly
    raw_latex_code = latex_string_with_delimiters.strip()
    original_input_for_log = raw_latex_code  # Keep original for logging

    # --- Improved Delimiter Removal ---
    # Repeatedly strip common delimiters until none are found at the ends
    cleaned = True
    while cleaned:
        cleaned = False
        original_length = len(raw_latex_code)
        if raw_latex_code.startswith(r"\(") and raw_latex_code.endswith(r"\)"):
            raw_latex_code = raw_latex_code[2:-2].strip()
            logging.debug("Stripped '\\(...\\)' delimiters.")
        elif raw_latex_code.startswith(r"\[") and raw_latex_code.endswith(r"\]"):
            raw_latex_code = raw_latex_code[2:-2].strip()
            logging.debug("Stripped '\\[...\\]' delimiters.")
        elif raw_latex_code.startswith("$$") and raw_latex_code.endswith("$$"):
            raw_latex_code = raw_latex_code[2:-2].strip()
            logging.debug("Stripped '$$...$$' delimiters.")
        elif (
            raw_latex_code.startswith("$")
            and raw_latex_code.endswith("$")
            and not raw_latex_code.startswith(
                "$$"
            )  # Avoid stripping single $ from $$..$$
        ):
            raw_latex_code = raw_latex_code[1:-1].strip()
            logging.debug("Stripped '$...$' delimiters.")

        if len(raw_latex_code) < original_length:
            cleaned = True  # Mark that we changed something and should loop again

    if not raw_latex_code:
        logging.error(
            f"LaTeX string became empty after removing delimiters. Original: '{original_input_for_log}'"
        )
        return None
    # --- End Improved Delimiter Removal ---

    # 2. Wrap the *cleaned* LaTeX code with $ for mathtext
    mathtext_string = f"${raw_latex_code}$"
    logging.debug(f"Attempting to render mathtext: {mathtext_string}")

    # Generate a unique filename
    filename = f"eq_{uuid.uuid4().hex}.png"
    # Use the passed-in directory
    filepath = equation_image_dir / filename

    fig = plt.figure()
    try:
        # Render the LaTeX using mathtext
        fontsize = 12 if not is_display_style else 14
        fig.text(
            0, 0, mathtext_string, fontsize=fontsize, va="bottom", math_fontfamily="cm"
        )

        # Adjust figure size to tightly fit the text
        renderer = fig.canvas.get_renderer()
        try:
            bbox = fig.texts[0].get_window_extent(renderer=renderer)
            if bbox.width <= 1 or bbox.height <= 1:
                raise ValueError(
                    f"Rendered bounding box is too small ({bbox.width}x{bbox.height}), likely mathtext render error."
                )
        except IndexError:
            raise ValueError(
                "Matplotlib fig.text failed to render (no text objects found)."
            )
        except ValueError as ve:
            raise ValueError(f"Matplotlib rendering error: {ve}")

        dpi = fig.get_dpi()
        # Calculate base width/height from bbox, add a minimal buffer to prevent clipping
        base_width = (bbox.width / dpi) + 0.01
        base_height = (bbox.height / dpi) + 0.01

        # --- Determine padding based on style ---
        if is_display_style:
            # Keep existing padding for display math
            figure_padding = 0.1
            save_pad_inches = 0.05
            logging.debug("Using display style padding.")
        else:
            # Use zero padding for inline math
            figure_padding = 0.0
            save_pad_inches = 0.0
            logging.debug("Using zero padding for inline style.")
        # --- End padding determination ---

        # Set final figure size including the calculated padding
        final_width = base_width + figure_padding
        final_height = base_height + figure_padding
        fig.set_size_inches(final_width, final_height)

        # Calculate text position within the final figure size
        # If figure_padding is 0, text starts near bottom-left (considering the small base buffer)
        text_x = (figure_padding / 2) / final_width if final_width > 0 else 0
        text_y = (figure_padding / 2) / final_height if final_height > 0 else 0
        fig.texts[0].set_position((text_x, text_y))

        # Make sure the directory exists (it should be created in main, but double-check)
        equation_image_dir.mkdir(parents=True, exist_ok=True)

        # Save the figure using the determined save padding
        plt.savefig(
            filepath,
            dpi=300,
            bbox_inches="tight",  # Still use tight bbox to help
            pad_inches=save_pad_inches,  # Use conditional padding here
            transparent=True,
        )
        plt.close(fig)
        logging.debug(f"Rendered LaTeX to {filepath}")
        return str(filepath)
    except Exception as e:
        # Log the original input (with delimiters) and the final mathtext string
        logging.error(
            f"Failed to render LaTeX input: '{original_input_for_log}'. Processed mathtext: '{mathtext_string}'. Error: {e}"
        )
        plt.close(fig)
        return None

def clean_text_for_docx(text):
    """
    Replaces internal newlines (and surrounding whitespace) with a single space.
    This is useful for ensuring that text content in DOCX paragraphs does not
    have unexpected line breaks.
    """
    if not text:
        return ""
    return re.sub(r"\s*\n\s*", " ", text.strip())
