import sys
import argparse
import logging
import matplotlib
import pathlib
from book_generator.config import load_config
from book_generator.orchestrator import run_generation_process

def run_cli():
    """Runs the command-line book generation process."""
    # Load configuration
    config = load_config()
    debug_options = config.get("debug_options", {})
    log_output_path = debug_options.get("log_output_path")

    # Configure logging
    log_handlers = [logging.StreamHandler(sys.stdout)]
    if log_output_path:
        try:
            log_dir = pathlib.Path(log_output_path).parent
            log_dir.mkdir(parents=True, exist_ok=True)
            log_handlers.append(logging.FileHandler(log_output_path, mode='w'))
        except Exception as e:
            # Can't use logger here.
            print(f"Failed to configure file logging: {e}. "
                  f"Falling back to console logging.", file=sys.stderr)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=log_handlers,
        force=True
    )

    logging.info("Starting book generation process from CLI...")
    run_generation_process()
    logging.info("Book generation process finished.")

def main():
    """Main entry point for the application."""
    parser = argparse.ArgumentParser(description="LLM Book Generator")
    parser.add_argument('--cli', action='store_true', help="Run in command-line interface mode.")
    args = parser.parse_args()

    # Configure Matplotlib to use LaTeX and load amsmath
    matplotlib.rcParams["text.usetex"] = True
    matplotlib.rcParams["text.latex.preamble"] = (
        r"\usepackage{amsmath}  \usepackage{amssymb}"
    )
    # Ensure matplotlib doesn't try to use a GUI backend
    matplotlib.pyplot.switch_backend("Agg")

    if args.cli:
        run_cli()
    else:
        # Run the GUI application
        try:
            from gui import main as run_gui
            run_gui()
        except ImportError as e:
            print(f"Could not import GUI components: {e}", file=sys.stderr)
            print("Please ensure PySide6 is installed (`pip install PySide6`).", file=sys.stderr)
            sys.exit(1)


if __name__ == "__main__":
    main()
