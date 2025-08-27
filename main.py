import logging
import matplotlib
import pathlib
from book_generator.config import load_config
from book_generator.orchestrator import run_generation_process

if __name__ == "__main__":
    # Configure Matplotlib to use LaTeX and load amsmath
    matplotlib.rcParams["text.usetex"] = True
    matplotlib.rcParams["text.latex.preamble"] = (
        r"\usepackage{amsmath}  \usepackage{amssymb}"
    )

    # Ensure matplotlib doesn't try to use a GUI backend
    matplotlib.pyplot.switch_backend("Agg")

    # Load configuration
    config = load_config()
    debug_options = config.get("debug_options", {})
    log_output_path = debug_options.get("log_output_path")

    # Configure logging
    log_handlers = []
    if log_output_path:
        try:
            # Create the directory if it doesn't exist
            log_dir = pathlib.Path(log_output_path).parent
            log_dir.mkdir(parents=True, exist_ok=True)
            log_handlers.append(logging.FileHandler(log_output_path, mode='w'))
        except Exception as e:
            logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
            logging.error(f"Failed to configure file logging: {e}. Falling back to console logging.")

    # Always add a console logger
    log_handlers.append(logging.StreamHandler())

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=log_handlers
    )

    run_generation_process()
