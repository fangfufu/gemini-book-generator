import logging
import sys
import yaml

CONFIG_FILE_PATH = "config.yaml"

def get_config_path():
    """Returns the path to the configuration file."""
    return CONFIG_FILE_PATH

def load_config(config_path=CONFIG_FILE_PATH):
    """Loads the configuration file."""
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
        logging.info(f"Configuration loaded successfully from {config_path}")
        return config
    except FileNotFoundError:
        logging.error(f"Error: Configuration file not found at {config_path}")
        sys.exit(1)
    except yaml.YAMLError as e:
        logging.error(f"Error parsing configuration file {config_path}: {e}")
        sys.exit(1)
