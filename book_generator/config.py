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
        # The logging is not configured yet.
        # logging.info(f"Configuration loaded successfully from {config_path}")
        return config
    except FileNotFoundError:
        print(f"Error: Configuration file not found at {config_path}", file=sys.stderr)
        sys.exit(1)
    except yaml.YAMLError as e:
        print(f"Error parsing configuration file {config_path}: {e}", file=sys.stderr)
        sys.exit(1)
