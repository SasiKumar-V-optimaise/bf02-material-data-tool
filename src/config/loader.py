import logging
import logging.config
import yaml
from pathlib import Path

def load_config(config_path: str = "src/config/setting.yaml"):
    """
    Load configuration from a YAML file.
    """
    config_file_path = Path(config_path).resolve()
    if not config_file_path.is_file():
        raise FileNotFoundError(f"Configuration file not found: {config_file_path}")
    with open(config_file_path, "r") as file:
        return yaml.safe_load(file)