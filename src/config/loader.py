import logging
import logging.config
import yaml
import os
from pathlib import Path
from dotenv import load_dotenv

def load_config(config_path: str = "src/config/setting.yaml"):
    """
    Load configuration from YAML and .env (for sensitive values like token).
    """
    # Load environment variables from .env file
    load_dotenv()

    config_file_path = Path(config_path).resolve()
    if not config_file_path.is_file():
        raise FileNotFoundError(f"Configuration file not found: {config_file_path}")

    with open(config_file_path, "r") as file:
        config = yaml.safe_load(file)

    # Inject token from environment into influxdb config
    if "influxdb" not in config:
        config["influxdb"] = {}

    config["influxdb"]["token"] = os.getenv("INFLUX_TOKEN")

    if not config["influxdb"].get("token"):
        raise ValueError("❌ INFLUX_TOKEN not set in .env file")

    # Optional: Set up logger
    if "logging" in config:
        logging.config.dictConfig(config["logging"])
        config["logger"] = logging.getLogger("main")

    return config
