from config.loader import load_config
from datetime import datetime
from pathlib import Path
from typing import List
import os
import logging

logger = logging.getLogger("root")
config = load_config()

project_root = Path(__file__).resolve().parents[2]
def remove_file() -> None:
  os.remove(project_root / "data" / "data.csv")

def preprocess_timestamp(time_status: datetime) -> datetime:
  return time_status

def get_cols() -> List[str]:
  """
  Load required variables from a YAML configuration file.
  
  Parameters:
      config_path (str): Path to the YAML configuration file.

  Returns:
      list: List of variables to be used as column names.
  """
  # Extract all variables under 'data_tags'
  data_tags = config.get("data_tags", {})
  all_variables = []
  
  for tag_variables in data_tags.values():
      if isinstance(tag_variables, list):  # Ensure the value is a list
          all_variables.extend(tag_variables)