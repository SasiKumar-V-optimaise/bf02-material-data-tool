from models.data_model import SensorData
from datetime import datetime
from utils.helper_functions_transformer import remove_file, preprocess_timestamp, get_cols
import pandas as pd
from typing import Dict
import logging

log = logging.getLogger("root")

def preprocess_file(file_path: str, time_status: datetime, tags: Dict):
    """
    Read, validate, and preprocess the downloaded file.
    """
    # Load Excel into a DataFrame
    USECOLS = get_cols()
    df = pd.read_csv(file_path, usecols=USECOLS, nrows=1)
    log.info(f"timestep for file - {time_status}")
    time_status = preprocess_timestamp(time_status)
    remove_file()
    # Convert to dictionary for validation
    validated_data = []
    for measurement, variables in tags.items():
        data = {
            "timestamp": time_status,
            "measurement": measurement,
            "values": {var: df[var].iloc[0] for var in variables if var in df.columns}
        }
        validated_data.append(SensorData.validate_data(data))
        log.info(f"Wrote {variables} of {measurement} for {time_status}.")
    return validated_data