from influxdb_client_3 import InfluxDBClient3, Point
from config.loader import load_config
from dotenv import load_dotenv
from typing import Dict
import os
import logging

# Load requirements:
config = load_config()
log = logging.getLogger("root")
load_dotenv()

TOKEN = os.getenv("TOKEN")

def write_to_influxdb(config: Dict, validated_data):
    """
    Write validated data to InfluxDB.
    """
    client = InfluxDBClient3(
        host=config["influxdb"]["host"],
        token=TOKEN,
        org=config["influxdb"]["org"]
    )
    log.info(f"Called the influxdb client {client}.")

    for data in validated_data:
        points = [
            Point(data.measurement)
            .tag("sensor", tag)
            .field("value", value)
            .time(data.timestamp)
            for tag, value in data.values.items()
        ]
        if config["influxdb"]["write"].upper() == "TRUE":
            client.write(database=config["influxdb"]["database"], record=points)