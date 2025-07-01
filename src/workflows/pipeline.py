from prefect import flow, task
from data.downloader import download_csv
from data.transformer import preprocess_file
from data.writer import write_to_influxdb
from config.loader import load_config
import logging

log = logging.getLogger("root")

@task(retries=3)
def download_data_task():
    """
    Task to download the data using Selenium.
    """
    log.info("Starting download task...")
    file_path, time_status = download_csv()
    log.info(f"Downloaded file to: {file_path}")
    return file_path, time_status

@task
def preprocess_data_task(file_path, time_status, config):
    """
    Task to preprocess and validate the data.
    """
    log.info("Starting data preprocessing task...")
    validated_data = preprocess_file(file_path, time_status, config["data_tags"])
    log.info("Data preprocessing complete.")
    return validated_data

@task(retries=3)
def write_data_task(validated_data, config):
    """
    Task to write validated data to InfluxDB.
    """
    log.info("Starting data write task...")
    write_to_influxdb(config, validated_data)
    log.info("Data successfully written to InfluxDB.")

@flow(name="bf-data-200var")
def data_pipeline():
    """
    Prefect flow to orchestrate the data pipeline.
    """
    log.info("Starting data pipeline...")
    config = load_config()
    
    # Execute tasks in sequence
    file_path, time_status = download_data_task()
    validated_data = preprocess_data_task(file_path, time_status, config)
    write_data_task(validated_data, config)

    log.info("Data pipeline completed successfully.")
