from utils.logger import setup_logger
from datetime import datetime
from workflows.pipeline import data_pipeline
import logging

# Setup the logger at the start of the application
setup_logger()

log = logging.getLogger("root")

if __name__ == "__main__":
    while True:
        start_time = datetime.now()
        log.info(f"Start Time {start_time}")
        log.info("Starting the data pipeline")
        data_pipeline()
        log.info("Finished")
        end_time = datetime.now()
        log.info(f"Took {end_time - start_time} seconds.")