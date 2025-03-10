import logging
import os

def setup_logging():
    log_directory, log_filename = "logs", "MieTrakAPI.log"
    os.makedirs(log_directory, exist_ok=True)
    log_file_path = os.path.join(log_directory, log_filename)

    logging.basicConfig(
        level=logging.DEBUG,
        format = "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        handlers=[logging.FileHandler(log_file_path, mode='a'), logging.StreamHandler()])
