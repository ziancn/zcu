import logging


def default_log_config(level=logging.INFO):
    logging.basicConfig(
        level=level,
        format="%(asctime)s - [%(levelname)s] - %(message)s",
    )