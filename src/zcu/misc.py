"""
Miscellaneous utility functions for the zcu package.
"""


import logging


def config_logging(level=logging.INFO):
    logging.basicConfig(
        level=level,
        format="%(asctime)s - [%(levelname)s] - %(message)s",
    )