import logging
import colorlog

# Creating logger
logger = logging.getLogger(__name__)

# Setting up logging
logger.setLevel(level=logging.DEBUG)

# create a color formatter
formatter = colorlog.ColoredFormatter(
    "%(log_color)s %(asctime)s %(filename)s %(levelname)s:%(message)s",
    log_colors={
        "DEBUG": "blue",
        "INFO": "green",
        "WARNING": "yellow",
        "ERROR": "red",
        "CRITICAL": "red,bg_white",
    },
)

# create a console handler
console_handler = logging.StreamHandler()

# set the formatter for the console handler
console_handler.setFormatter(formatter)

# add the console handler to the logger
logger.addHandler(console_handler)