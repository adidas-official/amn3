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

def message(name, sheet_index, data={}, new=False, pension=True) -> str:
    """Returns message for logging."""
    message  = f"{name}, sheet:{sheet_index}, new:({new}), pension:({pension})\n"
    for month in data:
        try:
            message += f" - month: {month}, payout: {data[month]['payout']}, cell: {data[month]['cell']}\n"
        except KeyError as err:
            message += f" - month: {month}, payout: {data[month]['payout']}, cell: {data[month]['cell']}\n"
            message += f"KeyError: {err}"
    message += "_"*64
    return message

def msgrow(name, sheet_index, data={}, new=False, pension=True) -> dict:
    """Returns message in form of a list to display on front-end."""
    row = {"name": name, "sheet": sheet_index, "new": new, "pension": pension, "months": []}
    for month in data:
        row["months"].append({"month": month, "payout": data[month]["payout"], "cell": data[month]["cell"]})
    return row

