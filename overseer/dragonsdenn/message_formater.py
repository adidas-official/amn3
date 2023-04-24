from .logger import logger

data = {
    '07': {
        "payout": 12345,
        "cell": 'c12'
    },
    '08': {
        "payout": 13757,
        "cell": 't12'
    }
}

def message(name, sheet_index, data={}, new=False, pension=True) -> str:
    """Returns message for logging."""
    logger.debug(data)
    message  = f"{name}, sheet:{sheet_index} new:({new}), pension:({pension})\n"
    for month in data:
        try:
            message += f" - month: {month}, payout: {data[month]['payout']}, cell: {data[month]['cell']}\n"
        except KeyError as err:
            message += f" - month: {month}, payout: {data[month]['payout']}, cell: {data[month]['cell']}\n"
            message += f"KeyError: {err}"
    message += "_"*64
    return message

print(message('Cenefels', 2, data, new=False, pension=False))