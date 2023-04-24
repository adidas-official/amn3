def message(name, sheet_index, data={}, new=False, pension=True) -> str:
    """Returns message for logging."""
    message  = f"{name}, sheet:{sheet_index} new:({new}), pension:({pension})\n"
    for month in data:
        try:
            message += f" - month: {month}, payout: {data[month]['payout']}, cell: {data[month]['cell']}\n"
        except KeyError as err:
            message += f" - month: {month}, payout: {data[month]['payout']}, cell: {data[month]['cell']}\n"
            message += f"KeyError: {err}"
    message += "_"*64
    return message