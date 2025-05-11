def number_to_words(amount):
    """
    Convert a numerical amount to words (e.g., 123 -> 'One Hundred Twenty-Three').
    Args:
        amount: Integer or float amount to convert.
    Returns:
        String representation of the amount in words.
    """
    try:
        from num2words import num2words
        amount = int(round(float(amount)))
        return num2words(amount, lang='en').replace(',', '').title()
    except ImportError:
        return str(amount)
    except (ValueError, TypeError):
        return str(amount)

def is_extra_item_sheet_empty(ws_extra):
    """
    Check if the Extra Items sheet DataFrame is empty or contains only null values.
    Args:
        ws_extra: pandas DataFrame representing the Extra Items sheet.
    Returns:
        Boolean indicating if the sheet is empty.
    """
    import pandas as pd
    if not isinstance(ws_extra, pd.DataFrame):
        return True
    return ws_extra.empty or ws_extra.dropna(how='all').empty