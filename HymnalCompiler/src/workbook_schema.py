"""Workbook parsing rules shared by validation and transformation."""


def read_title_page(ws):
    """Return normalized title-page values and any row-specific errors.

    Reject duplicate keys instead of silently replacing earlier values.
    Required values must be nonblank text, as expected by the Typst template.
    """
    info = {}
    seen = {}
    errors = []
    for row_num, (key, value) in enumerate(
        ws.iter_rows(min_row=2, max_col=2, values_only=True), start=2
    ):
        if key is None or (isinstance(key, str) and not key.strip()):
            continue
        if not isinstance(key, str):
            errors.append(f"Title Page row {row_num}, Key: expected text, got {type(key).__name__}")
            continue
        key = key.strip().lower()
        if key in seen:
            errors.append(
                f"Title Page row {row_num}: duplicate key '{key}' "
                f"(first used at row {seen[key]})"
            )
            continue
        seen[key] = row_num
        info[key] = value.strip() if isinstance(value, str) else value

    for key in ("title", "subtitle"):
        value = info.get(key)
        if value is None or value == "":
            errors.append(f"Title Page: missing or empty value for key '{key}'")
        elif not isinstance(value, str):
            errors.append(
                f"Title Page row {seen[key]}, Value (key '{key}'): "
                f"expected text, got {type(value).__name__}"
            )
    return info, errors
