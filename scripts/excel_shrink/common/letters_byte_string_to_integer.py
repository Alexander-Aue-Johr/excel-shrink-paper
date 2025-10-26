col_cache = {}

def convert_letters_byte_string_to_integer(col_bytes: bytes) -> int:
    """
    Convert a column-letter bytes object (e.g. b'A', b'B', b'AA', etc.)
    into its 1-based column index (A=1, B=2, ... AA=27, etc.).
    """
    col_index = 0
    for b in col_bytes:
        # Uppercase letters: 65-90 (A-Z)
        # Lowercase letters: 97-122 (a-z)
        if 65 <= b <= 90:
            # A-Z
            col_index = col_index * 26 + (b - 65 + 1)
        elif 97 <= b <= 122:
            # a-z (normalize to uppercase)
            col_index = col_index * 26 + (b - 97 + 1)
    return col_index

def letters_byte_string_to_integer(col_bytes: bytes | None) -> int:
    """Return a cached or newly computed column index for col_b."""
    if col_bytes is None:
        return 0
    if col_bytes not in col_cache:
        col_cache[col_bytes] = convert_letters_byte_string_to_integer(col_bytes)
    return col_cache[col_bytes]