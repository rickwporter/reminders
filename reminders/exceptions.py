from typing import List


class MissingUser(Exception):
    """
    Simple wrapper to format a missing user correctly.
    """
    def __init__(self, username: str, record: str, row: str):
        message = f"Missing '{username}' on {record}, row={row}"
        super().__init__(message)


class AmbiguousUser(Exception):
    """
    Simple wrapper to format errors
    """
    def __init__(self, username: str, matches: List[str]):
        message = f"Ambiguous result for '{username}' matches: {', '.join(matches)}"
        super().__init__(message)
