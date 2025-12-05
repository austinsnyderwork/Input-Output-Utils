from pathlib import Path
import logging
import json
import os
from typing import Any, Optional


class JsonManager:
    """
    Manages saving/loading a *single* JSON dictionary to/from a file.
    Mirrors the structure and behavior style of JsonlManager.
    """

    def __init__(self,
                 path: Path,
                 logger: Optional[logging.Logger] = None):
        self.path = Path(path)

        # Ensure the file exists
        if not self.path.exists():
            # Create an empty JSON file (empty dict)
            with open(self.path, "w", encoding="utf-8") as f:
                f.write("{}")

        # Optional logger
        if logger is None:
            self._logger = logging.getLogger("dummy")
            self._logger.addHandler(logging.NullHandler())
        else:
            self._logger = logger

    def save(self,
             data: dict,
             existing_ok: bool = True):
        """
        Saves a dictionary into the JSON file.
        If existing_ok is False and the file already has data,
        an error is raised.
        """
        if not isinstance(data, dict):
            raise TypeError(f"JsonManager.save requires a dict, got {type(data)}")

        # Detect existing data
        if not existing_ok:
            if os.path.getsize(self.path) > 0:
                current = self.read()
                if current:  # not empty dict
                    raise ValueError(
                        f"existing_ok=False but JSON file already contains data:\n{self.path}"
                    )

        # Write JSON
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)

        self._logger.info(f"Saved JSON to {self.path}")

    def read(self) -> dict:
        """
        Loads and returns the dictionary stored in the JSON file.
        If the file is empty or corrupted, returns {}.
        """
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if not content:
                    return {}
                return json.loads(content)
        except json.JSONDecodeError:
            self._logger.error(f"JSON decode error in {self.path}. Returning empty dict.")
            return {}

    def has_data(self) -> bool:
        data = self.read()
        return bool(data)

    def delete_all_data(self):
        """
        Clears the JSON file and replaces it with an empty dict.
        """
        with open(self.path, "w", encoding="utf-8") as f:
            f.write("{}")

        self._logger.info(f"Cleared JSON file: {self.path}")
