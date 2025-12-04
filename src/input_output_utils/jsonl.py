
from pathlib import Path
import pandas as pd
import logging
from typing import Iterable
import json
import os

class JsonlManager:

    def __init__(self,
                 path: Path,
                 logger: logging.Logger = None):
        self.path = path
        if not os.path.exists(self.path):
            open(self.path, "w").close()

        if logger is None:
            self._logger = logging.getLogger("dummy")
            self._logger.addHandler(logging.NullHandler())
        else:
            self._logger = logger

    def save(self, items: Iterable[object | dict]):
        with open(self.path, "a", encoding="utf-8") as f:
            for item in items:
                if isinstance(item, dict):
                    f.write(json.dumps(item) + "\n")
                elif isinstance(item, object):
                    item_d = item.to_dict()
                    json_str = json.dumps(item_d)
                    f.write(json_str + "\n")

    def read(self) -> list:
        jsonl_lines = []
        with open(self.path, "r", encoding="utf-8") as f:
            for line in f:
                line_dict = json.loads(line)  # each line is a dict
                jsonl_lines.append(line_dict)

        return jsonl_lines

    def read_into_dataframe(self) -> pd.DataFrame:
        lines = self.read()

        agg_data = {
            k: []
            for k in lines[0].keys()
        }
        for data_dict in lines:
            missing_keys = [k for k in data_dict.keys() if k not in agg_data]
            if missing_keys:
                random_col = list(data_dict.values())[0]
                current_len = len(random_col) + 1
                for k in missing_keys:
                    agg_data[k] = [None] * current_len

            for k in agg_data.keys():
                v = data_dict.get(k, [None])
                if isinstance(v, Iterable):
                    agg_data[k].extend(v)
                else:
                    agg_data[k].append(v)

        df = pd.DataFrame(agg_data)
        return df

    def delete_all_data(self):
        with open(self.path, "w"):
            pass



