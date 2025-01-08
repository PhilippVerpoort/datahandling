from pathlib import Path
from typing import Optional, Final

import pandas as pd


FILE_TYPE_XLS = 'xls'
FILE_TYPE_CSV = 'csv'
FILE_TYPE_CONFIG_COMMON = ['skiprows', 'nrows', 'decimal', 'usecols', 'names']
FILE_TYPE_CONFIG = {
    FILE_TYPE_XLS: FILE_TYPE_CONFIG_COMMON + ['sheet_name'],
    FILE_TYPE_CSV: FILE_TYPE_CONFIG_COMMON + ['sep', 'quotechar'],
}
FILE_SEP_OPTIONS = {
    ',': 'Comma (,)',
    ';': 'Semicolon (;)',
    '\t': 'Tab',
}
FILE_QUOTE_OPTIONS = {
    '"': 'Single quote (")',
    "'": "Double quote (')",
}


def excel_col_name(n: int):
    if not isinstance(n, int) or n < 0:
        raise ValueError('Number must be positive')
    ret = ''
    n += 1
    while n:
        n = n - 1
        ret = chr(n % 26 + 65) + ret
        n = n // 26
    return ret


max_attempts: Final[int] = 30
def _my_read_csv(file_path: Path, **read_kws):
    read_kws['skip_blank_lines'] = False
    read_width = None
    for i in range(max_attempts):
        try:
            if read_width is not None:
                read_kws['names'] = list(range(read_width))
            return pd.read_csv(file_path, **read_kws)
        except pd.errors.ParserError as ex:
            try:
                read_width = int(str(ex).split()[-1])
            except:
                continue

    raise Exception('Could not determine format of CSV file.')


class DataHandle:
    _file_path: None | Path
    _file_type: None | str
    _file_config: dict
    _stats: dict

    def __init__(self, file_path: Path, file_type: str):
        self._file_path = file_path
        self._file_type = file_type

        self._file_config = {
            'first': None,
            'last': None,
            'decimal': ',',
            'sheet_name': 0,  # only excel
            'sep': ';',  # only csv
            'quotechar': '"',  # only csv
        }
        self._stats = {}

    @property
    def file_path(self):
        return self._file_path

    @file_path.setter
    def file_path(self, file_path: Path):
        self._file_path = file_path

    @property
    def file_type(self):
        return self._file_type

    @property
    def file_config(self) -> dict:
        return self._file_config

    def update_file_config(self, **file_config_updated):
        self._file_config |= file_config_updated

    @property
    def stats(self) -> dict:
        return self._stats

    @property
    def title(self):
        return self._file_path.name

    def preview(self, file_config: Optional[dict] = None):
        df = self.read(file_config=file_config)
        self._stats['len_rows'] = len(df)
        self._stats['len_columns'] = len(df.columns)
        return df

    def read(self, file_config: Optional[dict] = None):
        # load default file config from datafile object
        file_config = self._file_config | (file_config if file_config is not None else {})

        # set default read keywords for all read functions
        read_kws = {'index_col': False, 'header': None, 'dtype': 'object'}

        # add read_kws from file_config
        read_kws |= {
            key: val
            for key, val in file_config.items()
            if key in FILE_TYPE_CONFIG[self._file_type]
        }

        # read file via pandas, depending on file type
        if self._file_type == FILE_TYPE_XLS:
            df = pd.read_excel(self._file_path, **read_kws)
        elif self._file_type == FILE_TYPE_CSV:
            df = _my_read_csv(self._file_path, **read_kws)
        else:
            if self._file_type is None:
                raise Exception('File type not set.')
            else:
                raise Exception(f"Unknown file type {self._file_type}.")

        df.index += 1

        return df
