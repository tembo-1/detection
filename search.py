from __future__ import annotations

import collections
import gzip
import json
import lzma
import mimetypes
from math import log10
from os import PathLike
from pathlib import Path
from typing import ClassVar, Literal, Dict, NamedTuple, TYPE_CHECKING, Any

import openpyxl
from openpyxl.workbook import Workbook

if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet

FreqType = Dict[Literal['frequency', 'intensity'], float]


class Record(NamedTuple):
    tag: str
    freq: FreqType


class Search:
    error: ClassVar[float] = 0.3
    line_width: ClassVar[float] = 0.6

    def __init__(self, excel_path: str, json_path: str | PathLike[str]) -> None:
        self.json_path: Path = Path(json_path)
        self.excel_path: Path = Path(excel_path)
        self.excel: Workbook = openpyxl.load_workbook(excel_path)

        self.data: dict[str, dict[str, Any]] = {}

        self.__prepareJsonData()

    def __prepareJsonData(self) -> None:
        mimetype: str | None
        encoding: str | None
        mimetype, encoding = mimetypes.guess_type(self.json_path, strict=False)
        if mimetype != 'application/json':
            raise ValueError(f'Not a JSON file: {self.json_path}')
        if encoding is None:
            opener = Path.open
        elif encoding == 'gzip':
            opener = gzip.open
        elif encoding == 'xz':
            opener = lzma.open
        else:
            raise ValueError(f'Unknown type: {self.json_path}')
        with opener(self.json_path, 'r') as file:
            data = json.load(file)
            if 'catalog' in data:
                self.data = data['catalog']
            else:
                raise RuntimeError("No 'catalog' entry in file")
            # TODO: check that the structure in `dict[str, dict[str, Any]]` indeed

    def search(self):
        sheet: Worksheet = self.excel.active

        header = sheet['H1'] 
        header.value = 'Frequency'

        header = sheet['I1'] 
        header.value = 'Intensity'

        header = sheet['J1'] 
        header.value = 'Substance'

        accord = []

        omni_buff: dict[float, list[Record]] = {}

        for row in sheet.iter_rows(min_row=2, min_col=1, max_col=3):
            excel_freq = row[0].value

            if excel_freq is None:
                break

            for tag, array in self.data.items():
                flag: bool = False
                for freq in array['lines']:
                    if freq['frequency'] + Search.error >= excel_freq >= freq['frequency'] - Search.error:
                        flag = True
                        if excel_freq in omni_buff:
                            omni_buff[excel_freq].append(Record(tag, freq))
                        else:
                            omni_buff[excel_freq] = [Record(tag, freq)]

                if flag:
                    break

        common: dict[str, int] = dict(collections.Counter(
            tag for o in omni_buff.values() for tag, freq in o).most_common())

        # combine lines of the same species at the same frequency
        for excel_freq in omni_buff:
            tags: list[str] = [tag for tag, freq in omni_buff[excel_freq]]
            # are there non-unique tags for the frequency?
            if len(tags) != len(set(tags)):
                new_found: list[Record] = []
                for tag, freq in omni_buff[excel_freq]:
                    if tags.count(tag) != 1:
                        # the tag isn't unique
                        # combine the lines
                        # FIXME: shouldn't we consider the frequency offsets of the lines?
                        lines_of_the_tag: list[FreqType] = [f for t, f in omni_buff[excel_freq] if t == tag]
                        total_frequency: float = sum(
                            line['frequency'] * 10 ** line['intensity'] for line in lines_of_the_tag) / sum(
                            10 ** line['intensity'] for line in lines_of_the_tag)
                        total_intensity: float = log10(sum(10 ** line['intensity'] for line in lines_of_the_tag))
                        new_found.append(
                            Record(
                                tag=tag,
                                freq={'frequency': total_frequency, 'intensity': total_intensity},
                            )
                        )
                    else:
                        new_found.append(Record(tag, freq))
                omni_buff[excel_freq] = new_found

        # find the lines that are the most frequently occurred, intense, and closest to the given frequency
        for excel_freq in omni_buff:
            min_weight: float = 0
            min_tag: str | None = None
            min_freq: None | FreqType = None
            for tag, freq in omni_buff[excel_freq]:
                weight: float = 10 ** freq['intensity'] / (
                        1 + (abs(freq['frequency'] - excel_freq) / Search.line_width) ** 2) * common[tag]
                if weight > min_weight:
                    min_weight = weight
                    min_freq = freq
                    min_tag = tag
            if min_tag is not None and min_freq is not None:
                accord.append({self.data[min_tag]['name']: (min_freq['frequency'], min_freq['intensity'])})

        for index, key in enumerate(accord, start=2):
            for key, value in key.items():
                cell = sheet['J'+ str(index)]
                cell.value = key
                
                cell = sheet['H'+ str(index)]
                cell.value = value[0]

                cell = sheet['I'+ str(index)]
                cell.value = value[1]
        
        self.excel.save(self.excel_path)

        # Запись данных в json
        with open('data.json', 'w') as file:
            json.dump(accord, file)


        return accord
