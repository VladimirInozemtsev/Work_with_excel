"""
Стоит весьма простая задача, не требующая особых знаний.
Далее я написал примитивный класс для ЗАПИСИ в эксель.
Твоя задача состоит в том что бы, во-первых, улучшить его, добавить новые фишки или оптимизировать не оптимизированное.
Во-вторых, уже конкретно, наследуя его, переопределить нужные методы и создать класс для записи какого-то конкретного документа,
имея его шаблон, или как я написал ниже, создать начальную конфигурацию этого документа в виде DATAframe.
В data/uploads шаблон авансового отчёта. Только не надо в нём объединять ячейки без нужды! Это плохо.
https://xlsxwriter.readthedocs.io/index.html
"""
import io
import os

import xlsxwriter

from contextlib import suppress
from pathlib import Path
from openpyxl import load_workbook

from django.conf import settings
from django.http import HttpResponse

#from intra.tools.different_functions import decorator_directory_path, user_directory_path


class ExcelWriter:
    """ Первичный класс для создания эксель документов.
    Данный класс требует наследования для, как минимум, переопределения  add """

    def_options = {"in_memory": True}

    def _default_file_path(self):
        return os.path.join("data/advance report template.xlsx")

    def __init__(self, file_path=None, options=None):
        file_path = file_path or self._default_file_path()
        self.wb = load_workbook(file_path) if Path(file_path).exists() else None
        self.output = io.BytesIO()
        self.workbook = xlsxwriter.Workbook(self.output, options=options or self.def_options)

    def write(self, filename=None):
        """ Это исходный метод и потребитель данного класса по идее вызывает только его """
        with self.workbook as workbook:
            if self.wb:
                self._copy(workbook)
            self.configuration(workbook)
            self.add(workbook)
        self.output.seek(0)
        response = HttpResponse(self.output,
                                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        response["Content-Disposition"] = "attachment; filename=%s" % self.validate_filename(filename)
        return response

    def validate_filename(self, filename, default="output.xlsx"):
        if isinstance(filename, str):
            if not filename.endswith('.xlsx'):
                filename += '.xlsx'
        else:
            filename = default
        return filename

    def _copy(self, workbook):
        """ Дописать разных вариантов """
        with self.wb as wb:
            for sheet_name in wb.sheetnames:
                source_sheet = wb[sheet_name]
                worksheet = workbook.add_worksheet(sheet_name)

                # Копируем размеры столбцов
                for col_idx, col_dim in source_sheet.column_dimensions.items():
                    if col_dim.width is not None:
                        worksheet.set_column(f'{col_idx}:{col_idx}', col_dim.width)

                # Копируем размеры строк
                for row in source_sheet.iter_rows():
                    for cell in row:
                        if source_sheet.row_dimensions[cell.row].height:
                            worksheet.set_row(cell.row - 1, source_sheet.row_dimensions[cell.row].height)

                # Копируем данные и стили
                for row_idx, row in enumerate(source_sheet.iter_rows(), start=0):
                    for col_idx, cell in enumerate(row, start=0):
                        # Копируем значение
                        worksheet.write(row_idx, col_idx, cell.value)
                        # Применяем шрифт и стили
                        if cell.has_style:
                            cell_format = workbook.add_format(self._get_cell_format(cell))
                            worksheet.write(row_idx, col_idx, cell.value, cell_format)

    def _get_cell_format(self, cell):
        """ Получает формат для ячейки. """
        format_dict = {}
        font = cell.font
        if font.bold: format_dict['bold'] = True
        if font.italic: format_dict['italic'] = True
        if font.name: format_dict['font_name'] = font.name
        if font.size: format_dict['font_size'] = font.size
        if font.color and font.color.rgb:
            format_dict['font_color'] = f'#{font.color.rgb[2:]}'

        align = cell.alignment
        if align.horizontal: format_dict['align'] = align.horizontal
        if align.vertical: format_dict['valign'] = align.vertical

        return format_dict

    def add(self, workbook):
        """ Добавляем чё надо """
        pass

    def configuration(self, workbook):
        """ Начальная конфигурация.
        Если она есть имеет смысл её представить в виде Dataframe и записать используя Pandas или Polars """
        pass
