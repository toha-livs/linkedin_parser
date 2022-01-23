from typing import Union, Optional, List

import xlwt


def save_excel_func(data: List[Union[tuple, dict]], headers: list, path: str, file_name: str, sheet_name: Optional[str] = 'default'):
    if not data:
        print('Excel не сохранился, данных для сохранения нет...')
        return
    excel_file = xlwt.Workbook()
    sheet = excel_file.add_sheet(sheet_name, cell_overwrite_ok=True)

    # write headers
    for ind, header in enumerate(headers):
        sheet.write(0, ind, header)

    for row, candidate in enumerate(data):
        row += 1

        is_candidate_dict = True
        if isinstance(candidate, (list, tuple)):
            is_candidate_dict = False

        for ind, header in enumerate(headers):

            try:
                if is_candidate_dict:
                    value = candidate[header]
                else:
                    value = candidate[ind]
            except (IndexError, KeyError):
                value = None

            sheet.write(row, ind, str(value))

    excel_file.save(f'{path}{file_name}.xlsx')
    print(f'В файл excel сохранено {len(data)} записей, по пути {path}{file_name}.xlsx')