from .save_excel import save_excel_func


def clean_id_from_link(link: str):
    return link[28:link.find('?')].replace('/', '')
