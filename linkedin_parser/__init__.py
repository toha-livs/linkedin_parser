import os


class Menu:

    def __init__(self):
        self.current_function = None

    def confirm(self, message=None):
        if message is None:
            message = 'это сделать'
        confirm = input(f'Вы уверенны что хотите  {message}?\nОствьте поле пустым для отмены\nВведите "1" для подтверждения')
        if confirm == '1' and self.current_function:
            self.current_function()
        self.current_function = None

    def __call__(self):
        print('''Чтобы запустить парсер из поиска линкедин введите 1.
Чтобы запустить сортировку для дальнейшего полного парсинга введите 2.
Чтобы запустить полный парсинг введите 3.
Чтобы запустить сортировку для дальнейшего инвайтинга или формирования в xlsx файл введите 4.
Чтобы запустить инвайтинг введите 5.
Чтобы запустить запись в xlsx файл введите 6.
Чтобы запустить инвайтинг с сообщением введите 7.
Чтобы запустить рассылку друзьям введите 8.
Чтобы запустить поиск контактов по базе введите 9. 
Для завершения работы введите 0.\n''')




class LinkedinParserManager:

    def __init__(self):
        self.path = None
        self.session_name = None

    def start(self):
        self.session_name = input('Введите название вакансии:')
        self._setup_path()
        self._request_menu()

    def _request_menu(self):
        pass

    def _setup_path(self):
        path = f'{os.getcwd()}/data/{self.session_name}/'
        try:
            os.mkdir(path)
            os.mkdir(f'{path}nsort/')
            os.mkdir(f'{path}sort/')
            os.mkdir(f'{path}full/')
            os.mkdir(f'{path}invite/')
            print('Создал папку...')
        except OSError:
            print('Такая вакансия найдена, продолжаю работу с ней...')

