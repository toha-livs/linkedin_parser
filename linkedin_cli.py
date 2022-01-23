import os
import sys
import time
import json
import traceback

import xlwt
import datetime
import subprocess
from random import randint
from linkedin_api import Linkedin
from tinydb import TinyDB, where
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains

import logging
from selenium.webdriver.remote.remote_connection import LOGGER as LOGGER_SELENIUM

from utils import save_excel_func, clean_id_from_link

LOGGER_SELENIUM.setLevel(logging.ERROR)


profile = webdriver.ChromeOptions()
# 1 - Allow all images
# 2 - Block all images
# 3 - Block 3rd party images
prefs = {"profile.managed_default_content_settings.images": 2}
profile.add_argument('--headless')
profile.add_experimental_option("prefs", prefs)


def save_to_file(file_path, content):
    with open(file_path, 'w', encoding="utf-8") as fo:
        json.dump(content, fo, ensure_ascii=False)


def add_to_list(path):
    json_files = [f for f in os.listdir(path) if f.endswith('.json')]
    people = []

    for i in range(len(json_files)):
        with open(path + json_files[i]) as file:
            data = json.load(file)
            people.append(data)

    return people


def linkedin_login(login, password, driver):
    driver.get('https://www.linkedin.com/login')
    login_input = driver.find_element(By.XPATH, '//*[@id="username"]')
    pwd_input = driver.find_element(By.XPATH, '//*[@id="password"]')
    signupbtn = driver.find_element_by_css_selector('.btn__primary--large')
    login_input.clear()
    login_input.send_keys(login)
    time.sleep(randint(1, 2))
    pwd_input.clear()
    pwd_input.send_keys(password)
    time.sleep(randint(1, 2))
    signupbtn.click()


def parse_from_linkedin_search(url, page, path, driver):
    search_url = f'{url}&page={page}'
    driver.get(search_url)
    time.sleep(3)
    driver.execute_script("window.scrollTo(0, 450);")
    time.sleep(3)
    driver.execute_script("window.scrollTo(0, 850);")
    time.sleep(1)
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    # urls = driver.find_elements_by_class_name('app-aware-link')
    # urls = [url.get_attribute("href") for url in urls]
    candidates = []
    for row in range(10):
        li_index = row + 1
        try:
            link = driver.find_element_by_xpath(f'//*[@id="main"]/div/div/div[2]/ul/li[{li_index}]/div/div/div[2]/div[1]/div[1]/div/span[1]/span/a').get_attribute('href')
        except NoSuchElementException:
            link = None
        try:
            position = driver.find_element_by_xpath(f'//*[@id="main"]/div/div/div[2]/ul/li[{li_index}]/div/div/div[2]/div[1]/div[2]/div/div[1]').text
        except NoSuchElementException:
            position = None
        candidates.append({'position': position, 'linkedin': link})
    # //*[@id="main"]/div/div/div[2]/ul/li[1]/div/div/div[2]/div[1]/div[1]/div/span[1]/span/a
    # //*[@id="main"]/div/div/div[2]/ul/li[2]/div/div/div[2]/div[1]/div[1]/div/span[1]/span/a
    # //*[@id="main"]/div/div/div[2]/ul/li[1]/div/div/div[2]/div[1]/div[2]/div/div[1]
    # positions = driver.find_elements_by_class_name('entity-result__primary-subtitle')
    # position_list = [i.text for i in positions]
    # profiles_url = re.findall(r'https:\/\/www\.linkedin\.com\/in\/\S+', ' '.join(urls))
    # links = list(dict.fromkeys(profiles_url))
    # links = [link for link in links if link.replace('https://www.linkedin.com/in/', '').replace('/', '')[:2] != 'AC']
    # print(len(links))

    # for ind, link in enumerate(links):
    for candidate in candidates:
        name = candidate['linkedin'][28:candidate['linkedin'].find('?')].replace('/', '')  # removed  'https://www.linkedin.com/in/'
        parser = candidate
        save_to_file(f'{path}/nsort/{name}.json', parser)


def start_parse(start, end, login, password, url, path, excel: str = None):
    def parse_candidates(_page, _driver):
        search_url = f'{url}&page={_page}'
        _driver.get(search_url)
        time.sleep(3)
        _driver.execute_script("window.scrollTo(0, 450);")
        time.sleep(3)
        _driver.execute_script("window.scrollTo(0, 850);")
        time.sleep(1)
        _driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        _candidates = []
        for row in range(10):
            li_index = row + 1
            try:
                link = _driver.find_element_by_xpath(
                    f'//*[@id="main"]/div/div/div[2]/ul/li[{li_index}]/div/div/div[2]/div[1]/div[1]/div/span[1]/span/a').get_attribute(
                    'href')
            except NoSuchElementException:
                link = None
            try:
                position = _driver.find_element_by_xpath(
                    f'//*[@id="main"]/div/div/div[2]/ul/li[{li_index}]/div/div/div[2]/div[1]/div[2]/div/div[1]').text
            except NoSuchElementException:
                position = None
            _candidates.append({'position': position, 'linkedin': link})
        return _candidates

    def save_json_candidates(_candidates):
        for candidate in _candidates:
            name = clean_id_from_link(candidate['linkedin'])
            save_to_file(f'{path}/nsort/{name}.json', candidate)


    with webdriver.Chrome(executable_path=f'{os.getcwd()}/chromedriver', chrome_options=profile) as driver:  # firefox_profile=firefox_profile
        driver.set_page_load_timeout(60)
        pages = [i for i in range(int(start), int(end) + 1)]
        linkedin_login(login, password, driver)

        candidates = []

        for page in pages:

            candidates.extend(parse_candidates(_page=page, _driver=driver))

            time.sleep(randint(1, 5))

            print(f'Парсинг страницы номер {str(page)}')

        save_json_candidates(candidates)
        print('json файлы сохранены')
        if excel:
            try:
                save_excel_func(
                    data=candidates,
                    headers=['linkedin', 'position'],
                    path=path,
                    file_name=excel,
                    sheet_name='nsort'
                )
                print(f"в excel сохранено {len(candidates)}")
            except:
                print(f'Ошибка сохранения excel...\n\n{traceback.format_exc()}')

    return menu(path)


def sort_for_parse(words, path, excel=None):
    data = add_to_list(path + 'nsort/')

    excel_added = []

    for item in data:
        parser = {
            'name': clean_id_from_link(item['linkedin']),
            'position' : item['position'],
            'linkedin' : item['linkedin']
        }
        keywords = str(item['position'])

        if 'Looking' in keywords:
            continue
        if 'HR' in keywords:
            continue
        if 'Recruiter' in keywords:
            continue

        for word in words:
            if str(word) in keywords:
                save_to_file(f'{path}sort/{parser["name"]}.json', parser)
                excel_added.append(parser)
                break
    if excel:
        save_excel_func(
            data=excel_added,
            headers=['name', 'position', 'linkedin'],
            path=path,
            file_name=excel,
            sheet_name='sort'
        )
    return menu(path)


def get_timedelta(el):
    start = el['timePeriod']['startDate']
    end = el['timePeriod'].get('endDate')
    start_date = datetime.date(start['year'], start.get('month', 1), 1)
    end_date = (datetime.date(end['year'], end.get('month', 1), 1) if end
                else datetime.datetime.today())

    delta = ((start_date.year - end_date.year) * 12 +
            start_date.month - end_date.month)

    return delta if delta != 0 else 1


def full_parser(path, api, start_slice, end_slice):
    counter = 0
    data = add_to_list(path + 'sort/')
    print(f'Найдено {len(data)} файлов для полного парсинга')

    for profile in data[int(start_slice) - 1:int(end_slice)]:
        profile = profile['linkedin']
        linkedin_url = profile

        profile_name = clean_id_from_link(profile)
        # profile_name = profile.replace('https://www.linkedin.com/in/', '').replace('/', '').replace('?locale=en_US', '')

        page = api.get_profile(public_id=profile_name)

        if page != {}:
            pass
        else:
            continue

        firstname = page['firstName']
        lastname = page['lastName']
        profile_id = page['profile_id']
        positions = []

        try:
            location = page['locationName']
        except:
            location = None

        try:
            experience = abs(sum([get_timedelta(el) for el in page['experience']])) if len(page['experience']) > 0 else 0
        except:
            experience = 0

        for i in range(len(page['experience'])):

            try:
                company_name = page['experience'][i]['companyName']
                title = page['experience'][i]['title']
                start_time = page['experience'][i]['timePeriod']['startDate']

                try:
                    end_time = page['experience'][i]['timePeriod']['endDate']
                except:
                    end_time = None

                positions_result = {'company_name' : company_name, 'title' : title, 'start_time' : start_time, 'end_time' : end_time}
                positions.append(positions_result) 

            except:
                continue 

        skills = []

        for i in range(len(page['skills'])):
            skills.append(page['skills'][i]['name'])

        education = []

        for i in range(len(page['education'])):
            try:
                university = page['education'][i]['schoolName']
            except:
                university = None

            try:
                start_education = (page['education'][i]['timePeriod']['startDate'])
                end_education = (page['education'][i]['timePeriod']['endDate'])
            except:
                start_education = None
                end_education = None

            education_result = {'name' : university, 'start_education' : start_education, 'end_education' : end_education}
            education.append(education_result)

        contact_info = api.get_profile_contact_info(profile_name) # GET a profiles contact info
        email = contact_info['email_address']
        phone_numbers = contact_info['phone_numbers']

        parser = {'profile_id': profile_id, 'firstname': firstname, 'lastname': lastname, 'linkedin_link': linkedin_url, 'location': location, 'experience': experience, 'positions': positions, 'skills': skills, 'education': education, 'email': email, 'phone_numbers' : phone_numbers}
        save_to_file(f'{path}full/{str(profile_id)}.json', parser)
        counter += 1
        print(f'Сохранено {str(counter)} профилей')
        
        if counter > 300:
            print('Сохранено больше 300 профилей с 1-го аккаунта. Угроза бана, смените аккаунт, завершаю работу...')

            return menu(path)

def start_full_parse(path):
    print('Введите код vpn сервера:')
    server = input()
    if os.environ.get('login'):
        login_full = os.environ['login']
    else:
        print('Введите логин linkedin аккаунта для парсинга:')
        login_full = str(input())
    if os.environ.get('password'):
        password_full = os.environ['password']
    else:
        print('Введите пароль linkedin аккаунта для парсинга:')
        password_full = str(input())
    print('Введите начальный срез:')
    start_slice = str(input())
    print('Введите конечный срез:')
    end_slice = str(input())
    print('Запускаю процесс полного парсинга, напоминаю, что с одного аккаунта нельза парсить больше 300 профилей в сутки, после окончания произойдёт возврат в главное меню')
    # subprocess.run(['sudo', 'nordvpn', 'c', server])
    api = Linkedin(login_full, password_full)

    # try:
    full_parser(path, api, start_slice, end_slice)
    # except Exception as e:
    #     print(e)

    # subprocess.run(['sudo', 'nordvpn', 'd'])

    return menu(path)


def sort_for_invite(path, exp_start, exp_end, skills_list, excel=None):
    data = add_to_list(path + 'full/')

    added_to_excel = []

    for item in data:
        parser = {
            'linkedin': item["linkedin_link"],
            'fill_name': f"{item['lastname']} {item['firstname']}",
            'skills': item['skills'],
            'positions': item['positions'],
            'education': item['education']
        }
        skills = item["skills"]
        exp = item["experience"]

        if exp >= exp_start is False:
            continue
        if exp_end != 0 and exp <= exp_end:
            continue

        for skill in skills_list:
            if skill in skills:
                save_to_file(f'{path}invite/{item["firstname"]} {item["lastname"]}.json', parser)
                added_to_excel.append(parser)
                break

    if excel:
        save_excel_func(
            data=added_to_excel,
            headers=['fill_name', 'linkedin', 'skills', 'positions', 'education'],
            path=path,
            file_name=excel,
            sheet_name='invite'
        )

    print(f'Всего подходящих для инвайтинга и конвертации в xslx файл {len(added_to_excel)}')

    return menu(path)


def invite(path, api, invite_start, invite_end):
    data = add_to_list(path + 'invite/')
    counter = 0

    for candidate in data[int(invite_start) - 1:int(invite_end)]:
        link = candidate['linkedin'].replace('https://www.linkedin.com/in/', '').replace('/', '')
        r = api.add_connection(link)
        counter += 1
        print(link)
        print(r)
        print(f'Наинвайчен {counter} пользователь')

    if counter > 70:
        print('Наинвайчено больше 70 профилей с 1-го аккаунта. Угроза бана, смените аккаунт, завершаю работу...')

        return menu(path)

def start_inviter(path):
    print('Введите код vpn сервера:')
    server = input()
    print('Введите логин linkedin аккаунта для инвайтинга:')
    login_invite = str(input())
    print('Введите пароль linkedin аккаунта для инвайтинга:')
    password_invite = str(input())
    print('Введите начальный срез:')
    invite_start = str(input())
    print('Введите конечный срез:')
    invite_end = str(input())
    print('Запускаю процесс инвайтинга, напоминаю, что с одного аккаунта нельза инвайтить больше 70 профилей в сутки, после окончания произойдёт возврат в главное меню')
    subprocess.run(['sudo', 'nordvpn', 'c', server])
    api = Linkedin(login_invite, password_invite)

    try:
        invite(path, api, invite_start, invite_end)
    except Exception as e:
        print(e)

    subprocess.run(['sudo', 'nordvpn', 'd'])

    return menu(path)

def xlsx_writer(path, filename):
    db = TinyDB('helpers/contacts.json')
    data = add_to_list(path + 'invite/')
    excel_file = xlwt.Workbook()
    sheet = excel_file.add_sheet('with contacts', cell_overwrite_ok=True)
    sheet.write(0, 0, 'Linkedin')
    sheet.write(0, 1, 'Skype')
    sheet.write(0, 2, 'Email')
    sheet.write(0, 3, 'Phone')
    sheet.write(0, 4, 'Facebook')
    counter = 0
    row = 1

    for item in data:
        contacts = db.search(where('link') == item['linkedin'])

        if contacts != []:
            skype = contacts[0]['skype'] if not [] else None
            email = contacts[0]['email'] if not [] else None
            phone = contacts[0]['phone'] if not [] else None

            sheet.write(row, 0, item['linkedin'])
            sheet.write(row, 1, skype)
            sheet.write(row, 2, email)
            sheet.write(row, 3, phone)
            sheet.write(row, 4, contacts[0]['facebook'])
            row += 1
        else:
            sheet.write(row, 0, item['linkedin'])
            sheet.write(row, 1, None)
            sheet.write(row, 2, None)
            sheet.write(row, 3, None)
            sheet.write(row, 4, None)
            row += 1
        
        counter += 1

    excel_file.save(f'{path}{filename}.xlsx')
    print(f'Записано ячеек {counter}')

    return menu(path)

def invite_witg_msg(start, end, login, password, msg, path):
    driver = webdriver.Chrome(executable_path=f'{os.getcwd()}/chromedriver', chrome_options=profile) # firefox_profile=firefox_profile
    driver.set_page_load_timeout(60)
    linkedin_login(login, password, driver)
    data = add_to_list(path + 'invite/')
    print(f'Длинна сообщения - {len(msg)}')

    counter = 0

    for item in data[start - 1: end]:
        driver.get(item['linkedin'])
        print(item['linkedin'])
        time.sleep(randint(1, 2))

        try:
            connect_btn = driver.find_element_by_class_name('pv-s-profile-actions--connect').click()
            time.sleep(randint(2, 3))
            msg_next = driver.find_element_by_class_name('artdeco-button--secondary')
            ActionChains(driver).move_to_element(msg_next).click().perform()
            time.sleep(randint(2, 3))
            textarea = driver.find_element_by_css_selector('#custom-message').send_keys(msg) #custom-message
            time.sleep(randint(2, 3))
            send_invite = driver.find_element_by_class_name('ml1')
            ActionChains(driver).move_to_element(send_invite).click().perform()
            counter += 1
            print('Инвайт отправлен.')
            
            if counter > 70:
                print('Наинвайчено больше 70 профилей с 1-го аккаунта. Угроза бана, смените аккаунт, завершаю работу...')

                return menu(path)
        except Exception as e:
            print(f'Erroe - {e}')
            continue

    print(f'{counter} заинвайчено')
    driver.quit()

    return menu(path)

def send_msg(path, api, msg):
    data = add_to_list(path + 'invite/')

    for item in data:
        recipt = item["linkedin"].replace('https://www.linkedin.com/in/', '').replace('/', '').replace('?locale=en_US', '')

        try:
            profile = api.get_profile(recipt)
            conv_id = api.get_conversation_details(profile_urn_id=profile["profile_id"])
            api.send_message(conversation_urn_id=conv_id["id"], message_body=msg)
            print(f'Сообщение {item["linkedin"]} отправлено')
        except Exception as e:
            print(e)
            continue

    subprocess.run(['sudo', 'nordvpn', 'd'])
    return menu(path)

def search_contacts(path):
    db = TinyDB('helpers/contacts.json')
    print('Введите ссылку на linkedin аккаунт (Важно! Ссылка должна быть без / в конце):')
    link = str(input())
    contacts = db.search(where('link') == link)

    if contacts != []:
        skype = str(contacts[0]['skype']).replace("'", "").replace("[", "").replace("]", "") if not [] else None
        email = str(contacts[0]['email']).replace("'", "").replace("[", "").replace("]", "") if not [] else None
        phone = str(contacts[0]['phone']).replace("'", "").replace("[", "").replace("]", "") if not [] else None
        facebook = contacts[0]['facebook'] if not [] else None
        print(f'''
Skype - {skype}
Email - {email}
Phone - {phone}
Facebook - {facebook}
''')
    else:
        print('Не найдено в базе.')

    return menu(path)

def menu(path):
    print(
f'''Чтобы запустить парсер из поиска линкедин введите 1.
Чтобы запустить сортировку для дальнейшего полного парсинга введите 2.
Чтобы запустить полный парсинг введите 3.
Чтобы запустить сортировку для дальнейшего инвайтинга или формирования в xlsx файл введите 4.
Чтобы запустить инвайтинг введите 5.
Чтобы запустить запись в xlsx файл введите 6.
Чтобы запустить инвайтинг с сообщением введите 7.
Чтобы запустить рассылку друзьям введите 8.
Чтобы запустить поиск контактов по базе введите 9. 
Для завершения работы введите 0.'''
    )

    choice = int(input())

    if choice == 1:
        print('Чтобы продолжить парсинг введите 1, чтобы вернуться в меню введите 0:')
        re_choice = int(input())

        if re_choice == 0:
            return menu(path)
        elif re_choice == 1:
            print('Введите логин для linkedin:')
            login = str(input())
            print('Введите пароль для linkedin:')
            password = str(input())
            print('Введите номер страницы для старта парсинга:')
            start = int(input())
            print('Введите номер страницы для окончания парсинга:')
            end = int(input())
            print('Введите url для парсинга:')
            url = str(input())
            excel = input('Если хотите сохранить результат в excel, введите название файла\n(Оставьте пустым если не желаете):\n')
            print(
                'Запускаю парсер, пожалуйста не закрывайте окно браузера, оно закроется автоматически по окончанию работы и вернёт вас в главное меню...'
            )
            start_parse(start, end, login, password, url, path, excel=excel)
    elif choice == 2:
        print('Чтобы продолжить сортировку введите 1, чтобы вернуться в меню введите 0:')
        re_choice = int(input())

        if re_choice == 0:
            return menu(path)
        elif re_choice == 1:
            words = []

            while True:
                print('Введите слово для сортировки, чтобы закончить с вводом слов введите 0')
                word = input()

                if word == '0':
                    break
                else:
                    words.append(word)

            excel = input('Если хотите сохранить результат в excel, введите название файла\n(Оставьте пустым если не желаете):\n')
            
            print('Запускаю процесс сортировки, после окончания произойдёт возврат в главное меню')
            sort_for_parse(words, path, excel)
    elif choice == 3:
        print('Чтобы продолжить полный парсинг введите 1, чтобы вернуться в меню введите 0:')
        re_choice = int(input())

        if re_choice == 0:
            return menu(path)
        elif re_choice == 1:
            start_full_parse(path)
    elif choice == 4:
        print('Чтобы продолжить сортировку введите 1, чтобы вернуться в меню введите 0:')
        re_choice = int(input())

        if re_choice == 0:
            return menu(path)
        elif re_choice == 1:
            print('Введите нижний порог опыта работы (Важно! Опыт работы вводить в месяцах!):')
            exp_start = int(input())
            print('Введите верхний порог опыта работы, если он не нужен введите 0 (Важно! Опыт работы вводить в месяцах!):')
            exp_end = int(input())
            skills_list = []

            while True:
                print('Введите умение, чтобы закончить с вводом умений введите 0')
                skill = input()

                if skill == '0':
                    break
                else:
                    skills_list.append(skill)

            excel = input('Если хотите сохранить результат в excel, введите название файла\n(Оставьте пустым если не желаете):\n')
            print('Запускаю процесс сортировки, после окончания произойдёт возврат в главное меню')
            sort_for_invite(path, exp_start, exp_end, skills_list, excel)
    elif choice == 5:
        print('Чтобы продолжить инвайтинг введите 1, чтобы вернуться в меню введите 0:')
        re_choice = int(input())

        if re_choice == 0:
            return menu(path)
        elif re_choice == 1:
            start_inviter(path)
    elif choice == 6:
        print('Чтобы продолжить запись в xlsx файл введите 1, чтобы вернуться в меню введите 0:')
        re_choice = int(input())

        if re_choice == 0:
            return menu(path)
        elif re_choice == 1:
            print('Введите имя для файла:')
            filename = str(input())
            print(f'Запускаю процесс записи в xlsx файл, файл будет сохранен по пути {path}, после окончания произойдёт возврат в главное меню')
            xlsx_writer(path, filename)
    elif choice == 7:
        print('Чтобы продолжить инвайтинг с сообщением введите 1, чтобы вернуться в меню введите 0:')
        re_choice = int(input())

        if re_choice == 0:
            return menu(path)
        elif re_choice == 1:
            print('Введите логин для linkedin:')
            login = str(input())
            print('Введите пароль для linkedin:')
            password = str(input())
            print('Введите стартовый срез:')
            start = int(input())
            print('Введите конечный срез:')
            end = int(input())
            print('Введите сообщение(Важно! Максимум 300 символов.):')
            msg = "\n".join(iter(input, ""))
            print('Запускаю инвайтер с сообщением, пожалуйста не закрывайте окно браузера, оно закроется автоматически по окончанию работы и вернёт вас в главное меню... (Важно! Не инвайтите больше 70 человек на один профиль.)')
            invite_witg_msg(start, end, login, password, msg, path)
    elif choice == 8:
        print('Чтобы продолжить рассылку введите 1, чтобы вернуться в меню введите 0:')
        re_choice = int(input())

        if re_choice == 0:
            return menu(path)
        elif re_choice == 1:
            print('Введите код vpn сервера:')
            server = input()
            print('Введите логин для linkedin:')
            login = str(input())
            print('Введите пароль для linkedin:')
            password = str(input())
            print('Введите сообщение:')
            msg = "\n".join(iter(input, ""))
            subprocess.run(['sudo', 'nordvpn', 'c', server])
            api = Linkedin(login, password)
            print('Запускаю рассылку...')
            send_msg(path, api, msg)
    elif choice == 9:
        print('Чтобы продолжить поиск по базе контактов введите 1, чтобы вернуться в меню введите 0:')
        re_choice = int(input())

        if re_choice == 0:
            return menu(path)
        elif re_choice == 1:
            search_contacts(path)
    elif choice == 0:
        print('Чтобы завершить работу введите 1, чтобы вернуться в меню введите 0:')
        re_choice = int(input())

        if re_choice == 0:
            return menu(path)
        elif re_choice == 1:
            sys.exit()
    else:
        print('Введен неверный номер, пожалуйста повторите попытку')

        return menu(path)


if __name__ == '__main__':
    if os.environ.get('vacancy'):
        vacancy = os.environ['vacancy']
    else:
        print('Введите название вакансии:')
        vacancy = str(input())
    path = f'{os.getcwd()}/data/{vacancy}/'

    try:
        os.mkdir(path)
        os.mkdir(f'{path}nsort/')
        os.mkdir(f'{path}sort/')
        os.mkdir(f'{path}full/')
        os.mkdir(f'{path}invite/')
    except OSError:
        print('Такая вакансия найдена, продолжаю работу с ней...')

    try:
        menu(path)
    except SystemExit:
        print('Пока...')
    except:
        print(traceback.format_exc())
