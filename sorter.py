import os
import re
from win32com.client import Dispatch
from datetime import datetime
import shutil
from pathlib import Path


main_path = 'C:\\UBS\\FORMAT'
os.chdir(main_path)
sorted_path = main_path.split('\\')
sorted_folder = (Path.home() / 'Desktop')

list_of_file = os.listdir(main_path)


def create_shortcut(path, target='', work_dir='', icon=''):
    shell = Dispatch('WScript.shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = work_dir
    shortcut.IconLocation = icon
    shortcut.save()


sum = 0
time_start = datetime.now()

print('Сортирую файлы в каталоге: C:\\UBS\\FORMAT')


if not os.path.exists(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ'):
    os.makedirs(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ')

loop = 0

while loop <= 1:
    for i in range(len(list_of_file)):  # create dir from brand names

        folder = list_of_file[i].split('_', 1)

        if not os.path.exists(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}'):
            os.makedirs(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}')

        sub_folder = re.findall(r'\-\d{2}\-\d{1}', folder[-1])

        if len(sub_folder) > 0:
            if sub_folder[0] in folder[-1]:
                sub_folder = sub_folder[0].replace('-', '', 1)

                try:
                    # create subfolders for files
                    if not os.path.exists(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}\\{sub_folder}'):
                        os.makedirs(
                            f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}\\{sub_folder}')
                except Exception as e:
                    e

                filename = list_of_file[i]
                target = main_path + '\\' + filename  # target for create link
                linkname = list_of_file[i].replace('.ubf', '.lnk')

                if len(sub_folder) > 0:
                    try:
                        path = f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}\\{sub_folder}\\{linkname}'
                        if sub_folder in filename:
                            create_shortcut(path, target, main_path, target)
                    except Exception as m:
                        m

        if 'thermofix' in folder[0]:  # copy files with wrong name

            list_of_subdirs = os.listdir(
                f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\thermofix\\{sub_folder}')
            try:
                for i in range(len(list_of_subdirs)):
                    os.replace(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\thermofix\\{sub_folder}\\{list_of_subdirs[i]}',
                               f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\termofix\\{sub_folder}\\{list_of_subdirs[i]}')
            except Exception as e:
                e

            try:
                shutil.move(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\thermofix\\{sub_folder}',
                            f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\termofix\\')
            except Exception as p:
                p

        if 'standardhidravlika' in folder[0]:

            list_of_subdirs = os.listdir(
                f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\standardhidravlika\\{sub_folder}')
            try:
                for i in range(len(list_of_subdirs)):
                    os.replace(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\standardhidravlika\\{sub_folder}\\{list_of_subdirs[i]}',
                               f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\standard\\{sub_folder}\\{list_of_subdirs[i]}')
            except Exception as e:
                e

            """ try:
                shutil.move(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\standardhidravlika\\{sub_folder}',
                            f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\standard\\')
            except Exception as p:
                print(p) """

        if 'standarthyrdavlika' in folder[0]:

            list_of_subdirs = os.listdir(
                f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\standarthyrdavlika\\{sub_folder}')
            try:
                for i in range(len(list_of_subdirs)):
                    os.replace(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\standarthyrdavlika\\{sub_folder}\\{list_of_subdirs[i]}',
                               f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\standard\\{sub_folder}\\{list_of_subdirs[i]}')
            except Exception as e:
                e

            """ try:
                shutil.move(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\standarthyrdavlika\\{sub_folder}',
                            f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\standard\\')
            except Exception as p:
                print(p) """

        if 'матадор' in folder[0]:

            list_of_subdirs = os.listdir(
                f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\матадор\\{sub_folder}')
            try:
                for i in range(len(list_of_subdirs)):
                    os.replace(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\матадор\\{sub_folder}\\{list_of_subdirs[i]}',
                               f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\matador\\{sub_folder}\\{list_of_subdirs[i]}')
            except Exception as e:
                e

        sum = sum + 1
        # print('#', end=' ')

    # __RISPA__ sorted
    for i in range(len(list_of_file)):  # create dir from brand names

        filename = list_of_file[i]
        folder = list_of_file[i].split('_', 1)

        if not os.path.exists(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}'):
            os.makedirs(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}')

        sub_folder_rispa = re.findall(r'\_\d{2}\-\d{1}', filename)

        if len(sub_folder_rispa) > 0:
            #print(sub_folder_heaton[0].replace('_', '', 1))
            if sub_folder_rispa[0] in folder[-1]:
                sub_folder_rispa = sub_folder_rispa[0].replace('_', '', 1)

                try:
                    # create subfolders for files
                    if not os.path.exists(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}\\{sub_folder_rispa}'):
                        os.makedirs(
                            f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}\\{sub_folder_rispa}')
                except Exception as e:
                    e

                target = main_path + '\\' + filename  # target for create link
                linkname = list_of_file[i].replace('.ubf', '.lnk')

                if len(sub_folder_rispa) > 0:
                    try:
                        path = f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}\\{sub_folder_rispa}\\{linkname}'

                        if sub_folder_rispa in filename:
                            create_shortcut(path, target, main_path, target)
                    except Exception as m:
                        m

    # __HEATON__ sorted
    for i in range(len(list_of_file)):

        filename = list_of_file[i]
        folder = list_of_file[i].split('_', 1)

        if not os.path.exists(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}'):
            os.makedirs(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}')

        sub_folder_heaton = re.findall(r'\w{2}[^_-]\d{2}\D\d{1}', filename)

        if len(sub_folder_heaton) > 0:

            if sub_folder_heaton[0] in filename:

                sub_folder_heaton = re.sub(
                    r'\_\w{2}', '', sub_folder_heaton[0])
                sub_folder_heaton = re.sub(r'\_', '-', sub_folder_heaton)

                try:
                    # create subfolders for files
                    if not os.path.exists(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}\\{sub_folder_heaton}'):
                        os.makedirs(
                            f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}\\{sub_folder_heaton}')
                except Exception as e:
                    e

                target = main_path + '\\' + filename  # target for create link
                linkname = list_of_file[i].replace('.ubf', '.lnk')
                sub_folder_heaton_link = re.sub(r'\-', '_', sub_folder_heaton)

                if len(sub_folder_heaton) > 0:
                    try:
                        path = f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}\\{sub_folder_heaton}\\{linkname}'

                        if sub_folder_heaton or sub_folder_heaton_link in filename:
                            create_shortcut(path, target, main_path, target)
                    except Exception as m:
                        m

        if 'haeton' in folder[0]:  # copy files with wrong name

            list_of_subdirs = os.listdir(
                f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\haeton\\{sub_folder_heaton}')
            try:
                for i in range(len(list_of_subdirs)):
                    os.replace(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\haeton\\{sub_folder_heaton}\\{list_of_subdirs[i]}',
                               f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\heaton\\{sub_folder_heaton}\\{list_of_subdirs[i]}')
                shutil.rmtree(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\haeton')
            except Exception as e:
                e

    # __WATTSON__ sorted
    for i in range(len(list_of_file)):

        filename = list_of_file[i]
        folder = list_of_file[i].split('_', 1)

        if not os.path.exists(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}'):
            os.makedirs(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}')

        sub_folder_wattson = re.findall(r'\d{1,2}\-\D\-\d{1}', filename)
        sub_folder_wattson_etc = re.findall(
            r'\_\d{2}[^-_0]\D{,1}[^_-x]\d{,1}[^0]', filename)

        if len(sub_folder_wattson) > 0:

            if sub_folder_wattson[0] in filename:

                sub_folder_wattson = re.findall(
                    r'\d{1,2}\-\D\-\d{1}', filename)
                sub_folder_wattson = re.sub(
                    r'\-\D\-', '-', sub_folder_wattson[0])

                try:
                    # create subfolders for files
                    if not os.path.exists(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}\\{sub_folder_wattson}'):
                        os.makedirs(
                            f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}\\{sub_folder_wattson}')
                except Exception as e:
                    e

                target = main_path + '\\' + filename  # target for create link
                linkname = list_of_file[i].replace('.ubf', '.lnk')

                if len(sub_folder_wattson) > 0:
                    try:
                        path = f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}\\{sub_folder_wattson}\\{linkname}'
                        create_shortcut(path, target, main_path, target)
                    except Exception as m:
                        m
                else:
                    print(f'Не знаю как сортировать это: {filename}')

        if len(sub_folder_wattson_etc) > 0:

            if sub_folder_wattson_etc[0] in filename:

                sub_folder_wattson_etc = re.sub(
                    r'\_', '', sub_folder_wattson_etc[0])
                sub_folder_wattson_etc = re.sub(
                    r'\D{1,2}\.', '-', sub_folder_wattson_etc)

                try:
                    # create subfolders for files
                    if not os.path.exists(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}\\{sub_folder_wattson_etc}'):
                        os.makedirs(
                            f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}\\{sub_folder_wattson_etc}')
                except Exception as e:
                    e

                target = main_path + '\\' + filename  # target for create link
                linkname = list_of_file[i].replace('.ubf', '.lnk')

                if len(sub_folder_wattson_etc) > 0:
                    try:
                        path = f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\{folder[0]}\\{sub_folder_wattson_etc}\\{linkname}'
                        create_shortcut(path, target, main_path, target)
                    except Exception as m:
                        m
    loop = loop + 1

try:
    shutil.rmtree(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\thermofix')
    shutil.rmtree(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\haeton')
    shutil.rmtree(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\standardhidravlika')
    shutil.rmtree(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\standarthyrdavlika')
    shutil.rmtree(f'{sorted_folder}\\ШАБЛОНЫ ДЛЯ ПЕЧАТИ\\матадор')
except Exception as e:
    print(e)

time_end = datetime.now()
timer = time_end - time_start


print('\nГотово!')
print('Отсортировано файлов: ' + str(sum/2))
print('Время выполнения: ', ('%d.%d' %
      (timer.seconds, timer.microseconds//1000)), 'сек')
a = input('Нажмите Enter для выхода!')
