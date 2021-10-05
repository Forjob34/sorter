import os


main_path = input('Укажите путь до каталога: ')
os.chdir(main_path)
sorted_path = main_path.split('\\')
sorted_folder = '\\'.join(sorted_path[:-1])

list_of_file = os.listdir(main_path)

def rename_file(list_of_file):

    for i in range(len(list_of_file)):
        # print(list_of_file[i])
        if 'haeton' in list_of_file[i]:
            new_name = list_of_file[i].replace('haeton', 'heaton')
            name_true = os.rename(list_of_file[i], new_name)

        if 'матадор' in list_of_file[i]:
            new_name = list_of_file[i].replace('матадор', 'matador')
            name_true = os.rename(list_of_file[i], new_name)

        if 'ларко-' in list_of_file[i]:
            new_name = list_of_file[i].replace('ларко-', 'ларко_')
            name_true = os.rename(list_of_file[i], new_name)

        if 'thermofix' in list_of_file[i]:
            new_name = list_of_file[i].replace('thermofix', 'termofix')
            name_true = os.rename(list_of_file[i], new_name)
        
        if 'standardhidravlika' in list_of_file[i]:
            new_name = list_of_file[i].replace('standardhidravlika', 'standart')
            name_true = os.rename(list_of_file[i], new_name)
        
        if 'standarthyrdavlika' in list_of_file[i]:
            new_name = list_of_file[i].replace('standarthyrdavlika', 'standart')
            name_true = os.rename(list_of_file[i], new_name)
        
        if 'standard_hidravlika' in list_of_file[i]:
            try:
                new_name = list_of_file[i].replace('standard_hidravlika', 'standart')
                name_true = os.rename(list_of_file[i], new_name)
            except Exception as e:
                e
                
        if 'standard' in list_of_file[i]:
            try:
                new_name = list_of_file[i].replace('standard', 'standart')
                name_true = os.rename(list_of_file[i], new_name)
            except Exception as e:
                print(e)
                #os.remove(list_of_file[i])

        if 'Standard' in list_of_file[i]:
            new_name = list_of_file[i].replace('Standard', 'standart')
            name_true = os.rename(list_of_file[i], new_name)

    list_of_file = os.listdir(main_path)
    
    return list_of_file

if __name__ == '__main__':
    a = rename_file(list_of_file)
    print(list_of_file)