import os
import unicodedata
import re
import subprocess
from tkinter import messagebox

from Export_columns import exp_format
import json

template_folder_name = os.path.expanduser(r'~\Documents\DB Process\# Custom Templates')
# template_folder_name = resource_path(r'# Custom Templates')


def _check_templates_folder():
    """
    Прверяем на наличие папки с шаблонами и дэфолтного файла
    Eсли нет ни того ни другого - создаем
    """
    if not os.path.exists(template_folder_name):
        os.makedirs(template_folder_name)
        print(f"### Info: Default templates folder created: {template_folder_name}")

    default_template_path = os.path.join(template_folder_name, "Default.json")
    if not os.path.exists(default_template_path):
        save_template("Default", exp_format)


def get_templates_folder_path() -> str:
    """
    Возвращаем путь к папке шаблонов
    """

    return template_folder_name


def open_templates_folder():
    subprocess.Popen(fr'explorer /select,{template_folder_name}')


def get_templates_list() -> list:
    """
    Возвращаем список файлов в папке шаблонов без расширений
    """
    _check_templates_folder()
    templates_list = []
    for filename in os.listdir(get_templates_folder_path()):
        if filename.endswith('.json'):
            filename_extless = filename[:-5]
            templates_list.append(filename_extless)

    return templates_list


def read_template(template_name: str) -> list:
    template_path = os.path.join(get_templates_folder_path(), template_name + '.json')

    if os.path.exists(template_path):
        with open(template_path, encoding='utf-8') as json_file:
            template_columns = json.load(json_file)
            return template_columns
    else:
        messagebox.showwarning(f'Read template Warning', f"<{template_name}> template is not exist!!",
                               icon='warning')
        print("### Error: Template file is not exist!")
        return []


def delete_template(template_name: str):
    """
    Удаляем шаблон
    """
    if template_name.lower() != 'default':
        template_path = os.path.join(get_templates_folder_path(), template_name + '.json')

        if os.path.exists(template_path):
            os.remove(template_path)
        messagebox.showwarning(f'Delete template Info', f"<{template_name}> deleted.",
                               icon='info')
    else:
        messagebox.showwarning(f'Delete template Info', f"You can't delete <{template_name}> template!", icon='info')
        print("### Info: Default template file deletion is prohibited!")


def save_template(template_name: str, columns_list: list):
    """
    Сохраняем шаблон в JSON
    """
    template_name_valid = slugify(template_name)
    template_path = os.path.join(get_templates_folder_path(), template_name_valid + '.json')

    if not os.path.exists(template_path):
        with open(template_path, 'w', encoding='utf8') as file:
            file.write(json.dumps(columns_list, indent=3, ensure_ascii=False))
            messagebox.showwarning(f'Save template Info', f"<{template_name}> template Saved",
                                   icon='info')
            print("### Info: Template Saved!")
    else:
        rewrite = messagebox.askquestion(f'Save template Warning',
                                         f"<{template_name}> template already exists!!\n\t"
                                         f"Rewrite?", icon='warning')

        if rewrite == 'yes':
            rewrite_template(template_name=template_name, columns_list=columns_list)
            # messagebox.showwarning(f'Save template Warning', f"<{template_name}> template already exists!!", icon='warning')
            print(f"### Error: <{template_name}> template already exists!! Rewrite?")

def rewrite_template(template_name: str, columns_list: list):

    template_path = os.path.join(get_templates_folder_path(), template_name + '.json')
    if os.path.exists(template_path):
        # open('file.txt', 'w').close()
        with open(template_path, 'w', encoding='utf8') as file:
            file.write(json.dumps(columns_list, indent=3, ensure_ascii=False))
            messagebox.showwarning(f'Rewrite template Info', f"<{template_name}> template Rewrited",
                                   icon='info')
            print("### Info: Template Rewrited!")


def slugify(value, allow_unicode=True):
    """
    Удаляет недопустимые символы в названии файла
    """
    value = str(value)
    if allow_unicode:
        value = unicodedata.normalize('NFKC', value)
    else:
        value = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore').decode('ascii')

    value = re.sub(r'[^\w\s-]', '', value)
    return value


if __name__ == '__main__':
    z = "Ямал FR"
    print(slugify(z))
