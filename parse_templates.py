import os
import unicodedata
import re

from resource_path import resource_path
from Export_columns import exp_format
import json

template_folder_name = resource_path(r'# Custom Templates')


def _check_templates_folder():
    """
    Прверяем на наличие папки с шаблонами и дэфолтного файла
    Eсли нет ни того ни другого - создаем
    """
    if not os.path.exists(template_folder_name):
        os.makedirs(template_folder_name)

    default_template_path = os.path.join(template_folder_name, "Default.json")
    if not os.path.exists(default_template_path):
        save_template("Default", exp_format)


def get_templates_folder_path() -> str:
    """
    Возвращаем путь к папке шаблонов
    """

    return template_folder_name


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
    else:
        print("### Info: Default template file deletion is prohibited!")


def save_template(template_name: str, columns_list: list):
    """
    Сохраняем шаблон в JSON
    """
    template_name_valid = slugify(template_name).capitalize()
    template_path = os.path.join(get_templates_folder_path(), template_name_valid + '.json')

    if not os.path.exists(template_path):
        with open(template_path, 'w', encoding='utf8') as file:
            file.write(json.dumps(columns_list, indent=3, ensure_ascii=False))
    else:
        print("### Error: Template already exist with the same name!")


def rewrite_template(template_name: str, columns_list: list):

    template_path = os.path.join(get_templates_folder_path(), template_name + '.json')
    if os.path.exists(template_path):
        open('file.txt', 'w').close()
        with open(template_path, 'w', encoding='utf8') as file:
            file.write(json.dumps(columns_list, indent=3, ensure_ascii=False))


def slugify(value, allow_unicode=True):
    """
    Удаляет недопустимые символы в названии файла
    """
    value = str(value)
    if allow_unicode:
        value = unicodedata.normalize('NFKC', value)
    else:
        value = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore').decode('ascii')
    value = re.sub(r'[^\w\s-]', '', value.lower())
    return re.sub(r'[-\s]+', '-', value).strip('-_')


if __name__ == '__main__':
    z = read_template("Заебись шаблон")
    print(z)
