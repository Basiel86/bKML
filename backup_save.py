import collections
import os
import pathlib
from datetime import datetime
import zipfile


class BackupFile:
    def __init__(self, file_path: str, mode: str, total_backups=10):
        """
        Any file backup
        :param file_path: as it is
        :param mode: 'auto' or 'manual' - auto will keep total_backups count
        :param total_backups: count
        """

        self.modes = 'auto', 'manual'
        if mode in self.modes:
            self.mode = mode
        else:
            self.mode = 'manual'

        self.backup_folder_name = 'Backup'
        self.file_path = file_path
        self.backup_folder_dir_path = os.path.dirname(file_path)

        self.total_backups = total_backups
        self.backups_list = []

    @staticmethod
    def stack_put(value: str, lst: list) -> (list, str):
        last_element = lst[len(lst) - 1]
        lst_shift = collections.deque(lst)
        lst_shift.rotate(1)
        lst_shift[0] = value
        return list(lst_shift), last_element

    def get_backup_name_time(self):
        """
        Return Backup name with Time
        """
        now = datetime.now()
        dt_string = now.strftime("%Y.%m.%d %H.%M.%S")
        return self.get_base_name() + f" {self.mode} {dt_string}"

    def check_backup_folder(self):
        """
        Check and create backup folder
        """
        backup_folder_path = os.path.join(os.path.dirname(self.file_path), self.backup_folder_name)
        if not os.path.exists(backup_folder_path):
            os.mkdir(backup_folder_path)

        file_ext = pathlib.Path(self.file_path).suffix
        backup_file_path = backup_folder_path + rf"\{os.path.basename(self.file_path)}"
        if not os.path.exists(backup_file_path):
            os.mkdir(backup_file_path)
        else:
            if len(self.backups_list) == 0:
                for filename in os.listdir(backup_file_path):
                    if 'auto' in filename:
                        self.backups_list.append(filename)
                self.backups_list.sort(reverse=True)

    def get_base_name(self) -> (str, None):
        """
        Return file base name without extension
        :return: Str or None
        """
        basename = os.path.basename(self.file_path)
        if os.path.isfile(self.file_path):
            file_ext = pathlib.Path(self.file_path).suffix
            return basename[:len(basename) - len(file_ext)]
        return None

    def total_backups_process(self, backup_name, backup_folder_dir_path):
        """
        append new element and if more than max value of backups - remove
        :param backup_name: new backup name
        :param backup_folder_dir_path: full path of folder file exists
        """
        if len(self.backups_list) < self.total_backups:
            self.backups_list.append(backup_name)
        else:
            self.backups_list, last_element = self.stack_put(value=backup_name, lst=self.backups_list)
            backup_path = os.path.join(backup_folder_dir_path, last_element)
            os.remove(backup_path)

    def save_backup(self, print_status=False):
        """
        Save backup, if mode is auto - remove oldest file
        :param print_status: print or not status
        """
        if os.path.exists(self.file_path):
            self.check_backup_folder()
            backup_folder_dir_path = os.path.join(self.backup_folder_dir_path,
                                                  self.backup_folder_name,
                                                  os.path.basename(self.file_path))
            backup_name = self.get_backup_name_time() + '.zip'
            backup_path = os.path.join(backup_folder_dir_path, backup_name)

            if self.mode == 'auto':
                self.total_backups_process(backup_name=backup_name, backup_folder_dir_path=backup_folder_dir_path)

            # название и путь Zip файла
            with zipfile.ZipFile(backup_path, 'w',
                                 compression=zipfile.ZIP_DEFLATED,
                                 compresslevel=9) as zf:
                # добавление файлов построчно в архив (можно много)
                zf.write(self.file_path, arcname=os.path.basename(self.file_path))

            if print_status:
                print(f'# Backup saved: {backup_name}')
        else:
            print('# ERROR: Path is wrong!!!')


if __name__ == '__main__':
    from queue import LifoQueue

    path = r"c:\Users\Vasily\OneDrive\Macro\PYTHON\bKML\file.txt"

    z = BackupFile(file_path=path, mode='auto')
    z.save_backup(print_status=True)
