import configparser
import os


# https://docs.python.org/3/library/configparser.html

class CFG(object):
    def __init__(self, cfg_name):
        self.cfg = configparser.ConfigParser()
        self.cfg_name = cfg_name
        self.cfg_default_path = os.path.expanduser(r'~\Documents')
        self.cfg_project_path = os.path.join(self.cfg_default_path, self.cfg_name)
        self.cfg_path = os.path.join(self.cfg_project_path, self.cfg_name + '.ini')

        self._check_folder()

    def _is_bool(self, value):
        if str(value).lower() == 'true':
            return True
        if str(value).lower() == 'false':
            return False
        return None

    def _check_folder(self):
        """
        Прверяем на наличие папки и конфига
        Eсли нет ни того ни другого - создаем
        """
        if not os.path.exists(self.cfg_project_path):
            os.makedirs(self.cfg_project_path)
            print(f"### Info: Poject CFG folder created: {self.cfg_project_path}")

        if not os.path.exists(self.cfg_path):
            with open(self.cfg_path, 'w', encoding='utf8') as configfile:
                self.cfg.write(configfile)

    def _check_section(self, section):
        if section not in self.cfg.sections():
            self.cfg[section] = {}
            return False
        return True

    def _check_key(self, section, key):
        self._check_section(section)
        if key not in list(self.cfg[section].keys()):
            self.cfg[section][key] = ''
            return False
        return True

    def store_settings(self, section, key, value):

        self.cfg.read(self.cfg_path)
        self._check_key(section=section, key=key)

        self.cfg[section][key] = str(value)
        with open(self.cfg_path, 'w', encoding='utf8') as configfile:
            self.cfg.write(configfile)

    def read_cfg(self, section, key, create_if_none=False, default=''):

        self.cfg.read(self.cfg_path)
        if self._check_key(section=section, key=key) is True:
            value = self.cfg[section][key]
            # проверка на Bool
            if self._is_bool(value) is not None:
                return self._is_bool(value)
            # проверка на Int
            elif str(value).isdigit():
                if float(value).is_integer():
                    return int(value)
            else:
                return str(value)
        else:
            if create_if_none is True or default != '':
                self.store_settings(section=section, key=key, value=default)
            return None


if __name__ == '__main__':
    cfg1 = CFG("PY_project")
    indext_path = cfg1.read_cfg(section="PATHS", key="INDEX_CUSTOM", create_if_none=True)
    struct_path = cfg1.read_cfg(section="PATHS", key="STRUCT_CUSTOM", create_if_none=True)

    print(indext_path)
    print(struct_path)
    # cfg1.store_settings(section="PATHS", key="INDEX_CUSTOM", value=20)
    # cfg1.store_settings(section="PATHS", key="STRUCT_CUSTOM", value=20)
