import logging
import logging.handlers
import os

# DEBUG, INFO, WARNING, ERROR, CRITICAL


def init_logger(name, log_pth, print_console=False):
    logger = logging.getLogger(name)
    FORMAT = '%(asctime)s - %(name)s:%(lineno)s - %(levelname)s - %(message)s'
    datefmt = '%Y.%m.%d %H:%M:%S'
    logger.setLevel(logging.DEBUG)
    # хэндлер в консоль
    if print_console is True:
        sh = logging.StreamHandler()
        sh.setFormatter(logging.Formatter(FORMAT, datefmt))
        sh.setLevel(logging.DEBUG)
        logger.addHandler(sh)
        # sh.addFilter(disable_debug)
    # хэндлер записи в файл
    fh = logging.handlers.RotatingFileHandler(filename=log_pth)
    fh.setFormatter(logging.Formatter(FORMAT, datefmt))
    fh.setLevel(logging.DEBUG)

    logger.addHandler(fh)
    usr, pc_name = get_user_sats()
    #logger.debug(f'Logger initialized by Username: {usr}, PC name: {pc_name}')


def disable_debug(log: logging.LogRecord) -> int:
    if 'DEBUG' in log.levelname:
        return 0
    return 1


def get_user_sats():
    """
    :return: tuple (username, pc_name)
    """
    username = os.getenv('username')
    pc_name = os.environ['COMPUTERNAME']
    return username, pc_name