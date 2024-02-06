import os
import sys

from win32com.client import Dispatch

__all__ = ['printer', 'colored', 'create_shortcut', 'convert_bytes', 'file_to_list', 'list_to_file']


def printer(data: str):
    """Динамический вывод на одну строку в терминал"""
    sys.stdout.write("\r\x1b[K" + data.__str__())
    sys.stdout.flush()


def colored(r: int, g: int, b: int, text: str):
    """Цветной вывод в терминал"""
    return f"\033[38;2;{r};{g};{b}m{text}\033[0m"


def create_shortcut(path, target):
    """Создать ярлык Windows (*.lnk)"""
    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(path + ".lnk")
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = os.path.dirname(path)
    shortcut.save()


def convert_bytes(num: float, power_of_10=False):
    """Преобразовать байты в Mb... Gb... и т.д."""
    base = 1000.0 if power_of_10 else 1024.0
    for x in ["bytes", "K", "M", "G", "T"]:
        if num < base:
            return f"{num:3.1f}{x}"
        num /= base


def file_to_list(file: str):
    """Загрузить файл в список с пропуском пустых строк."""
    with open(file, "r", encoding="utf-8") as f:
        list = [x.rstrip() for x in f if not x.strip() == ""]
        return list


def list_to_file(file: str, list: list):
    """Записать список в текстовый файл"""
    with open(file, "w", encoding="utf-8") as f:
        for item in list:
            f.write(item + "\n")


if __name__ == "__main__":

    pass
