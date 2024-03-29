import os
import sys

from win32com.client import Dispatch

__all__ = [
    'printer', 'colored', 'create_shortcut', 'convert_bytes', 'file_to_list', 'list_to_file', 'print_r', 'print_g', 'print_b', 'print_m',
    'print_c', 'print_y', 'print_w', 'print_gr'
]


def printer(data: str):
    """Динамический вывод на одну строку в терминал"""
    sys.stdout.write("\r\x1b[K" + data.__str__())
    sys.stdout.flush()


def colored(r: int, g: int, b: int, text: str):
    """Цветной вывод в терминал"""
    return f"\033[38;2;{r};{g};{b}m{text}\033[0m"


def print_r(text: str):
    """Вывод в терминал красным цветом"""
    print(colored(255, 0, 0, text))


def print_g(text: str):
    """Вывод в терминал зеленым цветом"""
    print(colored(0, 255, 0, text))


def print_b(text: str):
    """Вывод в терминал голубым цветом"""
    print(colored(0, 0, 255, text))


def print_y(text: str):
    """Вывод в терминал желтым цветом"""
    print(colored(255, 255, 0, text))


def print_c(text: str):
    """Вывод в терминал сине-зелёным (cyan) цветом"""
    print(colored(0, 255, 255, text))


def print_m(text: str):
    """Вывод в терминал пурпурным цветом"""
    print(colored(255, 0, 255, text))


def print_w(text: str):
    """Вывод в терминал белым цветом"""
    print(colored(255, 255, 255, text))


def print_gr(text: str):
    """Вывод в терминал серым цветом"""
    print(colored(100, 100, 100, text))


def create_shortcut(path: str, target: str):
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
