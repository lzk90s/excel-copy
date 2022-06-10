import fcntl
import json
import numbers
import os
import re
import sys
import tempfile
from typing import Optional

import PIL.ImageGrab as ImageGrab
import pyperclip


def obj2json(obj: object, indent=4) -> str:
    return json.dumps(obj, default=lambda o: o.__dict__, sort_keys=True, indent=indent, ensure_ascii=False)


def json2dict(json_str: str) -> dict:
    return json.loads(json_str)


def is_chinese(string: str) -> bool:
    for ch in string:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True
    return False


def get_number_from_str(desc: str):
    pattern = re.compile(r'\d+')
    res = re.findall(pattern, desc)
    res = list(map(int, res))
    res.sort(reverse=True)
    if res:
        return res[0]
    return None


def is_excel_serial_no(txt, max=100) -> bool:
    """
    判断是否是excel的序号，序号暂时最大到100
    :param txt:
    :return:
    """
    num = get_number_from_str(txt)
    if not num:
        return False
    return int(num) < max


def read_binary_file(path: str) -> bytes:
    with open(path, 'rb') as f:
        return f.read()


def read_image_from_clipboard() -> Optional[bytes]:
    file = '/tmp/snapshot.png'
    img = ImageGrab.grabclipboard()
    if not img:
        return None
    img.save(file, 'PNG')
    return read_binary_file(file)


def clip_copy(content: str) -> None:
    if not content:
        print("No content to copy")
        return
    pyperclip.copy(content)


__fh = 0


def run_once(lock_file=None):
    if not lock_file:
        basename = os.path.splitext(os.path.abspath(sys.argv[0]))[0].replace(
            "/", "-").replace(":", "").replace("\\", "-") + '.lock'
        lock_file = os.path.normpath(tempfile.gettempdir() + '/' + basename)

    global __fh
    f = os.path.abspath(lock_file)
    # print("lock file " + f)

    if not os.path.exists(f):
        with open(f, 'w') as fw:
            fw.close()

    __fh = open(f, 'r')
    try:
        fcntl.flock(__fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except:
        print("Only allow run single instance")
        os._exit(0)
