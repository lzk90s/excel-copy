import time
from typing import Optional

import requests

import util

UPLOAD_IMAGE_URL = 'https://cs8.intsig.net/sync/upload_jpg?last=1&pages=&token={}'
RECOGNIZE_URL = 'https://cs8.intsig.net/sync/cloud_ocr?file_name={}&token={}'


def upload_image(token: str, image_bytes: bytes) -> Optional[str]:
    url = UPLOAD_IMAGE_URL.format(token)
    res = requests.post(url, data=image_bytes)
    if res.status_code != 200:
        print("Upload file failed")
        return None
    data = util.json2dict(res.text)
    doc_id = data["page_id"]
    return str(doc_id)


def recognize(token: str, doc_id: str) -> str:
    url = RECOGNIZE_URL.format(doc_id + ".jpage", token)
    res = requests.get(url)
    if res.status_code != 200:
        print("Cloud ocr failed")
        return None

    data = util.json2dict(res.text)
    ocr_result = util.json2dict(data["cloud_ocr"])
    text = str(ocr_result['ocr_user_text'])
    content = text.encode('latin').decode('utf-8')
    return content


def parse_recognize_result(content: str, item_per_line: int) -> Optional[str]:
    ss = content.replace(" ", "").split("\n")
    if not ss:
        return None

    lines = []
    for i in range(0, len(ss), item_per_line):
        line = ss[i: i + item_per_line]
        lines.append("\t".join(line))

    cc = "\n".join(lines)
    return cc


def ocr(token: str, item_per_line: int, image_bytes: bytes) -> Optional[str]:
    if not image_bytes:
        return None

    doc_id = upload_image(token, image_bytes)
    if not doc_id:
        print("Upload image failed")
        return None

    content = recognize(token, doc_id)
    if not content:
        print("Ocr result is empty")
        return None

    return parse_recognize_result(content, item_per_line)
