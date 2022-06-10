import cloud_ocr
import config
import timer
import util


def fun_timer(c: config.OcrConfig) -> None:
    img = util.read_image_from_clipboard()
    result = cloud_ocr.ocr(c.token, c.item_per_line, img)
    if result:
        print("Ocr recognize succeed, copy to clipboard")
        util.clip_copy(result)


if __name__ == '__main__':
    util.run_once()
    cfg = config.parse_config()
    print("Start image ocr")
    t = timer.UseTimer(1, fun_timer, {cfg})
    t.timer_start()
