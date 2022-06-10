import configparser
import os


class OcrConfig:
    token = ""
    item_per_line = 1


def generate_default_config(cfg: str) -> None:
    print("Generate default config")
    with open(cfg, 'w') as fp:
        cf = configparser.ConfigParser()
        cf.add_section("account")
        cf.set("account", "token", "")
        cf.write(fp)


def parse_config(cfg: str = os.environ['HOME'] + os.sep + "ocr.ini") -> OcrConfig:
    if not os.path.exists(cfg):
        generate_default_config(cfg)

    cf = configparser.ConfigParser()
    cf.read(cfg)

    config = OcrConfig()
    config.token = cf.get("account", "token")
    config.item_per_line = int(cf.get("ocr", "item_per_line"))
    assert config.token
    assert config.item_per_line
    return config
