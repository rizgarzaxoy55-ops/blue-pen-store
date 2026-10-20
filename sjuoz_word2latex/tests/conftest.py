# -*- coding: utf-8 -*-
"""Ortak test ayarları.

Çalıştırma:
    pytest sjuoz_word2latex/tests                  # bütün testler
    pytest sjuoz_word2latex/tests --update-golden  # beklenen .tex çıktılarını yeniden yaz
"""
import glob
import json
import os
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
FIXTURES = os.path.join(HERE, "fixtures")
GOLDEN = os.path.join(HERE, "golden")

sys.path.insert(0, ROOT)
import sjuoz as conv  # noqa: E402


def pytest_addoption(parser):
    parser.addoption("--update-golden", action="store_true",
                     help="beklenen .tex çıktılarını (tests/golden) mevcut çıktıyla değiştir")


def load_cases():
    cases = []
    for path in sorted(glob.glob(os.path.join(FIXTURES, "*", "case.json"))):
        with open(path, encoding="utf-8") as f:
            case = json.load(f)
        case["name"] = os.path.basename(os.path.dirname(path))
        case["docx"] = sorted(glob.glob(os.path.join(os.path.dirname(path), "*.docx")))
        cases.append(case)
    return cases


CASES = load_cases()


def run_case(case, outdir, compile_pdf=False):
    """Bir test makalesini dönüştürür; convert() sonucunu döndürür."""
    args = list(case["docx"]) + list(case["args"]) + ["-o", str(outdir), "--no-zip"]
    if not compile_pdf:
        args.append("--no-compile")
    opts = conv.build_parser().parse_args(args)
    return conv.convert(opts)


@pytest.fixture
def update_golden(request):
    return request.config.getoption("--update-golden")
