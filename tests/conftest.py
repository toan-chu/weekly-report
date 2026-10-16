# -*- coding: utf-8 -*-
"""Số liệu giả trong phép thử rất nhỏ (100, 70…): hạ các ngưỡng tính bằng VND cho khỏi lọc mất khách."""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


@pytest.fixture(autouse=True)
def _nguong_nho(monkeypatch):
    from wr import tindung
    monkeypatch.setattr(tindung, "NGUONG_DU_LE", 0.5)


@pytest.fixture(autouse=True)
def _khong_file_master_ngoai():
    """WR-08: file Client Master ngoài sample/ là cài đặt toàn cục — mỗi phép thử bắt đầu không có."""
    from wr import ghep
    ghep.dat_file_master(None)
    yield
    ghep.dat_file_master(None)
