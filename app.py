import os
import re
import math
import subprocess
from datetime import datetime
from typing import List, Dict, Any, Optional

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import ezdxf
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(title="Martech Professional BOM Engine")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# === CẤU HÌNH ĐƯỜNG DẪN ===
import sys
if sys.platform == "linux":
    ODA_EXE_PATH = "/usr/bin/ODAFileConverter"
else:
    ODA_EXE_PATH = r"D:\Program Files\ODA\ODAFileConverter 27.1.0\ODAFileConverter.exe"

os.makedirs("temp_dwg_in", exist_ok=True)
os.makedirs("temp_dxf_out", exist_ok=True)
os.makedirs("exports", exist_ok=True)


# --- 1. LÀM SẠCH KÝ TỰ RÁC AUTOCAD MTEXT ---
def clean_cad_mtext(text: str) -> str:
    if not text:
        return ""
    text = text.replace("%%c", "Ø").replace("%%C", "Ø")
    text = text.replace("%%d", "°").replace("%%D", "°")
    text = text.replace("%%p", "±").replace("%%P", "±")
    text = re.sub(r"\\P", " / ", text, flags=re.IGNORECASE)
    text = re.sub(r"/ xq[cl][^;]*;?", "", text)
    text = re.sub(r"\\pxq[cl];?", "", text)
    text = re.sub(r"\\[A-Za-z0-9]+;", "", text)
    text = re.sub(r"\\f[^;]+;", "", text)
    text = re.sub(r"\\[LlOoKk]", "", text)
    text = re.sub(r"\\[HhWwQqCcTtFfAa][^;]*;", "", text)
    text = re.sub(r"\\[HhWwQqCcTtFfAa][0-9.]+x?", "", text)
    text = text.replace("{", "").replace("}", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip(" /;:,")


# --- 2. TỪ ĐIỂN QUY CÁCH TIÊU CHUẨN MARTECH (ROTARY 700 & BỒN NƯỚC TP2887) ---
MARTECH_SPEC_DATABASE = {
    # Rotary 700
    "1-P1": {"TenVT": "Thép tấm SS400", "MoTa": "Tôn 16mm trước tiện", "Dai": 805, "Rong": 805, "Cao": 16, "Dvt": "Tấm", "KL": 62.7, "DT": 1.02},
    "1-P2": {"TenVT": "Thép tấm SS400", "MoTa": "SS400 20mm", "Dai": 219, "Rong": 219, "Cao": 20, "Dvt": "Tấm", "KL": 5.8, "DT": 0.10},
    "1-P3": {"TenVT": "Thép tấm SS400", "MoTa": "SS400 20mm", "Dai": 219, "Rong": 219, "Cao": 20, "Dvt": "Tấm", "KL": 4.6, "DT": 0.10},
    "1-O1": {"TenVT": "Thép ống đúc", "MoTa": "Ống Ø219 (Khoan lỗ bơm mỡ)", "Dai": 135, "Rong": 219, "Cao": 12.7, "Dvt": "Cái", "KL": 8.7, "DT": 0.18},
    "1-O1_TRUC": {"TenVT": "Thép láp S45C", "MoTa": "Láp C45 Ø110x1730mm", "Dai": 1730, "Rong": 110, "Cao": 110, "Dvt": "Cây", "KL": 128.9, "DT": 0.60},
    "1-O2": {"TenVT": "Thép láp S45C", "MoTa": "Láp Ø195 (Khoan lỗ sau tiện)", "Dai": 95, "Rong": 195, "Cao": 195, "Dvt": "Cái", "KL": 22.2, "DT": 0.12},
    "2-P1": {"TenVT": "Thép tấm SS400", "MoTa": "SS400 8mm", "Dai": 800, "Rong": 291, "Cao": 8, "Dvt": "Tấm", "KL": 8.7, "DT": 0.46},
    "2-P2": {"TenVT": "Thép tấm SS400", "MoTa": "SS400 8mm", "Dai": 1000, "Rong": 310, "Cao": 8, "Dvt": "Tấm", "KL": 19.5, "DT": 0.62},
    "2-P3": {"TenVT": "Thép tấm SS400", "MoTa": "SS400 8mm", "Dai": 900, "Rong": 302, "Cao": 8, "Dvt": "Tấm", "KL": 10.2, "DT": 0.54},
    "2-P4": {"TenVT": "Thép tấm SS400", "MoTa": "SS400 8mm", "Dai": 1000, "Rong": 335, "Cao": 8, "Dvt": "Tấm", "KL": 21.0, "DT": 0.67},
    "2-P5": {"TenVT": "Thép tấm SS400", "MoTa": "SS400 16mm trước tiện tấm", "Dai": 805, "Rong": 805, "Cao": 16, "Dvt": "Tấm", "KL": 16.3, "DT": 1.02},
    "2-P6": {"TenVT": "Thép tấm SS400", "MoTa": "SS400 12mm (vỏ cuốn)", "Dai": 2160, "Rong": 1090, "Cao": 12, "Dvt": "Tấm", "KL": 221.8, "DT": 4.71},
    "3-P1": {"TenVT": "Thép tấm SS400", "MoTa": "SS400 10mm", "Dai": 636, "Rong": 636, "Cao": 10, "Dvt": "Tấm", "KL": 24.2, "DT": 0.81},
    "3-P2": {"TenVT": "Thép tấm SS400", "MoTa": "SS400 10mm (cánh quay)", "Dai": 1010, "Rong": 209.5, "Cao": 10, "Dvt": "Tấm", "KL": 14.0, "DT": 0.42},
    "NẮP CHẶN MOTOR": {"TenVT": "Thép tấm SS400", "MoTa": "Nắp chặn Motor 10mm", "Dai": 120, "Rong": 120, "Cao": 10, "Dvt": "Tấm", "KL": 0.9, "DT": 0.03},
    "LA 10X50": {"TenVT": "Thép thanh La", "MoTa": "La 10x50mm viền mặt bích", "Dai": 1000, "Rong": 50, "Cao": 10, "Dvt": "Thanh", "KL": 3.9, "DT": 0.12},
    "LA 8X50": {"TenVT": "Thép thanh La", "MoTa": "La 8x50mm viền mặt bích", "Dai": 1000, "Rong": 50, "Cao": 8, "Dvt": "Thanh", "KL": 3.1, "DT": 0.12},

    # Bồn nước 20m3 & Khử khí TP2887
    "P2": {"TenVT": "Thép tấm SS400", "MoTa": "Tôn thân bồn nước ID1970x7092x5mm (4 khoang lốc cuộn)", "Dai": 1773, "Rong": 6189, "Cao": 5, "Dvt": "Tấm", "KL": 430.7, "DT": 21.95},
    "P1.1": {"TenVT": "Chỏm cầu Ellipsoidal", "MoTa": "Chỏm cầu bồn nước ID1970x6mm (bán kính R=1773/r=335)", "Dai": 1970, "Rong": 1970, "Cao": 6, "Dvt": "Cái", "KL": 198.3, "DT": 8.42},
    "P4": {"TenVT": "Thép tấm SS400", "MoTa": "Tôn thân bình khử khí ID1200x2209x5mm", "Dai": 2209, "Rong": 3770, "Cao": 5, "Dvt": "Tấm", "KL": 326.9, "DT": 16.66},
    "P3": {"TenVT": "Chỏm cầu Ellipsoidal", "MoTa": "Chỏm cầu bình khử khí ID1200x6mm", "Dai": 1200, "Rong": 1200, "Cao": 6, "Dvt": "Cái", "KL": 73.5, "DT": 3.12},
    "ELD-ID1970": {"TenVT": "Chỏm cầu Ellipsoidal", "MoTa": "Chỏm cầu bồn nước ID1970x6mm", "Dai": 1970, "Rong": 1970, "Cao": 6, "Dvt": "Cái", "KL": 198.3, "DT": 8.42},
    "ELD-ID1200": {"TenVT": "Chỏm cầu Ellipsoidal", "MoTa": "Chỏm cầu bình khử khí ID1200x6mm", "Dai": 1200, "Rong": 1200, "Cao": 6, "Dvt": "Cái", "KL": 73.5, "DT": 3.12},
    "P5": {"TenVT": "Thép tấm SS400", "MoTa": "Yên đỡ bồn nước (Saddle support)", "Dai": 1200, "Rong": 450, "Cao": 16, "Dvt": "Bộ", "KL": 98.0, "DT": 1.88},
    "P5.1": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm đế yên đỡ bồn t=10mm", "Dai": 1200, "Rong": 350, "Cao": 10, "Dvt": "Tấm", "KL": 33.0, "DT": 0.84},
    "P5.2": {"TenVT": "Thép tấm A36", "MoTa": "Tấm vách đứng yên đỡ bồn t=16mm", "Dai": 1150, "Rong": 450, "Cao": 16, "Dvt": "Tấm", "KL": 65.0, "DT": 1.04},
    "P6": {"TenVT": "Thép tấm SS400", "MoTa": "Cụm gối đỡ phụ chân bồn", "Dai": 450, "Rong": 300, "Cao": 8, "Dvt": "Bộ", "KL": 21.0, "DT": 0.45},
    "P6.1": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm vách gối đỡ bồn t=8mm", "Dai": 450, "Rong": 280, "Cao": 8, "Dvt": "Tấm", "KL": 7.9, "DT": 0.25},
    "P6.2": {"TenVT": "Thép tấm SS400", "MoTa": "Gân tăng cứng gối đỡ t=6mm", "Dai": 350, "Rong": 200, "Cao": 6, "Dvt": "Tấm", "KL": 3.3, "DT": 0.14},
    "P6.3": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm đệm gối đỡ t=4mm", "Dai": 300, "Rong": 180, "Cao": 4, "Dvt": "Tấm", "KL": 1.7, "DT": 0.11},
    "P6.4": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm gân liên kết gối đỡ t=6mm", "Dai": 280, "Rong": 150, "Cao": 6, "Dvt": "Tấm", "KL": 2.0, "DT": 0.08},
    "P6.5": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm giằng gối đỡ t=6mm", "Dai": 260, "Rong": 140, "Cao": 6, "Dvt": "Tấm", "KL": 1.7, "DT": 0.07},
    "P6.6": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm bịt đầu gối đỡ t=6mm", "Dai": 250, "Rong": 120, "Cao": 6, "Dvt": "Tấm", "KL": 1.4, "DT": 0.06},
    "P7": {"TenVT": "Thép tấm SS400", "MoTa": "Cụm gân vành tăng cứng chân bồn", "Dai": 500, "Rong": 250, "Cao": 10, "Dvt": "Bộ", "KL": 28.5, "DT": 0.52},
    "P7.1": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm vành tăng cứng chân bồn t=10mm", "Dai": 450, "Rong": 220, "Cao": 10, "Dvt": "Tấm", "KL": 7.8, "DT": 0.20},
    "P7.2": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm sườn tăng cứng chân bồn t=10mm", "Dai": 400, "Rong": 200, "Cao": 10, "Dvt": "Tấm", "KL": 6.3, "DT": 0.16},
    "P7.3": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm lót đệm chân bồn t=4mm", "Dai": 350, "Rong": 180, "Cao": 4, "Dvt": "Tấm", "KL": 2.0, "DT": 0.13},
    "P7.4": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm giằng chéo chân bồn t=6mm", "Dai": 320, "Rong": 160, "Cao": 6, "Dvt": "Tấm", "KL": 2.4, "DT": 0.10},
    "P7.5": {"TenVT": "Thép tấm SS400", "MoTa": "Gân tăng cường đế bồn t=6mm", "Dai": 300, "Rong": 150, "Cao": 6, "Dvt": "Tấm", "KL": 2.1, "DT": 0.09},
    "P7.6": {"TenVT": "Thép tấm SS400", "MoTa": "Gân chặn chân bồn t=6mm", "Dai": 280, "Rong": 140, "Cao": 6, "Dvt": "Tấm", "KL": 1.8, "DT": 0.08},
    "P7.7": {"TenVT": "Thép tấm SS400", "MoTa": "Gân khóa chân bồn t=6mm", "Dai": 260, "Rong": 120, "Cao": 6, "Dvt": "Tấm", "KL": 1.5, "DT": 0.06},
    "P8": {"TenVT": "Thép tấm SS400", "MoTa": "Cụm tai cẩu nâng bồn nước (Lifting lug)", "Dai": 450, "Rong": 320, "Cao": 6, "Dvt": "Bộ", "KL": 11.5, "DT": 0.49},
    "P8.1": {"TenVT": "Thép tấm SS400", "MoTa": "Bản tai cẩu bồn nước t=6mm", "Dai": 380, "Rong": 260, "Cao": 6, "Dvt": "Tấm", "KL": 4.7, "DT": 0.20},
    "P8.2": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm đệm tai cẩu bồn nước t=6mm", "Dai": 450, "Rong": 320, "Cao": 6, "Dvt": "Tấm", "KL": 6.8, "DT": 0.29},
    "P9": {"TenVT": "Thép tấm SS400", "MoTa": "Cụm tai cẩu bình khử khí", "Dai": 380, "Rong": 280, "Cao": 6, "Dvt": "Bộ", "KL": 8.3, "DT": 0.35},
    "P9.1": {"TenVT": "Thép tấm SS400", "MoTa": "Bản tai cẩu bình khử khí t=6mm", "Dai": 320, "Rong": 220, "Cao": 6, "Dvt": "Tấm", "KL": 3.3, "DT": 0.14},
    "P9.2": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm đệm tai cẩu bình khử khí t=6mm", "Dai": 380, "Rong": 280, "Cao": 6, "Dvt": "Tấm", "KL": 5.0, "DT": 0.21},
    "P10": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm lót đáy chân bồn t=4mm", "Dai": 1100, "Rong": 300, "Cao": 4, "Dvt": "Tấm", "KL": 10.4, "DT": 0.66},
    "P11.1": {"TenVT": "Thép tấm SS400", "MoTa": "Bản mã liên kết dầm chân H t=12mm", "Dai": 280, "Rong": 240, "Cao": 12, "Dvt": "Tấm", "KL": 6.3, "DT": 0.13},
    "P11.3": {"TenVT": "Thép tấm SS400", "MoTa": "Gân tăng cứng dầm chân H t=12mm", "Dai": 250, "Rong": 200, "Cao": 12, "Dvt": "Tấm", "KL": 4.7, "DT": 0.10},
    "P12.1": {"TenVT": "Thép tấm SS400", "MoTa": "Bản mã giằng chân bồn t=6mm", "Dai": 320, "Rong": 220, "Cao": 6, "Dvt": "Tấm", "KL": 3.3, "DT": 0.14},
    "P12.2": {"TenVT": "Thép tấm SS400", "MoTa": "Bản mã đệm bulông móng t=12mm", "Dai": 350, "Rong": 250, "Cao": 12, "Dvt": "Tấm", "KL": 8.2, "DT": 0.18},
    "P13.1": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm đệm tăng cường Nozzle t=6mm", "Dai": 300, "Rong": 300, "Cao": 6, "Dvt": "Tấm", "KL": 4.2, "DT": 0.18},
    "P13.2": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm đệm cổ ống Nozzle t=6mm", "Dai": 280, "Rong": 280, "Cao": 6, "Dvt": "Tấm", "KL": 3.7, "DT": 0.16},
    "P14.1": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm bản mã giá đỡ phụ t=6mm", "Dai": 260, "Rong": 200, "Cao": 6, "Dvt": "Tấm", "KL": 2.5, "DT": 0.10},
    "P14.2": {"TenVT": "Thép tấm SS400", "MoTa": "Gân tăng cứng giá đỡ phụ t=6mm", "Dai": 240, "Rong": 180, "Cao": 6, "Dvt": "Tấm", "KL": 2.0, "DT": 0.09},
    "P14.3": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm đệm liên kết giá đỡ t=6mm", "Dai": 220, "Rong": 160, "Cao": 6, "Dvt": "Tấm", "KL": 1.7, "DT": 0.07},
    "P15.1": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm giá đỡ ống xả đáy t=6mm", "Dai": 250, "Rong": 180, "Cao": 6, "Dvt": "Tấm", "KL": 2.1, "DT": 0.09},
    "P15.2": {"TenVT": "Thép tấm SS400", "MoTa": "Gân đỡ ống xả đáy t=6mm", "Dai": 220, "Rong": 150, "Cao": 6, "Dvt": "Tấm", "KL": 1.6, "DT": 0.07},
    "P15.3": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm kẹp ống xả đáy t=6mm", "Dai": 200, "Rong": 120, "Cao": 6, "Dvt": "Tấm", "KL": 1.1, "DT": 0.05},
    "P16.1": {"TenVT": "Thép tấm SS400", "MoTa": "Bản mã chân thang leo t=6mm", "Dai": 220, "Rong": 150, "Cao": 6, "Dvt": "Tấm", "KL": 1.6, "DT": 0.07},
    "P16.2": {"TenVT": "Thép tấm SS400", "MoTa": "Bản mã đệm chân thang leo t=10mm", "Dai": 250, "Rong": 180, "Cao": 10, "Dvt": "Tấm", "KL": 3.5, "DT": 0.09},
    "P17": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm chắn bảo vệ mức nước t=6mm", "Dai": 450, "Rong": 200, "Cao": 6, "Dvt": "Tấm", "KL": 4.2, "DT": 0.18},
    "P18.1": {"TenVT": "Thép tấm Inox 304", "MoTa": "Vách ngăn dòng bình khử khí Inox 304 t=6mm", "Dai": 850, "Rong": 500, "Cao": 6, "Dvt": "Tấm", "KL": 20.2, "DT": 0.85},
    "P18.2": {"TenVT": "Thép tấm Inox 304", "MoTa": "Tấm hướng dòng bình khử khí Inox 304 t=6mm", "Dai": 600, "Rong": 400, "Cao": 6, "Dvt": "Tấm", "KL": 11.4, "DT": 0.48},
    "P18.3": {"TenVT": "Thép tấm Inox 304", "MoTa": "Bản mã liên kết khay khử khí Inox t=16mm", "Dai": 350, "Rong": 250, "Cao": 16, "Dvt": "Tấm", "KL": 11.1, "DT": 0.18},
    "P18.4": {"TenVT": "Thép tấm Inox 304", "MoTa": "Bản mã gối đỡ khay khử khí Inox t=16mm", "Dai": 400, "Rong": 300, "Cao": 16, "Dvt": "Tấm", "KL": 15.2, "DT": 0.24},
    "P18.5": {"TenVT": "Thép tấm Inox 304", "MoTa": "Tấm đệm gối đỡ khay khử khí Inox t=16mm", "Dai": 380, "Rong": 280, "Cao": 16, "Dvt": "Tấm", "KL": 13.5, "DT": 0.21},
    "P19.1": {"TenVT": "Thép tấm Inox 304", "MoTa": "Tấm chắn buồng khử khí Inox t=3mm", "Dai": 500, "Rong": 350, "Cao": 3, "Dvt": "Tấm", "KL": 4.2, "DT": 0.35},
    "P19.2": {"TenVT": "Thép tấm Inox 304", "MoTa": "Tấm phân phối hơi buồng khử khí Inox t=3mm", "Dai": 480, "Rong": 320, "Cao": 3, "Dvt": "Tấm", "KL": 3.7, "DT": 0.31},
    "P19.3": {"TenVT": "Thép tấm Inox 304", "MoTa": "Vách ngăn phân dòng buồng khử khí Inox t=3mm", "Dai": 450, "Rong": 300, "Cao": 3, "Dvt": "Tấm", "KL": 3.2, "DT": 0.27},
    "P19.4": {"TenVT": "Thép tấm Inox 304", "MoTa": "Tấm dẫn hướng buồng khử khí Inox t=3mm", "Dai": 420, "Rong": 280, "Cao": 3, "Dvt": "Tấm", "KL": 2.8, "DT": 0.24},
    "P19.5": {"TenVT": "Thép tấm Inox 304", "MoTa": "Tấm hướng dòng buồng khử khí Inox t=3mm", "Dai": 400, "Rong": 260, "Cao": 3, "Dvt": "Tấm", "KL": 2.5, "DT": 0.21},
    "P20": {"TenVT": "Thép tấm Inox 304", "MoTa": "Thanh đỡ khay khử khí Inox 304 t=10mm", "Dai": 1150, "Rong": 100, "Cao": 10, "Dvt": "Tấm", "KL": 9.1, "DT": 0.23},
    "P21.1": {"TenVT": "Thép tấm Inox 304", "MoTa": "Khay phân phối nước dập lỗ Inox 304 t=2mm", "Dai": 1100, "Rong": 150, "Cao": 2, "Dvt": "Tấm", "KL": 2.6, "DT": 0.33},
    "P21.2": {"TenVT": "Thép tấm Inox 304", "MoTa": "Khay khử khí dập lỗ Inox 304 t=2mm", "Dai": 1100, "Rong": 150, "Cao": 2, "Dvt": "Tấm", "KL": 2.6, "DT": 0.33},
    "P22": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm gia cường Nozzle N11 bồn nước t=6mm", "Dai": 350, "Rong": 350, "Cao": 6, "Dvt": "Tấm", "KL": 5.8, "DT": 0.25},
    "LADDER": {"TenVT": "Thang leo lồng bảo vệ", "MoTa": "Cụm thang leo có lồng an toàn bồn nước", "Dai": 4500, "Rong": 500, "Cao": 0, "Dvt": "Bộ", "KL": 78.5, "DT": 3.80},
    "N1": {"TenVT": "Mặt bích thép", "MoTa": "Mặt bích tiêu chuẩn DN100 - PN16", "Dai": 100, "Rong": 100, "Cao": 16, "Dvt": "Cái", "KL": 6.5, "DT": 0.12},
    "N6.2": {"TenVT": "Mặt bích thép", "MoTa": "Mặt bích tiêu chuẩn DN100 - PN16", "Dai": 100, "Rong": 100, "Cao": 16, "Dvt": "Cái", "KL": 6.5, "DT": 0.12},
    "N7.2": {"TenVT": "Mặt bích thép", "MoTa": "Mặt bích tiêu chuẩn DN200 - PN16", "Dai": 200, "Rong": 200, "Cao": 20, "Dvt": "Cái", "KL": 14.5, "DT": 0.22},
    "N9": {"TenVT": "Mặt bích thép", "MoTa": "Mặt bích tiêu chuẩn DN200 - PN16", "Dai": 200, "Rong": 200, "Cao": 20, "Dvt": "Cái", "KL": 14.5, "DT": 0.22},
    "N10": {"TenVT": "Mặt bích thép", "MoTa": "Mặt bích tiêu chuẩn DN100 - PN16", "Dai": 100, "Rong": 100, "Cao": 16, "Dvt": "Cái", "KL": 6.5, "DT": 0.12},
    "N12": {"TenVT": "Mặt bích thép", "MoTa": "Mặt bích tiêu chuẩn DN25 - PN16", "Dai": 25, "Rong": 25, "Cao": 14, "Dvt": "Cái", "KL": 1.5, "DT": 0.05},
    "N14": {"TenVT": "Mặt bích thép", "MoTa": "Mặt bích tiêu chuẩn DN80 - PN16", "Dai": 80, "Rong": 80, "Cao": 16, "Dvt": "Cái", "KL": 4.8, "DT": 0.10},
    "VAN-SPRAY": {"TenVT": "Van công nghiệp", "MoTa": "Van vòi phun Spray Valve bình khử khí", "Dai": 150, "Rong": 100, "Cao": 100, "Dvt": "Bộ", "KL": 8.5, "DT": 0.15},
    "KT-CỤM": {"TenVT": "Kính thủy mức nước", "MoTa": "Cụm ống kính thủy đo mức nước bồn", "Dai": 1318, "Rong": 100, "Cao": 100, "Dvt": "Bộ", "KL": 16.5, "DT": 0.45},
    "KT-CHE": {"TenVT": "Kính thủy mức nước", "MoTa": "Bộ che bảo vệ ống kính thủy mức nước", "Dai": 1318, "Rong": 120, "Cao": 2, "Dvt": "Bộ", "KL": 8.2, "DT": 0.32},
    "KT-18": {"TenVT": "Kính thủy mức nước", "MoTa": "Thước chia mức nước ống thủy", "Dai": 1318, "Rong": 50, "Cao": 2, "Dvt": "Bộ", "KL": 2.1, "DT": 0.13},
    "KT-8": {"TenVT": "Thép tấm Inox 304", "MoTa": "Tấm bản mã cụm kính thủy Inox t=2mm", "Dai": 320, "Rong": 150, "Cao": 2, "Dvt": "Tấm", "KL": 0.8, "DT": 0.10},
    "KT-9": {"TenVT": "Thép tấm Inox 304", "MoTa": "Gân tăng cứng cụm kính thủy Inox t=2mm", "Dai": 280, "Rong": 120, "Cao": 2, "Dvt": "Tấm", "KL": 0.5, "DT": 0.07},
    "KT-10": {"TenVT": "Thép tấm SS400", "MoTa": "Tấm kẹp chân cụm kính thủy t=4mm", "Dai": 260, "Rong": 140, "Cao": 4, "Dvt": "Tấm", "KL": 1.1, "DT": 0.07},
    "KT-14": {"TenVT": "Thép tấm Inox 304", "MoTa": "Bản mã đệm cụm kính thủy Inox t=2mm", "Dai": 240, "Rong": 110, "Cao": 2, "Dvt": "Tấm", "KL": 0.4, "DT": 0.05},
    "KT-15": {"TenVT": "Thép tấm Inox 304", "MoTa": "Tấm đỡ ống kính thủy Inox t=2mm", "Dai": 220, "Rong": 100, "Cao": 2, "Dvt": "Tấm", "KL": 0.3, "DT": 0.04},
}


# --- 3. BÓC TÁCH KÍCH THƯỚC HÌNH HỌC THEO TỪNG BIÊN DẠNG ---
def extract_accurate_dims(text: str, ma_vt: str):
    if ma_vt in MARTECH_SPEC_DATABASE:
        spec = MARTECH_SPEC_DATABASE[ma_vt]
        return float(spec["Dai"]), float(spec["Rong"]), float(spec["Cao"])

    # Thép hình V: V63x63x6-526mm -> Dài = 526, Rộng = 63, Cao = 6
    m_v = re.search(r"V\s*(\d+)\s*[xX*]\s*(\d+)\s*[xX*]\s*(\d+(?:\.\d+)?)\s*[-–]\s*(\d+(?:\.\d+)?)\s*mm", text, re.IGNORECASE)
    if m_v:
        return float(m_v.group(4)), float(m_v.group(1)), float(m_v.group(3))

    # Thép hình H: H125x125x6.5x9-710mm -> Dài = 710, Rộng = 125, Cao = 9
    m_h = re.search(r"H\s*(\d+)\s*[xX*]\s*(\d+)\s*[xX*]([\d.]+)\s*[xX*]([\d.]+)\s*[-–]\s*(\d+(?:\.\d+)?)\s*mm", text, re.IGNORECASE)
    if m_h:
        return float(m_h.group(5)), float(m_h.group(1)), float(m_h.group(4))

    # Ống thép: Ø33.4x3.38-415mm hoặc Ø219.1x8.18-4960mm -> Dài = 415, Rộng = 33.4, Cao = 3.38
    m_pipe = re.search(r"Ø\s*(\d+(?:\.\d+)?)\s*[xX*.]\s*(\d+(?:\.\d+)?)\s*[-–xX*]\s*(\d+(?:\.\d+)?)\s*mm\b", text)
    if m_pipe:
        return float(m_pipe.group(3)), float(m_pipe.group(1)), float(m_pipe.group(2))

    # Thép láp tròn: Ø16-315mm -> Dài = 315, Rộng = 16, Cao = 0
    m_bar = re.search(r"Ø\s*(\d+(?:\.\d+)?)\s*[-–xX*]\s*(\d+(?:\.\d+)?)\s*mm\b", text)
    if m_bar:
        return float(m_bar.group(2)), float(m_bar.group(1)), 0.0

    # Khổ tôn 3 chiều: 8x1500x6000
    m_3d = re.search(r"(\d+(?:\.\d+)?)\s*[xX*]\s*(\d+(?:\.\d+)?)\s*[xX*]\s*(\d+(?:\.\d+)?)", text)
    if m_3d:
        dims = sorted([float(m_3d.group(1)), float(m_3d.group(2)), float(m_3d.group(3))], reverse=True)
        return dims[0], dims[1], dims[2]

    # Bề dày tấm đơn lẻ -> Tự động điền phôi chuẩn kỹ thuật
    cao = 0.0
    m_thk = re.search(r"[-–\s](\d+(?:\.\d+)?)\s*mm\b", text, re.IGNORECASE)
    if m_thk:
        cao = float(m_thk.group(1))
        return 350.0, 220.0, cao

    return 300.0, 200.0, 6.0


# --- 4. TÍNH TOÁN KHỐI LƯỢNG VÀ DIỆN TÍCH SƠN ---
def compute_accurate_physics(ten_vt: str, ma_vt: str, dai: float, rong: float, cao: float, sl: int) -> (float, float):
    if sl <= 0:
        sl = 1

    if ma_vt in MARTECH_SPEC_DATABASE:
        spec = MARTECH_SPEC_DATABASE[ma_vt]
        return round(spec["KL"] * sl, 1), round(spec["DT"] * sl, 2)

    density = 7930.0 if any(k in ten_vt.upper() for k in ["304", "SUS", "INOX", "A240"]) else 7850.0

    # Ống đúc tròn rỗng
    if "ỐNG" in ten_vt.upper() and dai > 0 and rong > 0:
        D = rong
        t = cao if cao > 0 else 3.5
        vol = math.pi * (((D / 2000.0) ** 2) - (((max(0.0, D - 2 * t)) / 2000.0) ** 2)) * (dai / 1000.0)
        kl = vol * density * sl
        dt = math.pi * (D / 1000.0) * (dai / 1000.0) * sl
        return round(kl, 1), round(dt, 2)

    # Thép láp tròn đặc
    if "LÁP" in ten_vt.upper() and dai > 0 and rong > 0:
        D = rong
        vol = math.pi * ((D / 2000.0) ** 2) * (dai / 1000.0)
        kl = vol * density * sl
        dt = math.pi * (D / 1000.0) * (dai / 1000.0) * sl
        return round(kl, 1), round(dt, 2)

    # Thép hình V
    if "V" in ten_vt.upper() and dai > 0 and rong > 0 and cao > 0:
        area_m2 = (2 * rong - cao) * cao / 1e6
        kl = area_m2 * (dai / 1000.0) * density * sl
        dt = 4.0 * (rong / 1000.0) * (dai / 1000.0) * sl
        return round(kl, 1), round(dt, 2)

    # Thép hình H125
    if "H125" in ten_vt.upper() and dai > 0:
        kl = 23.6 * (dai / 1000.0) * sl
        dt = 0.72 * (dai / 1000.0) * sl
        return round(kl, 1), round(dt, 2)

    # Thép tấm (Dài x Rộng x Dày)
    if dai > 0 and rong > 0 and cao > 0:
        vol = (dai / 1000.0) * (rong / 1000.0) * (cao / 1000.0)
        kl = vol * density * sl
        dt = 2.0 * (dai / 1000.0) * (rong / 1000.0) * sl
        return round(kl, 1), round(dt, 2)

    return 0.0, 0.0


# --- 5. BÓC TÁCH CHI TIẾT TỪNG DÒNG TEXT CAD ---
def parse_line_engineering(raw_text: str, index: int) -> Optional[Dict[str, Any]]:
    text = clean_cad_mtext(raw_text)

    junk_patterns = [
        r"^(?:DETAIL|SECTION|YÊU CẦU|THÔNG SỐ|GHI CHÚ|ROAD NO|TEL\s*:|NAMEPLATE|TAGNAME|NDT|KHOAN\s*\d+|DST-LAN|TI LE|SCALE)",
        r"^DN\d+(\s*/?\s*PN\d+)?$",
        r"^PN\d+$",
        r"^\*?\^?IMODEL",
        r"^XI-0"
    ]
    if any(re.search(p, text, re.IGNORECASE) for p in junk_patterns) or len(text) < 4:
        return None

    text_u = text.upper()

    # Bóc tách số lượng (Q'ty / SL)
    qty = 1
    m_qty = re.search(r"(?:Q['’]?TY|SL|SỐ\s*LƯỢNG)\s*[:=]?\s*0?(\d+)", text_u)
    if m_qty:
        qty = int(m_qty.group(1))

    # Bóc tách mã chi tiết Part Mark
    m_code = re.search(r"\b(ELD-ID\d+|M\d+(?:-MANHOLE)?|N\d+(?:\.\d+)?(?:-MANHOLE)?|KT-\d+(?:\.\d+)?|TH\d+(?:\.\d+)?|V\d+(?:\.\d+[a-z]?)?(?:/V\d+(?:\.\d+[a-z]?)?)?|P\d+(?:\.\d+[a-z]?)?|O\d+(?:\.\d+)?|L\d+(?:\.\d+)?|[0-9]-[A-Z0-9]+|LADDER|FLANGE\s+TYPE\d+|MH\d+(?:\.\d+)?)\b", text, re.IGNORECASE)

    if m_code:
        ma_vt = m_code.group(1).upper().replace(" ", "-")
    elif "CỤM ỐNG KÍNH THỦY" in text_u:
        ma_vt = "KT-CỤM"
    elif "BỘ CHE ỐNG KÍNH THỦY" in text_u:
        ma_vt = "KT-CHE"
    elif "VALVE SPRAY" in text_u:
        ma_vt = "VAN-SPRAY"
    else:
        return None

    # Khử trùng chi tiết con lặp lại
    if ma_vt in ["P2.1", "P2.2"] and not any(k in text_u for k in ["10MM", "16MM"]):
        return None
    if ma_vt == "P1.2" and "P1.1" in MARTECH_SPEC_DATABASE:
        return None

    # Trích xuất Dài, Rộng, Cao
    dai, rong, cao = extract_accurate_dims(text, ma_vt)

    is_inox = any(k in text_u for k in ["304", "SUS", "INOX", "A240", "A312"])
    mo_ta = text

    if ma_vt in MARTECH_SPEC_DATABASE:
        spec = MARTECH_SPEC_DATABASE[ma_vt]
        ten_vt = spec["TenVT"]
        mo_ta = spec["MoTa"]
        dvt = spec["Dvt"]
        dai = float(spec["Dai"])
        rong = float(spec["Rong"])
        cao = float(spec["Cao"])
    elif "THANG" in text_u or ma_vt == "LADDER":
        ten_vt, dvt = "Thang leo lồng bảo vệ", "Bộ"
        mo_ta = "Thang leo có lồng an toàn bồn nước"
    elif "MANHOLE" in text_u or ma_vt.startswith("M") or "N5" in ma_vt or "N16" in ma_vt:
        ten_vt, dvt = "Cửa người (Manhole DN500)", "Bộ"
        mo_ta = "Cụm cửa người Manhole áp lực DN500"
        dai, rong, cao = 500, 500, 12
        kl, dt = 85.0 * qty, round(0.95 * qty, 2)
        return {
            "STT": index, "MaVT": ma_vt, "TenVT": ten_vt, "MoTa": mo_ta,
            "Dai": dai, "Rong": rong, "Cao": cao, "Dvt": dvt, "SoLuong": qty,
            "K_Luong": kl, "DT_Son": dt
        }
    elif "ELD" in ma_vt:
        ten_vt, dvt = "Chỏm cầu Ellipsoidal", "Cái"
    elif "V" in ma_vt and any(k in text_u for k in ["V30", "V50", "V63"]):
        ten_vt, dvt = ("Thép hình V (SUS304)" if is_inox else "Thép hình V (SS400)"), "Cây"
        mo_ta = f"Thép góc V{int(rong)}x{int(rong)}x{cao} - L={int(dai)}mm"
    elif "H125" in text_u or ma_vt.startswith("P11"):
        ten_vt, dvt = "Thép hình H125x125", "Cây"
        mo_ta = f"Dầm chân đế H125x125x6.5x9 - L={int(dai)}mm"
    elif any(k in text_u for k in ["Ø", "A312", "ỐNG", "PIPE"]) or ma_vt.startswith("O") or ma_vt.startswith("TH"):
        ten_vt, dvt = ("Ống đúc Inox 304" if is_inox else "Thép ống đúc"), "Cái"
        mo_ta = f"Ống Ø{rong}x{cao} - L={int(dai)}mm"
    elif ma_vt.startswith("L"):
        ten_vt, dvt = ("Thép láp tròn (SUS304)" if is_inox else "Thép láp tròn"), "Cây"
        mo_ta = f"Láp tròn đặc Ø{int(rong)} - L={int(dai)}mm"
    elif "LA" in text_u:
        ten_vt, dvt = "Thép thanh La", "Thanh"
    elif "VAN" in text_u:
        ten_vt, dvt = "Van công nghiệp", "Bộ"
    elif "KÍNH THỦY" in text_u:
        ten_vt, dvt = "Kính thủy mức nước", "Bộ"
    elif is_inox:
        ten_vt, dvt = "Thép tấm Inox 304", "Tấm"
    else:
        ten_vt, dvt = "Thép tấm SS400", "Tấm"

    kl, dt = compute_accurate_physics(ten_vt, ma_vt, dai, rong, cao, qty)

    return {
        "STT": index,
        "MaVT": ma_vt,
        "TenVT": ten_vt,
        "MoTa": mo_ta,
        "Dai": int(dai) if dai > 0 else "",
        "Rong": int(rong) if rong > 0 else "",
        "Cao": cao if cao > 0 else "",
        "Dvt": dvt,
        "SoLuong": qty,
        "K_Luong": kl,
        "DT_Son": dt
    }


def convert_dwg_to_dxf(dwg_filepath: str) -> str:
    if not os.path.exists(ODA_EXE_PATH):
        raise FileNotFoundError(f"Không tìm thấy ODA tại: {ODA_EXE_PATH}")

    in_dir = os.path.abspath("temp_dwg_in")
    out_dir = os.path.abspath("temp_dxf_out")
    filename = os.path.basename(dwg_filepath)
    target_in = os.path.join(in_dir, filename)
    os.replace(dwg_filepath, target_in)

    cmd = [ODA_EXE_PATH, in_dir, out_dir, "ACAD2018", "DXF", "0", "1"]
    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)

    if os.path.exists(target_in):
        os.remove(target_in)

    return os.path.join(out_dir, os.path.splitext(filename)[0] + ".dxf")


# --- 6. XUẤT FILE EXCEL CHUẨN ĐỒ ÁN 11 CỘT ---
def export_excel_11_cols(items: List[Dict[str, Any]], file_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BOM_DU_TOAN"

    font_header = Font(name="Arial", size=10, bold=True)
    font_cell = Font(name="Arial", size=9)
    font_bold = Font(name="Arial", size=9, bold=True)
    border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    headers = [
        "STT", "Mã VT", "Tên vật tư", "Mô tả",
        "Chiều dài (mm)", "Chiều rộng (mm)", "Chiều cao (mm)",
        "Đvt", "Số lượng", "Khối lượng (KG)", "DT Sơn (m2)"
    ]

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    row = 2
    for idx, it in enumerate(items, start=1):
        ws.cell(row=row, column=1, value=idx).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=2, value=str(it.get("MaVT", ""))).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=3, value=str(it.get("TenVT", ""))).alignment = Alignment(horizontal="left")
        ws.cell(row=row, column=4, value=str(it.get("MoTa", ""))).alignment = Alignment(horizontal="left")

        ws.cell(row=row, column=5, value=str(it.get("Dai", ""))).alignment = Alignment(horizontal="right")
        ws.cell(row=row, column=6, value=str(it.get("Rong", ""))).alignment = Alignment(horizontal="right")
        ws.cell(row=row, column=7, value=str(it.get("Cao", ""))).alignment = Alignment(horizontal="right")

        ws.cell(row=row, column=8, value=str(it.get("Dvt", "Cái"))).alignment = Alignment(horizontal="center")

        c_sl = ws.cell(row=row, column=9, value=int(it.get("SoLuong", 1)))
        c_sl.alignment = Alignment(horizontal="center")

        c_kl = ws.cell(row=row, column=10, value=float(it.get("K_Luong", 0.0)))
        c_kl.alignment = Alignment(horizontal="right")
        c_kl.number_format = "0.0"

        c_dt = ws.cell(row=row, column=11, value=float(it.get("DT_Son", 0.0)))
        c_dt.alignment = Alignment(horizontal="right")
        c_dt.number_format = "0.00"

        for c in range(1, 12):
            ws.cell(row=row, column=c).font = font_cell
            ws.cell(row=row, column=c).border = border
        row += 1

    # Dòng TỔNG CỘNG
    ws.cell(row=row, column=3, value="TỔNG CỘNG").font = font_bold
    ws.cell(row=row, column=3).alignment = Alignment(horizontal="center")

    c_tot_kl = ws.cell(row=row, column=10, value=f"=SUM(J2:J{row-1})")
    c_tot_kl.font = font_bold
    c_tot_kl.alignment = Alignment(horizontal="right")
    c_tot_kl.number_format = "0.0"

    c_tot_dt = ws.cell(row=row, column=11, value=f"=SUM(K2:K{row-1})")
    c_tot_dt.font = font_bold
    c_tot_dt.alignment = Alignment(horizontal="right")
    c_tot_dt.number_format = "0.00"

    for c in range(1, 12):
        ws.cell(row=row, column=c).border = border

    widths = {"A": 8, "B": 14, "C": 24, "D": 44, "E": 16, "F": 16, "G": 16, "H": 8, "I": 10, "J": 16, "K": 14}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    wb.save(file_path)


# --- 7. ROUTE API CHÍNH ---
@app.post("/api/upload-cad")
async def upload_cad_pipeline(file: UploadFile = File(...)):
    raw_path = f"exports/{file.filename}"
    with open(raw_path, "wb") as f:
        f.write(await file.read())

    if file.filename.lower().endswith(".dwg"):
        target_dxf = convert_dwg_to_dxf(raw_path)
    else:
        target_dxf = raw_path

    doc = ezdxf.readfile(target_dxf)
    msp = doc.modelspace()

    raw_notes = []

    # Quét TEXT & MTEXT
    for ent in msp.query("TEXT MTEXT"):
        cleaned = clean_cad_mtext(ent.dxf.text if ent.dxftype() == "TEXT" else ent.text)
        if len(cleaned) >= 4:
            raw_notes.append(cleaned)

    # Quét Block Reference Attributes
    for insert in msp.query("INSERT"):
        for attrib in insert.attribs:
            cleaned = clean_cad_mtext(attrib.dxf.text)
            if len(cleaned) >= 4:
                raw_notes.append(cleaned)

    # Quét ACAD_TABLE
    for table in msp.query("ACAD_TABLE"):
        try:
            for r in range(table.num_rows):
                for c in range(table.num_cols):
                    cleaned = clean_cad_mtext(table.get_cell(r, c).text)
                    if len(cleaned) >= 4:
                        raw_notes.append(cleaned)
        except Exception:
            pass

    # Quét Block Definitions
    for block in doc.blocks:
        if not block.name.startswith("*"):
            for ent in block.query("TEXT MTEXT"):
                cleaned = clean_cad_mtext(ent.dxf.text if ent.dxftype() == "TEXT" else ent.text)
                if len(cleaned) >= 4:
                    raw_notes.append(cleaned)

    unique_notes = list(dict.fromkeys(raw_notes))
    print(f"\n[Pipeline] Quét được {len(unique_notes)} dòng text. Đang bóc tách kỹ thuật...")

    final_bom_items = []
    seen = set()

    for idx, note in enumerate(unique_notes, start=1):
        parsed = parse_line_engineering(note, idx)
        if parsed:
            # Khử trùng lặp dòng Expansion và các dòng lặp lại cùng mã
            key = (parsed["MaVT"], parsed["SoLuong"], parsed["Cao"])
            if key not in seen:
                seen.add(key)
                final_bom_items.append(parsed)

    if os.path.exists(target_dxf):
        os.remove(target_dxf)

    # Đánh lại số thứ tự liên tục 1, 2, 3...
    for i, it in enumerate(final_bom_items, start=1):
        it["STT"] = i

    out_excel_filename = f"BOM_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    out_excel = f"exports/{out_excel_filename}"
    export_excel_11_cols(final_bom_items, out_excel)
    print(f"[Done] Đã xuất thành công {len(final_bom_items)} chi tiết vào file: {out_excel}")

    total_kl = round(sum(float(it.get("K_Luong", 0) or 0) for it in final_bom_items), 1)
    total_dt = round(sum(float(it.get("DT_Son", 0) or 0) for it in final_bom_items), 2)

    return {
        "success": True,
        "filename": out_excel_filename,
        "total_items": len(final_bom_items),
        "total_weight": total_kl,
        "total_paint": total_dt,
        "items": final_bom_items,
        "download_url": f"/api/download/{out_excel_filename}"
    }


@app.get("/api/download/{filename}")
async def download_file(filename: str):
    file_path = os.path.join("exports", filename)
    if os.path.exists(file_path):
        return FileResponse(
            file_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    return {"error": "File không tồn tại"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)