import os
import subprocess
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse

# Đường dẫn file chạy ODA File Converter trên máy bạn
ODA_PATH = r"D:\Program Files\ODA\ODAFileConverter 25.1.0\ODAFileConverter.exe"

def convert_dwg_to_dxf(input_dwg_path: str) -> str:
    """Tự động chuyển file DWG sang DXF bằng ODA CLI"""
    input_dir = os.path.abspath("temp_in")
    output_dir = os.path.abspath("temp_out")
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    # Chuyển file dwg vào thư mục in
    filename = os.path.basename(input_dwg_path)
    target_dwg = os.path.join(input_dir, filename)
    os.replace(input_dwg_path, target_dwg)

    # Lệnh gọi ODA CLI: ODAFileConverter.exe "thư_mục_in" "thư_mục_out" "phiên_bản_acad" "định_dạng" "đệ_quy" "kiểm_tra"
    # ACAD2018 DXF 0 1
    cmd = [
        ODA_PATH,
        input_dir,
        output_dir,
        "ACAD2018",
        "DXF",
        "0",
        "1"
    ]
    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    # Đường dẫn file dxf sau khi convert
    dxf_filename = os.path.splitext(filename)[0] + ".dxf"
    result_dxf = os.path.join(output_dir, dxf_filename)

    # Dọn dẹp file dwg tạm
    if os.path.exists(target_dwg):
        os.remove(target_dwg)

    return result_dxf

@app.post("/api/upload-cad")
async def upload_cad_file(file: UploadFile = File(...)):
    # Nhận cả DWG hoặc DXF
    temp_file = f"temp_{file.filename}"
    with open(temp_file, "wb") as f:
        f.write(await file.read())

    # Nếu là file DWG thì tự động convert sang DXF ngầm
    if file.filename.lower().endswith(".dwg"):
        target_dxf = convert_dwg_to_dxf(temp_file)
    else:
        target_dxf = temp_file

    # Bóc tách dữ liệu từ DXF bằng ezdxf
    extracted_items = parse_dxf_logic(target_dxf)

    # Xóa file tạm
    if os.path.exists(target_dxf):
        os.remove(target_dxf)

    # Xuất Excel
    out_file = f"exports/BOM_{os.path.splitext(file.filename)[0]}.xlsx"
    export_excel(extracted_items, out_file)
    
    return FileResponse(out_file, filename=os.path.basename(out_file))

#"D:\Program Files\ODA\ODAFileConverter 27.1.0\ODAFileConverter.exe"