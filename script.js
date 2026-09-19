/**
 * AutoCAD AI BOM Extractor - Frontend Script
 * Tương tác API FastAPI, render bảng dự toán động và quản lý tải file.
 */

// Cấu hình Base URL API Backend
const API_BASE = "https://cad-ai-extractor.onrender.com";

// DOM Elements: Khu vực Upload & Điều khiển
const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const fileCard = document.getElementById("fileCard");
const fileName = document.getElementById("fileName");
const fileSize = document.getElementById("fileSize");
const removeFileBtn = document.getElementById("removeFileBtn");
const processBtn = document.getElementById("processBtn");
const stepperWrapper = document.getElementById("stepperWrapper");

// DOM Elements: Thống kê Dashboard
const metricsGrid = document.getElementById("metricsGrid");
const metricTotalItems = document.getElementById("metricTotalItems");
const metricTotalWeight = document.getElementById("metricTotalWeight");
const metricTotalPaint = document.getElementById("metricTotalPaint");

// DOM Elements: Bảng & Tải file
const tableContainer = document.getElementById("tableContainer");
const emptyState = document.getElementById("emptyState");
const bomTable = document.getElementById("bomTable");
const bomTableBody = document.getElementById("bomTableBody");
const downloadBtn = document.getElementById("downloadBtn");

// Trạng thái ứng dụng
let selectedFile = null;
let downloadUrl = "";

// ==========================================
// 1. XỬ LÝ SỰ KIỆN KÉO THẢ & CHỌN FILE CAD
// ==========================================

// Sự kiện Drag & Drop
dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
});

dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("dragover");
});

dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        handleFileSelect(e.dataTransfer.files[0]);
    }
});

// Sự kiện duyệt file từ thẻ input
fileInput.addEventListener("change", (e) => {
    if (e.target.files && e.target.files.length > 0) {
        handleFileSelect(e.target.files[0]);
    }
});

/**
 * Kiểm tra định dạng và lưu file người dùng chọn
 * @param {File} file 
 */
function handleFileSelect(file) {
    const validExts = [".dwg", ".dxf"];
    const fileExt = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();
    
    if (!validExts.includes(fileExt)) {
        alert("Định dạng không hợp lệ! Vui lòng chọn bản vẽ AutoCAD (.dwg hoặc .dxf).");
        return;
    }

    selectedFile = file;
    fileName.textContent = file.name;
    fileSize.textContent = (file.size / (1024 * 1024)).toFixed(2) + " MB";

    // Cập nhật giao diện
    dropzone.style.display = "none";
    fileCard.style.display = "flex";
    processBtn.disabled = false;
}

// Xóa file đã chọn
removeFileBtn.addEventListener("click", () => {
    selectedFile = null;
    fileInput.value = "";
    dropzone.style.display = "block";
    fileCard.style.display = "none";
    processBtn.disabled = true;
    stepperWrapper.style.display = "none";
});

// ==========================================
// 2. ĐIỀU KHIỂN TIẾN TRÌNH (STEPPER)
// ==========================================

/**
 * Đổi trạng thái hiển thị của các bước bóc tách
 * @param {number} stepNumber (1 đến 4)
 */
function setStepActive(stepNumber) {
    for (let i = 1; i <= 4; i++) {
        const el = document.getElementById(`step${i}`);
        if (el) {
            if (i <= stepNumber) {
                el.classList.add("active");
            } else {
                el.classList.remove("active");
            }
        }
    }
}

// ==========================================
// 3. GỬI REQUEST BÓC TÁCH TỚI FASTAPI
// ==========================================

processBtn.addEventListener("click", async () => {
    if (!selectedFile) return;

    // Vô hiệu hóa nút và kích hoạt stepper
    processBtn.disabled = true;
    stepperWrapper.style.display = "flex";
    setStepActive(1);

    const formData = new FormData();
    formData.append("file", selectedFile);

    try {
        // Mô phỏng tiến trình phân tích giai đoạn
        setTimeout(() => setStepActive(2), 300);
        setTimeout(() => setStepActive(3), 600);

        const response = await fetch(`${API_BASE}/api/upload-cad`, {
            method: "POST",
            body: formData
        });

        if (!response.ok) {
            const errData = await response.json().catch(() => ({}));
            throw new Error(errData.detail || `Lỗi máy chủ (${response.status}): ${response.statusText}`);
        }

        setStepActive(4);

        const result = await response.json();
        
        // Cập nhật đường dẫn file tải về
        downloadUrl = `${API_BASE}${result.download_url}`;

        // Hiển thị dữ liệu thực tế lên bảng
        renderLiveBOMTable(result.items || [], result.total_weight || 0, result.total_paint || 0);
        
        // Hiển thị nút tải file Excel
        downloadBtn.style.display = "inline-flex";

    } catch (err) {
        console.error("Pipeline Error:", err);
        alert("Không thể bóc tách bản vẽ: " + err.message);
    } finally {
        processBtn.disabled = false;
    }
});

// ==========================================
// 4. RENDER BẢNG PREVIEW ĐỘNG & DASHBOARD
// ==========================================

/**
 * Hiển thị dữ liệu toàn bộ các chi tiết vật tư vào DOM
 * @param {Array} items Danh sách chi tiết bóc tách được
 * @param {number} totalWeight Tổng khối lượng phôi (KG)
 * @param {number} totalPaint Tổng diện tích sơn (m2)
 */
function renderLiveBOMTable(items, totalWeight, totalPaint) {
    // Ẩn thông báo trống, hiển thị bảng và các thẻ số liệu
    emptyState.style.display = "none";
    bomTable.style.display = "table";
    metricsGrid.style.display = "grid";

    // Cập nhật thẻ số liệu Dashboard
    metricTotalItems.textContent = items.length;
    metricTotalWeight.innerHTML = `${Number(totalWeight).toLocaleString("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 })} <small>KG</small>`;
    metricTotalPaint.innerHTML = `${Number(totalPaint).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} <small>m²</small>`;

    // Xây dựng danh sách hàng cho bảng
    let rowsHtml = items.map((r, index) => {
        const stt = r.STT || (index + 1);
        const maVt = r.MaVT || "-";
        const tenVt = r.TenVT || "Vật tư gia công";
        const moTa = r.MoTa || "";
        const dai = r.Dai !== "" && r.Dai !== null && r.Dai !== undefined ? r.Dai : "-";
        const rong = r.Rong !== "" && r.Rong !== null && r.Rong !== undefined ? r.Rong : "-";
        const cao = r.Cao !== "" && r.Cao !== null && r.Cao !== undefined ? r.Cao : "-";
        const dvt = r.Dvt || "Cái";
        const sl = r.SoLuong !== undefined ? r.SoLuong : 1;
        const kl = Number(r.K_Luong || 0).toFixed(1);
        const dt = Number(r.DT_Son || 0).toFixed(2);

        return `
            <tr>
                <td class="text-center font-mono">${stt}</td>
                <td class="text-center font-mono"><strong>${maVt}</strong></td>
                <td>${tenVt}</td>
                <td title="${moTa.replace(/"/g, '&quot;')}">${moTa}</td>
                <td class="text-right font-mono">${dai}</td>
                <td class="text-right font-mono">${rong}</td>
                <td class="text-right font-mono">${cao}</td>
                <td class="text-center">${dvt}</td>
                <td class="text-center font-mono">${sl}</td>
                <td class="text-right font-mono"><strong>${kl}</strong></td>
                <td class="text-right font-mono">${dt}</td>
            </tr>
        `;
    }).join("");

    // Tính tổng số lượng cấu kiện
    const totalQty = items.reduce((acc, it) => acc + Number(it.SoLuong || 1), 0);

    // Bổ sung hàng TỔNG CỘNG ở cuối bảng
    rowsHtml += `
        <tr class="total-row">
            <td colspan="4" class="text-center">TỔNG CỘNG</td>
            <td></td>
            <td></td>
            <td></td>
            <td></td>
            <td class="text-center font-mono">${totalQty}</td>
            <td class="text-right font-mono">${Number(totalWeight).toLocaleString("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 })}</td>
            <td class="text-right font-mono">${Number(totalPaint).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
        </tr>
    `;

    bomTableBody.innerHTML = rowsHtml;
}

// ==========================================
// 5. TẢI FILE EXCEL (.XLSX)
// ==========================================

downloadBtn.addEventListener("click", () => {
    if (!downloadUrl) {
        alert("Chưa có liên kết tải file! Vui lòng bóc tách lại bản vẽ.");
        return;
    }
    // Kích hoạt điều hướng để tải file trực tiếp từ backend
    window.location.href = downloadUrl;
});