# PROJECT CONTEXT — App Phân Tích Luân Chuyển Hàng Hoá

> **Quy tắc bắt buộc:** Sau mỗi lần sửa code, cập nhật section **[CHANGELOG]** ở cuối file này,
> ghi rõ ngày, nội dung thay đổi và lý do.

---

## 1. Mục đích ứng dụng

Desktop app (Python/Tkinter) giúp bộ phận bán lẻ phân tích dữ liệu bán hàng + tồn kho và đưa ra
2 loại đề xuất:

| Sheet output | Mục tiêu |
|---|---|
| **Fill từ kho** | Kho bổ sung hàng về từng cửa hàng để đảm bảo đủ X tháng sức bán |
| **Luân chuyển cửa hàng** | Chuyển hàng từ cửa hàng bán chậm sang cửa hàng bán nhanh đang thiếu hàng |

---

## 2. Cấu trúc file

```
main.py                  # Toàn bộ logic + GUI (single-file app)
requirements.txt         # Dependencies Python
app.spec                 # PyInstaller build spec
build_windows.bat        # Script build exe cho Windows
File_template/
  data sale out.xlsx     # Template file bán hàng (mẫu nhập liệu)
  inventory.xlsx         # Template file tồn kho (mẫu nhập liệu)
  Bluecircle.png         # Logo hiển thị trên header
.github/workflows/
  build.yml              # CI/CD: auto build exe khi push
```

---

## 3. Input data — cấu trúc cột bắt buộc

### File bán hàng (sales)

| Cột | Mô tả |
|---|---|
| `Date` | Ngày bán (parseable bởi pandas) |
| `Sale Team` | Kênh bán — chỉ xử lý hàng có giá trị **`Bán Lẻ`** |
| `location Name` | Tên cửa hàng |
| `SKU` | Mã hàng nội bộ |
| `Product Item` | Tên hiển thị sản phẩm |
| `Quantity` | Số lượng bán |

### File tồn kho (inventory)

| Cột | Mô tả |
|---|---|
| `Địa điểm/Team` | `Bán Lẻ` (tồn cửa hàng) hoặc `Kho` (tồn kho trung tâm) |
| `Địa điểm/Tên hiển thị` | Tên cửa hàng |
| `Sản phẩm/Mã nội bộ` | SKU |
| `Sản phẩm/Tên hiển thị` | Tên sản phẩm |
| `Số lượng` | Số lượng tồn |

---

## 4. Logic tính toán (`calculate()`)

### 4.1 Tiền xử lý
- Lọc `Sale Team == 'Bán Lẻ'` từ file bán hàng.
- Lấy **3 tháng gần nhất** tính từ `max(Date)` trong dữ liệu.
- Tính `actual_months = min(3, khoảng ngày thực tế / 30.44)` — tránh chia sai khi dữ liệu ngắn hơn 3 tháng.
- Có thể lọc theo danh sách `selected_stores` (do user chọn trên UI).

### 4.2 Tính velocity (sức bán)
```
AvgMonthly (store, SKU) = Sold3M / actual_months
```
Nhóm theo `(location Name, SKU)` trong 3 tháng, chia cho số tháng thực tế.

### 4.3 Sheet 1 — Fill từ kho
**Điều kiện bỏ qua:** `AvgMonthly <= min_fill_avg` (default 0.4 SP/tháng) → không fill.

Với mỗi cặp `(store, SKU)` còn lại:
```
TargetStock   = ceil(AvgMonthly × target_months)
SuggestedFill = max(0, TargetStock − StoreQty)
```

**Phân bổ kho có hạn:** Sắp xếp theo `AvgMonthly` giảm dần (ưu tiên hàng bán nhanh).
Duyệt từng dòng, trừ dần tồn kho còn lại:
```
FinalFill = min(SuggestedFill, kho_còn_lại_của_SKU)
```
Chỉ xuất ra dòng có `FinalFill > 0`.

### 4.4 Sheet 2 — Luân chuyển cửa hàng

**Xác định ngưỡng bán chậm:**
```
threshold(SKU) = avg(AvgMonthly tất cả stores) × (slow_pct / 100)
```

**Phân loại cửa hàng theo SKU:**
- `slow_store`: `StoreQty > 0` **và** `AvgMonthly <= threshold` → có hàng nhưng bán chậm → nguồn gửi đi
- `fast_store`: `AvgMonthly > threshold` **và** `Need > 0` → bán nhanh nhưng đang thiếu hàng → nhận hàng

```
Need(fast_store, SKU) = ceil(AvgMonthly × target_months) − StoreQty
```

**Matching:** Duyệt fast_stores theo Need giảm dần, lần lượt lấy từ các slow_stores:
```
TransferQty = min(need_còn_lại, tồn_slow_còn_lại)
```

**Quy tắc khu vực (allow_hcm_hn):**
- Nếu `allow_hcm_hn = False` (default): **chặn** luân chuyển HCM ↔ HN.
- Phát hiện khu vực từ tên cửa hàng: `'HCM' in name` → HCM, `'HN' in name` → HN, còn lại → OTHER.
- OTHER có thể luân chuyển với bất kỳ khu vực nào.

---

## 5. Tham số người dùng (UI)

| Tham số | Default | Ý nghĩa |
|---|---|---|
| `target_months` | 1.5 | Mức tồn mục tiêu (tháng sức bán) |
| `slow_pct` | 50% | Ngưỡng bán chậm (% so với TB mạng lưới) |
| `min_fill_avg` | 0.4 | AvgMonthly tối thiểu để fill |
| `allow_hcm_hn` | False | Cho phép luân chuyển liên vùng HCM↔HN |

---

## 6. Output Excel (`export_excel()`)

Dùng `xlsxwriter` engine. Mỗi sheet có:
- **Row 0:** Title bar màu navy (`#0D47A1`), merge toàn bộ cột.
- **Row 1:** Header màu xanh (`#1565C0`), chữ trắng, in đậm, wrap text.
- **Freeze panes:** 2 hàng đầu + 3 cột đầu.
- **Alternating rows:** trắng / xanh nhạt (`#DCEEFB`).
- **Subtotal row** cuối mỗi nhóm: nền xanh (`#1E88E5`), chữ trắng.
- Cột highlight: `Fill thực tế` (xanh đậm `#90CAF9`), `Đề xuất luân chuyển` (vàng `#FFD54F`).
- Auto-fit column width (capped 48 ký tự).

---

## 7. GUI — các thành phần

```
App (tk.Tk)
├── header          Logo + tiêu đề
├── body
│   ├── import_card     Chọn 2 file Excel (bán hàng + tồn kho)
│   ├── store_card      Listbox cửa hàng vật lý | kênh Shopee (multi-select)
│   ├── settings_card   Spinbox: target_months, slow_pct, min_fill_avg + checkbox HCM↔HN
│   ├── action_bar      Btn "Phân Tích" | Btn "Xuất Excel" | ProgressBar | status label
│   └── results         Notebook 2 tab → Treeview (Fill từ kho | Luân chuyển cửa hàng)
```

**Threading:** Đọc file + tính toán chạy trên daemon thread riêng. Cập nhật UI thông qua `self.after(0, callback)` để tránh block main thread.

**CheckListbox:** Widget tự build (Canvas + scrollable Frame + Checkbutton). Thay thế `tk.Listbox`. Hỗ trợ `set_items()`, `select_all()`, `deselect_all()`, `get_selected()`. Có mousewheel scroll.

**Auto-load stores:** Khi cả 2 path file thay đổi (`StringVar.trace`), tự động đọc danh sách cửa hàng và populate 2 CheckListbox. Quy tắc phân loại:
- **Kho Cửa Hàng** (cột trái): store có doanh số trong 3 tháng gần nhất VÀ không khớp từ khóa "other".
- **Các kho khác** (cột phải): store khớp `OTHER_KW = ('shopee', 'droppi', 'xe thuê', 'kho kg')` HOẶC không có doanh số 3 tháng gần nhất.

**DPI Windows:** `SetProcessDpiAwareness(2)` (per-monitor) được gọi trước khi tạo cửa sổ Tk, fallback về `SetProcessDPIAware()` nếu API cũ.

---

## 8. Build & Deploy

- **PyInstaller** (`app.spec`) đóng gói thành `.exe` kèm `File_template/`.
- `build_windows.bat` chạy PyInstaller trên Windows.
- **GitHub Actions** (`.github/workflows/build.yml`): tự build exe và tạo release khi push lên `main`.

---

## 9. Dependencies

| Package | Mục đích |
|---|---|
| `pandas >= 2.0` | Đọc Excel, xử lý DataFrame |
| `openpyxl >= 3.1` | Engine đọc `.xlsx` cho pandas |
| `xlsxwriter >= 3.1` | Ghi Excel có format phong phú |
| `Pillow >= 10.0` | Load logo PNG cho header |
| `pyinstaller >= 6.0` | Build exe |

---

## 10. Điểm cần chú ý khi sửa code

1. **SKU luôn cast sang `str`** trước khi merge/join — tránh mismatch int vs str.
2. **`actual_months`** phải > 0 (clamp về 0.1) — tránh chia cho 0.
3. **Kho bị trừ dần** trong vòng lặp fill → thứ tự duyệt quan trọng (sort by AvgMonthly desc).
4. **`self.after(0, ...)`** bắt buộc khi update widget từ thread phụ.
5. **`resource_path()`** dùng cho tất cả asset tĩnh — hỗ trợ cả dev lẫn PyInstaller bundle.
6. Khi thêm tham số mới vào `calculate()`, nhớ cập nhật cả: UI widget → `_analysis_worker` → `calculate()` signature.

---

## CHANGELOG

| Ngày | Thay đổi | Lý do |
|---|---|---|
| 2026-04-13 | Tạo file PROJECT_CONTEXT.md | Tạo mới để lưu context dự án |
| 2026-04-13 | Thay Listbox bằng CheckListbox (checkbox tick), đổi "Cửa hàng vật lý" → "Kho Cửa Hàng", "Kênh Shopee" → "Các kho khác"; phân loại store: chỉ store có doanh số 3T gần nhất vào Kho Cửa Hàng, còn lại (shopee/droppi/xe thuê/kho kg/không có số bán) vào Các kho khác; thêm DPI awareness cho Windows (`SetProcessDpiAwareness(2)`) | Cải thiện UX bộ lọc cửa hàng và độ nét trên màn hình HiDPI |
