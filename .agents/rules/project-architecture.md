# AIE — Excel Add-In Dự Toán & Thẩm Định Xây Dựng

## 1. Tổng Quan Dự Án

**AIE (AI Estimation)** là một Excel Add-In (.NET Framework 4.8 + ExcelDna) phục vụ hai nghiệp vụ chính:
1. **Thẩm định dự toán**: Kiểm tra, đối chiếu dự toán xây dựng với định mức nhà nước
2. **Lập dự toán mới**: Tạo dự toán xây dựng công trình từ đầu, tính giá hiện trường, tổng hợp kinh phí, xuất bảng biểu Excel

**Địa bàn áp dụng chính**: Đà Nẵng (Vùng II mặc định)

**Văn bản pháp lý**:
- **Thông tư 36/2026/TT-BXD**: Cấu trúc chi phí xây dựng (GXD = T + GT + TL + GTGT + LT)
- **Thông tư 37/2026/TT-BXD**: Định mức ca máy thiết bị thi công
- **Thông tư 38/2026/TT-BXD**: Định mức chi phí QLDA & Tư vấn đầu tư xây dựng

---

## 2. Kiến Trúc Hệ Thống

### 2.1 Solution Structure (AIE.slnx)
```
src/
├── AIE.Core/          → Domain models, enums, business logic, constants
├── AIE.Data/          → SQLite database, repositories (Dapper), import/export
├── AIE.ExcelAddIn/    → UI (WinForms), Ribbon, Services, Helpers — ENTRY POINT
└── AIE.Reports/       → (Placeholder, chưa triển khai)

tests/
├── AIE.Core.Tests/
└── AIE.Data.Tests/
```

### 2.2 Technology Stack
| Thành phần | Công nghệ |
|---|---|
| Framework | .NET Framework 4.8 |
| Excel Integration | ExcelDna.AddIn 1.8.0 + ExcelDna.Interop 16.0.0 |
| Database | SQLite (System.Data.SQLite.Core 1.0.118) |
| ORM | Dapper 2.1.79 |
| Excel Read/Write | ClosedXML 0.95.4 (cho export) + Microsoft.Office.Interop.Excel (cho runtime) |
| Serialization | Newtonsoft.Json 13.0.4 (lưu file .dt) |
| UI | System.Windows.Forms (WinForms modals) |

### 2.3 Dependency Flow
```
AIE.Core ← AIE.Data ← AIE.ExcelAddIn
                    ← AIE.Reports (placeholder)
```

---

## 3. Ribbon UI — Các Chức Năng Chính

Tab Ribbon: **"AIE Dự Toán"** gồm 4 nhóm:

### 3.1 Nhóm "Thiết lập dữ liệu"
| Nút | Handler | Chức năng |
|---|---|---|
| Nhập Database | `OnImportClicked` | Import file Excel định mức (VL, NC, Máy, Hao phí) vào SQLite |
| Tra cứu định mức | `OnTraCuuClicked` | Form tra cứu floating (không block Excel) |
| Tạo File mẫu | `OnTaoTemplateClicked` | Tạo 6 file Excel template |
| Trích xuất Đơn Giá | `OnXuatDonGiaClicked` | Export danh mục đơn giá từ HaoPhi |
| Quản lý Đơn giá | `OnQuanLyDonGiaClicked` | CRUD Vật liệu, Nhân công, Máy thi công |
| Xóa Database | `OnDeleteDbClicked` | Xóa toàn bộ hoặc từng bảng SQLite |

### 3.2 Nhóm "Thẩm định dự toán"
| Nút | Handler | Chức năng |
|---|---|---|
| Đơn giá thẩm định | `OnDonGiaThamDinhClicked` | Đọc dự toán → map định mức → tạo Bộ Đơn Giá thẩm định |
| Mở Bộ Đơn giá | `OnMoDonGiaThamDinhClicked` | Mở lại Bộ Đơn Giá đã lưu |
| Kiểm tra | `OnKiemTraClicked` | Thẩm định: so sánh DT trình vs Định mức chuẩn → sheet KQ_ThamDinh |
| Xuất Báo cáo | `OnBaoCaoTDClicked` | (Đang phát triển) |

### 3.3 Nhóm "Lập dự toán"
| Nút | Handler | Chức năng |
|---|---|---|
| Tạo Dự toán mới | `OnTaoDuToanMoiClicked` | Form nhập thông tin dự án → tạo BOQ trên Excel |
| Gọi Đơn giá | `OnGoiDonGiaClicked` | Chọn vùng mã hiệu trên Excel → tự fill tên, đơn vị từ DB |
| Giá VL, NC, MTC | `OnTinhGiaHienTruongClicked` | **Form chính**: Phân tích vật tư → Tính giá hiện trường → Bốc xếp, Vận chuyển |
| Tổng hợp kinh phí | `OnTongHopKinhPhiClicked` | **Form chính**: Tab 1 Chi phí XD, Tab 2 TMĐT & Tổng hợp DT |

### 3.4 Nhóm "File Dự toán"
| Nút | Handler | Chức năng |
|---|---|---|
| Lưu Dự toán | `OnLuuDuToanClicked` | Serialize `DuToan` object → file `.dt` (JSON) |
| Mở Dự toán | `OnMoDuToanClicked` | Deserialize `.dt` → restore BOQ + bảng biểu Excel |

---

## 4. Domain Models (AIE.Core)

### 4.1 Model Chính: DuToan
```
DuToan
├── TenDuAn, TenCongTrinh, LoaiCongTrinh, CapCongTrinh
├── SoBuocThietKe (0/1/2/3)
├── VungApDung (Vung enum: VungII=2, VungIII=3, VungIV=4, CuLaoCham=5)
├── BoDonGiaId (int?) → liên kết Bộ Đơn Giá trong DB
├── DanhSachHangMuc[] → HangMuc → DongDuToan[]
│   └── DongDuToan: STT, MaHieu, TenCongTac, DonVi, KhoiLuong
│       ├── DonGiaVL, DonGiaNC, DonGiaMay → DonGiaTongHop
│       ├── ThanhTien, ThanhTienVL, ThanhTienNC, ThanhTienMay
│       └── DanhSachHaoPhi[] (HaoPhi: MaHieuHP, TenHP, DonVi, DinhMuc, HeSo)
├── BangTongHop (BangTongHopVatTu)
│   ├── DanhSachVatLieu[] (VatLieuHienTruong)
│   │   └── GiaGoc + ChiPhiBocXep + CuocVCOTo + CuocVCBo = GiaHienTruong
│   ├── DanhSachNhanCong[] (NhanCongHienTruong)
│   └── DanhSachMay[] (MayThiCongHienTruong)
│       └── ChiPhiKhauHao + SuaChua + Khac + NhienLieu + NhanCong = GiaHienTruong
├── ChiPhiXD (ChiPhiXayDung)
│   └── VL, NC, M → T → CPC, TT → GT → TL → G → GTGT → Gxd → LT → GXD
├── BangKinhPhi (BangTongHopKinhPhiModel)
│   └── Items[] (ChiPhiKinhPhiItem) — QLDA, Tư vấn, Chi phí khác, Dự phòng
└── ChiPhiThietBi, ChiPhiQLDA, ChiPhiTuVan, ChiPhiKhac, ChiPhiDuPhong
```

### 4.2 Công Thức Chi Phí Xây Dựng (TT36)
```
T  = VL + NC + M                    (Chi phí trực tiếp)
CPC = T × %CPC (hoặc NC × %CPC)    (Chi phí chung)
TT  = T × %TT                       (CP không XĐ được KL)
GT  = CPC + TT                      (Chi phí gián tiếp)
TL  = (T + GT) × %TNCTTT            (Thu nhập chịu thuế tính trước)
G   = T + GT + TL                   (CP XD trước thuế)
GTGT = G × %GTGT                    (Thuế GTGT, thường 8%)
Gxd = G + GTGT                      (CP XD sau thuế)
LT  = Gxd × %NhaTam                 (Chi phí nhà tạm)
GXD = Gxd + LT                      (Tổng chi phí xây dựng)
```

### 4.3 Enums Quan Trọng
- **LoaiCongTrinh**: DanDung, CongNghiep, GiaoThong, ThuyLoi, HaTangKyThuat
- **Vung**: VungII(2), VungIII(3), VungIV(4), CuLaoCham(5)
- **LoaiHaoPhi**: VL(0), NC(1), MAY(2)
- **LoaiTiLe**: CPC, TT, TNCTTT, GTGT, NhaTam
- **LoaiNhanCong**: XayDung, LaiXe, VanHanh
- **NhomChiPhi**: BoiThuong_TDC, ChiPhiXayDung, ChiPhiThietBi, QuanLyDuAn, TuVanDauTuXD, ChiPhiKhac, ChiPhiDuPhong

---

## 5. Database (SQLite)

**File**: `AieData.sqlite` (cùng thư mục Add-In)

### Bảng chính:
| Bảng | Mô tả |
|---|---|
| CongTacXayDung | Danh mục công tác XD (MaHieu, TenCongTac, DonVi) |
| HaoPhi | Hao phí của công tác (VL/NC/MAY, MaHieuHP, DinhMuc, HeSo) — FK → CongTacXayDung |
| VatLieu | Giá vật liệu địa phương (MaVL, TenVL, DonGia) |
| NhanCong | Giá nhân công theo nhóm 1-6 (MaNC, TenNC, Nhom, DonGia) |
| MayThiCong | Giá ca máy (MaMay, TenMay, DonGia) |
| TiLePhanTram | Tỉ lệ % theo loại CT (CPC, TT, TNCTTT, GTGT, NhaTam) |
| DinhMucQLDA | Định mức % QLDA theo quy mô |
| DinhMucTuVan | Định mức % tư vấn theo quy mô & loại CT |

### Bảng bổ sung (tạo trong DatabaseManager):
| Bảng | Mô tả |
|---|---|
| BoDonGia | Bộ đơn giá thẩm định (TenBo, GiaXang, GiaDiezel, GiaDien) |
| BoDonGia_VatLieu | Giá VL trong bộ đơn giá |
| BoDonGia_NhanCong | Giá NC trong bộ đơn giá |
| BoDonGia_MayThiCong | Giá Máy trong bộ đơn giá |
| DinhMucCaMay | Định mức ca máy TT37 (KhauHao, SuaChua, NhienLieu, NhanCong) |
| TuDienDongNghia | Từ điển đồng nghĩa cho matching tên công tác |
| DinhMucBocXep | Định mức bốc xếp hàng hóa |
| DinhMucVanChuyen | Định mức vận chuyển (ô tô, bộ) |

### Repositories:
- `CongTacRepository` — CRUD công tác + hao phí
- `VatLieuRepository`, `NhanCongRepository`, `MayThiCongRepository` — CRUD giá địa phương
- `BoDonGiaRepository` — CRUD bộ đơn giá thẩm định
- `DinhMucCaMayRepository` — Tra cứu định mức ca máy TT37
- `DinhMucCPCRepository`, `DinhMucTTRepository` — Tra cứu tỉ lệ %

---

## 6. Services (AIE.ExcelAddIn)

### 6.1 Luồng Lập Dự Toán
```
TaoDuToanMoiForm → LapDuToanExcelService.WriteBOQToActiveSheet()
    ↓
GoiDonGia → Fill tên/đơn vị từ DB theo mã hiệu
    ↓
TinhGiaHienTruongForm (Form chính)
├── Tab VL: PhanTichVatTuService.PhanTich() → BangTongHopVatTu
│   ├── TinhChiPhiBocXepForm → BocXepStorage
│   ├── TinhCuocVCOToForm → VanChuyenStorage
│   └── TinhCuocVCBoForm → VanChuyenStorage
├── Tab NC: Nhập/chỉnh giá nhân công theo nhóm
├── Tab MTC: Tính giá ca máy theo TT37 (DinhMucCaMay_TT37)
└── PhanTichDonGiaService.TinhDonGiaChiTiet() → DonGiaVL/NC/May cho từng công tác
    ↓
TongHopKinhPhiForm (Form chính)
├── Tab 1: Chi phí XD (ChiPhiXayDungCalc) → T, GT, TL, G, GTGT, LT, GXD
├── Tab 2: Tổng hợp DT & TMĐT (BangTongHopKinhPhiModel)
│   └── QLDA, Tư vấn (TK, Giám sát, Thẩm tra...), Chi phí khác, Dự phòng
│   └── Nội suy theo DinhMucTT38Database + DinhMucChiPhiKhacBTC
└── Xuất Excel: XuatBangBieuService.Xuat7BangBieu()
```

### 6.2 Luồng Thẩm Định
```
ThamDinhConfigForm → Cấu hình map cột Excel
    ↓
DuToanExcelReader.Read() → Danh sách công tác từ file Excel
    ↓
ThamDinhEngine.KiemTra() → So sánh với định mức chuẩn
    ↓
ThamDinhExcelWriter.ExportResult() → Sheet KQ_ThamDinh
```

### 6.3 Services Chi Tiết
| Service | Vai trò |
|---|---|
| `LapDuToanExcelService` | Đọc/ghi BOQ từ/lên Excel active sheet |
| `PhanTichVatTuService` | Phân tích hao phí → tổng hợp VL, NC, Máy |
| `PhanTichDonGiaService` | Tính đơn giá chi tiết (HaoPhi × Giá = ĐG VL/NC/M) |
| `XuatBangBieuService` | **File lớn nhất** (~127KB) — Xuất 7+ bảng biểu Excel |
| `BocXepStorage` | Lưu/đọc cấu hình bốc xếp cho từng VL |
| `VanChuyenStorage` | Lưu/đọc cấu hình vận chuyển (ô tô + bộ) |
| `DuToanFileService` | Save/Load file `.dt` (JSON serialize DuToan) |
| `ThamDinhEngine` | Engine thẩm định: map mã hiệu, so sánh, trich xuất vật tư |
| `ThamDinhExcelWriter` | Xuất kết quả thẩm định ra Excel |
| `DuToanExcelReader` | Đọc dự toán từ Excel theo cấu hình map cột |

### 6.4 Các Sheet Excel Được Xuất (XuatBangBieuService)
1. **DuToan** — Bảng dự toán chi tiết (BOQ)
2. **TH_ChiPhiXD** — Tổng hợp chi phí xây dựng
3. **THVT** — Tổng hợp vật tư
4. **PTDG** — Phân tích đơn giá chi tiết
5. **TongHopDuToan** — Tổng hợp dự toán
6. **TongMucDauTu** — Tổng mức đầu tư
7. **Bia** — Trang bìa

---

## 7. Forms UI (WinForms)

| Form | Kích thước (bytes) | Chức năng |
|---|---|---|
| `TongHopKinhPhiForm` | 116KB | TabControl: Tab1=Chi phí XD, Tab2=TMĐT/DT |
| `ThamDinhDonGiaForm` | 69KB | Grid VL/NC/Máy, lưu Bộ Đơn Giá |
| `TinhGiaHienTruongForm` | 62KB | 3 tab VL/NC/MTC + tính giá hiện trường |
| `TinhCuocVCOToForm` | 57KB | Tính cước vận chuyển ô tô |
| `TinhCuocVCBoForm` | 43KB | Tính cước vận chuyển bộ (thồ) |
| `QuanLyDonGiaForm` | 37KB | Tab VL/NC/Máy CRUD |
| `ThamDinhConfigForm` | 28KB | Cấu hình map cột Excel |
| `ChonChiPhiThuVienForm` | 24KB | Chọn chi phí từ thư viện mẫu |
| `TinhChiPhiBocXepForm` | 22KB | Tính chi phí bốc xếp VL |
| `TraCuuDinhMucForm` | 20KB | Tra cứu định mức floating |
| `TinhTongHopForm` | 16KB | Tổng hợp (delegate to TongHopKinhPhiForm) |
| `TaoDuToanMoiForm` | 15KB | Form tạo dự toán mới |
| `ChonBangXuatExcelDialog` | 15KB | Chọn bảng biểu xuất Excel |
| `NhapDinhMucCaMayForm` | 15KB | Nhập định mức ca máy TT37 |
| `ImportDatabaseForm` | 10KB | Import Excel → SQLite |
| `ChonBoDonGiaForm` | 9KB | Chọn bộ đơn giá đã lưu |
| `DeleteDatabaseForm` | 4KB | Xóa database |
| `LoadingForm` | 1KB | Loading spinner |

### Helpers:
- `UIHelper` (13KB) — Tiện ích UI chung (style grid, format số, ...)
- `GridHelper` (6KB) — DataGridView utilities

---

## 8. Dữ Liệu Định Mức Hardcoded (AIE.Core)

### 8.1 DinhMucTT38Database — Định mức TT38/2026
- QLDA: 5 loại CT × 12 mốc quy mô (10-30,000 tỷ)
- Báo cáo KTKT: 5 loại CT × 5 mốc (1-15 tỷ)
- Lập FS: 5 loại CT × 10 mốc
- Thiết kế BVTC: 5 loại CT × 10 mốc (Cấp III chuẩn, hệ số cấp: ĐB=1.45, I=1.25, II=1.12, III=1.0, IV=0.85)
- Thẩm tra TK, Thẩm tra DT: 5 loại CT × 10 mốc
- Giám sát thi công: 5 loại CT × 10 mốc
- Lập HSMT, Đánh giá HSDT: 7 mốc

### 8.2 DinhMucChiPhiKhacBTC — Phí BTC
- Thẩm tra quyết toán, Kiểm toán độc lập
- Thẩm định dự án, Thẩm định TK, Thẩm định DT

### 8.3 Nội suy
- Công thức: `Nt = Nb - ((Nb - Na) / (Ga - Gb)) × (Gt - Gb)`
- Dùng chung `InterpolationHelper` và `DinhMucChiPhiKhacBTC.NoiSuy()`

---

## 9. Quy Tắc Làm Tròn (RoundingRules)

| Loại | Số thập phân |
|---|---|
| Thành tiền / Chi phí | 0 (phần nguyên) |
| Đơn giá (VL/NC/Máy) | 0 (phần nguyên) |
| Khối lượng | 3 |
| Hao phí định mức | 5 |
| Tỉ lệ % | 3 |
| Tổng mức đầu tư | Làm tròn đến hàng nghìn đồng |

---

## 10. Trạng Thái In-Memory

- `AieRibbon.CurrentDuToan` — DuToan object đang làm việc (static, shared giữa các form)
- `AieRibbon.CurrentFilePath` — Đường dẫn file .dt hiện tại (null nếu chưa lưu)
- Khi chuyển giữa các form, dữ liệu được bảo toàn qua `CurrentDuToan`
- File `.dt` = JSON serialization của toàn bộ `DuToan` object

---

## 11. Conventions & Patterns

### Naming
- Models: Tiếng Việt không dấu (DuToan, HangMuc, DongDuToan, HaoPhi)
- Enums: Tiếng Việt không dấu (LoaiCongTrinh, Vung, LoaiHaoPhi)
- Services: `[Chức năng]Service` hoặc `[Chức năng]Calc`
- Repositories: `[Entity]Repository`

### Error Handling
- Tất cả Ribbon handlers đều có try-catch → MessageBox.Show()
- Forms tự quản lý exception trong event handlers

### UI Pattern
- Modal dialogs (ShowDialog) cho hầu hết forms
- Show() chỉ cho TraCuuDinhMucForm (floating)
- LoadingForm hiển thị trước khi mở form nặng

### Build & Run
- Build: `dotnet build AIE.slnx --configuration Debug`
- Add-In tự load vào Excel qua ExcelDna
- File output: `bin/Debug/net48/AIE.ExcelAddIn-AddIn64.xll`
