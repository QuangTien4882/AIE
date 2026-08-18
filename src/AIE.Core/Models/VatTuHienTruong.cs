using System.Collections.Generic;

namespace AIE.Core.Models;

public class BangTongHopVatTu
{
    public List<VatLieuHienTruong> DanhSachVatLieu { get; set; } = new();
    public List<NhanCongHienTruong> DanhSachNhanCong { get; set; } = new();
    public List<MayThiCongHienTruong> DanhSachMay { get; set; } = new();
    
    // Giá nhiên liệu cho công trình (để tính giá máy tự động)
    public decimal GiaXang { get; set; }
    public decimal GiaDiezel { get; set; }
    public decimal GiaDien { get; set; }
    
    // Hệ số/Giá nhân công lái máy (để đơn giản thì cho nhập thẳng 1 giá trung bình hoặc chi tiết theo nhóm)
    public decimal GiaNhanCongLaiMay { get; set; }
}

public class VatTuHienTruongBase
{
    public string MaVatTu { get; set; } = string.Empty;
    public string TenVatTu { get; set; } = string.Empty;
    public string DonVi { get; set; } = string.Empty;
    public decimal TongKhoiLuong { get; set; }
    
    public decimal GiaGoc { get; set; } // Lấy từ Master Data
    public virtual decimal GiaHienTruong { get; set; } // Tính toán ra
    public decimal ThanhTien => TongKhoiLuong * GiaHienTruong;
}

public class VatLieuHienTruong : VatTuHienTruongBase
{
    public decimal CuocVanChuyen { get; set; }
    public string NguonCungCap { get; set; } = string.Empty;
    
    // Giá hiện trường của vật liệu = Giá mua + Cước vận chuyển
    public override decimal GiaHienTruong 
    {
        get => GiaGoc + CuocVanChuyen;
        set => base.GiaHienTruong = value;
    }
}

public class NhanCongHienTruong : VatTuHienTruongBase
{
    // Nhân công hiện trường thường nhập trực tiếp hoặc qua hệ số
    public decimal HeSo { get; set; } = 1;
}

public class MayThiCongHienTruong : VatTuHienTruongBase
{
    // Thông tin định mức hao phí máy lấy từ CSDL (Thông tư 37)
    public DinhMucCaMay_TT37? DinhMuc { get; set; }
    
    // Nguyên giá máy (để hiển thị cho người dùng kiểm tra)
    public decimal NguyenGia { get; set; }
    
    public decimal ChiPhiKhauHao { get; set; }
    public decimal ChiPhiSuaChua { get; set; }
    public decimal ChiPhiKhac { get; set; }
    public decimal ChiPhiNhiemLieu { get; set; }
    public decimal ChiPhiNhanCong { get; set; }
    
    // Giá hiện trường của máy = Tổng các thành phần chi phí
    public override decimal GiaHienTruong 
    {
        get => ChiPhiKhauHao + ChiPhiSuaChua + ChiPhiKhac + ChiPhiNhiemLieu + ChiPhiNhanCong;
        set => base.GiaHienTruong = value;
    }
}
