using AIE.Core.Enums;

namespace AIE.Core.Models;

/// <summary>
/// Kết quả thẩm định dự toán — tổng hợp tất cả các kiểm tra.
/// </summary>
public class KetQuaThamDinh
{
    public string TenCongTrinh { get; set; } = string.Empty;
    public string LoaiCongTrinh { get; set; } = string.Empty;
    public DateTime NgayThamDinh { get; set; } = DateTime.Now;

    /// <summary>Danh sách chi tiết từng mục kiểm tra</summary>
    public List<MucKiemTra> DanhSachKiemTra { get; set; } = [];

    /// <summary>Tổng hợp đối chiếu chi phí</summary>
    public DoiChieuChiPhi? DoiChieu { get; set; }

    public int SoLoiCritical => DanhSachKiemTra.Count(x => x.MucDo == MucDoThamDinh.KhongDat);
    public int SoLoiWarning => DanhSachKiemTra.Count(x => x.MucDo == MucDoThamDinh.CanhBao);
    public int SoMucDat => DanhSachKiemTra.Count(x => x.MucDo == MucDoThamDinh.Dat);

    /// <summary>Kết quả tổng thể</summary>
    public MucDoThamDinh KetQuaTongThe =>
        SoLoiCritical > 0 ? MucDoThamDinh.KhongDat :
        SoLoiWarning > 0 ? MucDoThamDinh.CanhBao :
        MucDoThamDinh.Dat;
}

/// <summary>
/// 1 mục kiểm tra trong kết quả thẩm định.
/// </summary>
public class MucKiemTra
{
    /// <summary>Tên mục kiểm tra (VD: "Mã hiệu AF.11112", "Tỉ lệ CPC",...)</summary>
    public string TenMuc { get; set; } = string.Empty;

    /// <summary>Nội dung chi tiết</summary>
    public string NoiDung { get; set; } = string.Empty;

    /// <summary>Giá trị trong dự toán trình</summary>
    public decimal? GiaTriTrinh { get; set; }

    /// <summary>Giá trị tính lại (đúng)</summary>
    public decimal? GiaTriTinhLai { get; set; }

    /// <summary>Chênh lệch</summary>
    public decimal? ChenhLech => GiaTriTrinh.HasValue && GiaTriTinhLai.HasValue
        ? GiaTriTrinh.Value - GiaTriTinhLai.Value
        : null;

    /// <summary>Mức độ đánh giá</summary>
    public MucDoThamDinh MucDo { get; set; }

    /// <summary>Ghi chú / lý do</summary>
    public string? GhiChu { get; set; }

    /// <summary>Vị trí cell trong Excel (VD: "D7", "E15",...)</summary>
    public string? ViTriCell { get; set; }
}

/// <summary>
/// Bảng đối chiếu chi phí tổng hợp.
/// </summary>
public class DoiChieuChiPhi
{
    public decimal GXD_Trinh { get; set; }
    public decimal GXD_TinhLai { get; set; }
    public decimal GTB_Trinh { get; set; }
    public decimal GTB_TinhLai { get; set; }
    public decimal GQLDA_Trinh { get; set; }
    public decimal GQLDA_TinhLai { get; set; }
    public decimal GTV_Trinh { get; set; }
    public decimal GTV_TinhLai { get; set; }
    public decimal GK_Trinh { get; set; }
    public decimal GK_TinhLai { get; set; }
    public decimal GDP_Trinh { get; set; }
    public decimal GDP_TinhLai { get; set; }

    public decimal Tong_Trinh => GXD_Trinh + GTB_Trinh + GQLDA_Trinh + GTV_Trinh + GK_Trinh + GDP_Trinh;
    public decimal Tong_TinhLai => GXD_TinhLai + GTB_TinhLai + GQLDA_TinhLai + GTV_TinhLai + GK_TinhLai + GDP_TinhLai;
    public decimal ChenhLech => Tong_Trinh - Tong_TinhLai;
    public decimal TiLeChenhLech => Tong_TinhLai != 0 ? ChenhLech / Tong_TinhLai * 100 : 0;
}
