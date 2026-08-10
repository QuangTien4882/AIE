namespace AIE.Core.Enums;

/// <summary>
/// Mức độ đánh giá kết quả thẩm định.
/// </summary>
public enum MucDoThamDinh
{
    /// <summary>🟢 Đạt — đúng quy định</summary>
    Dat,
    /// <summary>🟡 Cảnh báo — sai lệch nhỏ, cần xem xét</summary>
    CanhBao,
    /// <summary>🔴 Không đạt — sai lệch nghiêm trọng</summary>
    KhongDat,
    /// <summary>⚪ Bỏ qua — không đủ dữ liệu để kiểm tra</summary>
    BoQua
}
