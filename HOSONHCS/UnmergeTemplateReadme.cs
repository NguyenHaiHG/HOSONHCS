using System;
using System.Windows.Forms;

namespace HOSONHCS
{
    /// <summary>
    /// CÁCH SỬ DỤNG CÔNG CỤ TÁCH GỘP Ô TEMPLATE
    /// 
    /// ===== CÁCH 1: CHẠY TRỰC TIẾP =====
    /// Uncomment dòng dưới trong Program.cs Main() để chạy công cụ này:
    /// 
    ///     UnmergeTemplateForm.ShowUnmergeForm();
    ///     return;
    /// 
    /// ===== CÁCH 2: THÊM NÚT VÀO FORM1 =====
    /// Thêm nút trong Form1 và gắn sự kiện:
    /// 
    ///     btnUnmergeTemplate.Click += (s, e) => UnmergeTemplateForm.ShowUnmergeForm();
    /// 
    /// ===== CÁCH 3: GỌI TRỰC TIẾP HÀM =====
    /// Trong Form1.cs hoặc bất kỳ đâu:
    /// 
    ///     UnmergeTemplateUtil.UnmergeSXKDTemplate();
    /// 
    /// ===== SAU KHI TÁCH GỘP Ô XONG =====
    /// 1. File "01 SXKD.docx" đã được tách gộp tất cả các ô
    /// 2. File backup "01 SXKD_BACKUP_MERGED.docx" chứa file gốc (có ô gộp)
    /// 3. Có thể XÓA các file sau (không cần thiết nữa):
    ///    - UnmergeTemplateUtil.cs
    ///    - UnmergeTemplateForm.cs
    ///    - UnmergeTemplateReadme.cs (file này)
    /// </summary>
    internal static class UnmergeTemplateReadme
    {
        // File này CHỈ để hướng dẫn, KHÔNG chứa code thực thi
    }
}
