using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HOSONHCS
{
    /// <summary>
    /// Công cụ để tách gộp tất cả các ô trong file template Word
    /// Chỉ chạy một lần để sửa file template, sau đó có thể xóa class này
    /// </summary>
    public static class UnmergeTemplateUtil
    {
        /// <summary>
        /// Tách gộp tất cả các ô trong file Word template
        /// </summary>
        /// <param name="templatePath">Đường dẫn đầy đủ đến file template</param>
        public static void UnmergeCells(string templatePath)
        {
            if (!File.Exists(templatePath))
            {
                MessageBox.Show($"Không tìm thấy file: {templatePath}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Tạo bản sao file để an toàn
            string backupPath = templatePath.Replace(".docx", "_BACKUP_MERGED.docx");
            try
            {
                File.Copy(templatePath, backupPath, true);
                MessageBox.Show($"Đã tạo bản sao lưu:\n{backupPath}", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Không thể tạo bản sao lưu: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(templatePath, true))
                {
                    if (doc.MainDocumentPart == null)
                    {
                        MessageBox.Show("File không hợp lệ!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    int totalUnmerged = 0;

                    // Tách gộp ô trong tất cả các bảng
                    foreach (var table in doc.MainDocumentPart.Document.Descendants<Table>())
                    {
                        var rows = table.Elements<TableRow>().ToList();
                        
                        foreach (var row in rows)
                        {
                            foreach (var cell in row.Elements<TableCell>())
                            {
                                var tcPr = cell.GetFirstChild<TableCellProperties>();
                                if (tcPr != null)
                                {
                                    // Xóa gộp dọc (Vertical Merge)
                                    var vMerge = tcPr.GetFirstChild<VerticalMerge>();
                                    if (vMerge != null)
                                    {
                                        vMerge.Remove();
                                        totalUnmerged++;
                                    }

                                    // Xóa gộp ngang (Grid Span)
                                    var gridSpan = tcPr.GetFirstChild<GridSpan>();
                                    if (gridSpan != null)
                                    {
                                        gridSpan.Remove();
                                        totalUnmerged++;
                                    }

                                    // Xóa horizontal merge nếu có
                                    var hMerge = tcPr.GetFirstChild<HorizontalMerge>();
                                    if (hMerge != null)
                                    {
                                        hMerge.Remove();
                                        totalUnmerged++;
                                    }
                                }
                            }
                        }
                    }

                    doc.MainDocumentPart.Document.Save();

                    MessageBox.Show(
                        $"Đã tách gộp thành công!\n\n" +
                        $"File: {Path.GetFileName(templatePath)}\n" +
                        $"Số lượng merge đã xóa: {totalUnmerged}\n\n" +
                        $"Bản sao lưu (file gốc có ô gộp): {Path.GetFileName(backupPath)}",
                        "Hoàn tất",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi tách gộp ô:\n{ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Tách gộp ô cho file template "01 SXKD.docx"
        /// Tìm file trong thư mục Templates hoặc thư mục gốc
        /// </summary>
        public static void UnmergeSXKDTemplate()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string templateName = "01 SXKD.docx";

            // Thử tìm trong thư mục Templates
            string templatePath = Path.Combine(baseDir, "Templates", templateName);
            
            if (!File.Exists(templatePath))
            {
                // Thử tìm trong thư mục gốc
                templatePath = Path.Combine(baseDir, templateName);
            }

            if (!File.Exists(templatePath))
            {
                MessageBox.Show(
                    $"Không tìm thấy file template '{templateName}'\n\n" +
                    $"Đã tìm trong:\n" +
                    $"- {Path.Combine(baseDir, "Templates", templateName)}\n" +
                    $"- {Path.Combine(baseDir, templateName)}",
                    "Lỗi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            UnmergeCells(templatePath);
        }
    }
}
