using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HOSONHCS
{
    public partial class Form1
    {
        private string XuatHopDongGqvl(KhachHangGqvl item)
        {
            string templatePath = ResolveTemplatePath("HD.docx");
            string outputFolder = GetGqvlOutputFolder(item);
            Directory.CreateDirectory(outputFolder);

            string outputPath = Path.Combine(outputFolder, "HD_" + MakeFileSystemSafe(item.Tenkh) + ".docx");
            outputPath = UniqueGqvlFilePath(outputPath);
            File.Copy(templatePath, outputPath, true);

            DienPlaceholderGqvl(outputPath, item);
            return outputPath;
        }

        private string GetGqvlOutputFolder(KhachHangGqvl item)
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string root = Path.Combine(desktop, "Hồ sơ NHCS", "GQVL");
            string folder = MakeFileSystemSafe(item.Tenkh) + "_" + DateTime.Now.ToString("dd-MM-yyyy_HHmmss");
            return Path.Combine(root, folder);
        }

        private static string UniqueGqvlFilePath(string path)
        {
            if (!File.Exists(path)) return path;
            string dir = Path.GetDirectoryName(path);
            string name = Path.GetFileNameWithoutExtension(path);
            string ext = Path.GetExtension(path);
            int i = 2;
            string candidate;
            do
            {
                candidate = Path.Combine(dir, name + "_" + i + ext);
                i++;
            }
            while (File.Exists(candidate));
            return candidate;
        }

        private void DienPlaceholderGqvl(string docPath, KhachHangGqvl item)
        {
            var replacements = TaoPlaceholderGqvl(item);

            using (var doc = WordprocessingDocument.Open(docPath, true))
            {
                var mainPart = doc.MainDocumentPart;
                if (mainPart == null || mainPart.Document == null) return;

                ProcessConditionalBlock(mainPart, "chuyenkhoan_block", item.ChuyenKhoan && !string.IsNullOrWhiteSpace(item.Stk));
                ProcessConditionalBlock(mainPart, "uyquyen_block", !item.GiamDoc);
                ReplaceBlock(mainPart, "phanky_block", replacements["{{phanky_block}}"], "phankky_block");

                ReplacePlaceholdersInPart(mainPart, replacements);
                foreach (var header in mainPart.HeaderParts) ReplacePlaceholdersInPart(header, replacements);
                foreach (var footer in mainPart.FooterParts) ReplacePlaceholdersInPart(footer, replacements);
                IndentThoihanVayParagraph(mainPart);

                mainPart.Document.Save();
            }
        }

        private Dictionary<string, string> TaoPlaceholderGqvl(KhachHangGqvl item)
        {
            DateTime hanTraNo = TinhToanGqvl.TinhHanTraNo(item.Ngaygiaingan, item.Thoihanvay);

            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "{{sohd}}", item.Sohd ?? "" },
                { "{{sotienvaygqvl}}", item.Sotienvay ?? "" },
                { "{{bangchugqvl}}", DocTienBangChuGqvl(item.Sotienvay) },
                { "{{mucdichgqvl}}", item.Mucdich ?? "" },
                { "{{ngayhd}}", TinhToanGqvl.FormatDate(item.Ngayhopdong) },
                { "{{lsgqvl}}", item.Laisuat ?? "" },
                { "{{ngayvaygqvl}}", TinhToanGqvl.FormatDate(item.Ngaygiaingan) },
                { "{{stkgqvl}}", item.Stk ?? "" },
                { "{{dcphuongangqvl}}", item.DcPhuongAn ?? "" },
                { "{{tenkhgqvl}}", item.Tenkh ?? "" },
                { "{{ngaysinhkhgqvl}}", TinhToanGqvl.FormatDate(item.Ngaysinh) },
                { "{{sdtkhgqvl}}", item.SdtKh ?? "" },
                { "{{cccdgqvl}}", item.Cccd ?? "" },
                { "{{ngaycapccgqvl}}", TinhToanGqvl.FormatDate(item.NgaycapCccd) },
                { "{{thoihanccgqvl}}", string.IsNullOrWhiteSpace(item.ThoihanCccdText) ? TinhToanGqvl.FormatDate(item.ThoihanCccd) : item.ThoihanCccdText },
                { "{{noicapccgqvl}}", item.NoicapCccd ?? "" },
                { "{{diachikhgqvl}}", item.DiachiKh ?? "" },
                { "{{pgdgqvl}}", item.Pgd ?? "" },
                { "{{tenldgqvl}}", item.TenLanhDao ?? "" },
                { "{{sdtpgdgqvl}}", item.SdtPgd ?? "" },
                { "{{diachigqvl}}", item.DiachiPgd ?? "" },
                { "{{danhxunggqvl}}", item.Ong ? "Ông" : "Bà" },
                { "{{chucvugqvl}}", item.GiamDoc ? "Giám đốc" : "Phó giám đốc" },
                { "{{souqgqvl}}", item.SoUyQuyen ?? "" },
                { "{{ngayuqgqvl}}", TinhToanGqvl.FormatDate(item.NgayUyQuyen) },
                { "{{qhgqvl}}", TinhToanGqvl.TinhLaiQuaHan(item.Laisuat) },
                { "{{phanky_block}}", TinhToanGqvl.TaoPhanKyBlock(item.Ngaygiaingan, item.Sotienvay, item.Thoihanvay, item.Phanky) },
                { "{{namht}}", DateTime.Now.Year.ToString() },
                { "{{hantranogqvl}}", TinhToanGqvl.FormatDate(hanTraNo) },
                { "{{thoihanvaygqvl}}", item.Thoihanvay ?? "" },
                { "{{tm}}", item.ChuyenKhoan ? "☐" : "☑" },
                { "{{ck}}", item.ChuyenKhoan ? "☑" : "☐" }
            };
        }

        private static void ProcessConditionalBlock(OpenXmlPart part, string blockName, bool show)
        {
            if (show)
            {
                RemoveBlockMarkersOnly(part, blockName);
                return;
            }

            if (RemoveConditionalBlock(part, blockName))
                return;

            foreach (var paragraph in part.RootElement.Descendants<Paragraph>().ToList())
            {
                string text = GetParagraphTextGqvl(paragraph);
                if (text.IndexOf("{{" + blockName + "}}", StringComparison.OrdinalIgnoreCase) < 0 &&
                    text.IndexOf("{{/" + blockName + "}}", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                string pattern = Regex.Escape("{{" + blockName + "}}") + "(.*?)" + Regex.Escape("{{/" + blockName + "}}");
                string replaced = Regex.Replace(text, pattern, "", RegexOptions.IgnoreCase | RegexOptions.Singleline);

                WriteParagraphTextGqvl(paragraph, replaced.Trim());
            }
            part.RootElement.Save();
        }

        private static void RemoveBlockMarkersOnly(OpenXmlPart part, string blockName)
        {
            bool changed = false;
            string startMarker = "{{" + blockName + "}}";
            string endMarker = "{{/" + blockName + "}}";

            foreach (var paragraph in part.RootElement.Descendants<Paragraph>().ToList())
            {
                string paragraphText = GetParagraphTextGqvl(paragraph);
                string trimmed = paragraphText.Trim();

                if (string.Equals(trimmed, startMarker, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(trimmed, endMarker, StringComparison.OrdinalIgnoreCase))
                {
                    paragraph.Remove();
                    changed = true;
                    continue;
                }

                if (paragraphText.IndexOf(startMarker, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    paragraphText.IndexOf(endMarker, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    string replacedParagraph = ReplaceIgnoreCaseGqvl(paragraphText, startMarker, string.Empty);
                    replacedParagraph = ReplaceIgnoreCaseGqvl(replacedParagraph, endMarker, string.Empty);
                    WriteParagraphTextGqvl(paragraph, replacedParagraph.Trim());
                    changed = true;
                }
            }

            foreach (var text in part.RootElement.Descendants<Text>())
            {
                if (string.IsNullOrEmpty(text.Text)) continue;
                string replaced = ReplaceIgnoreCaseGqvl(text.Text, startMarker, string.Empty);
                replaced = ReplaceIgnoreCaseGqvl(replaced, endMarker, string.Empty);

                if (!string.Equals(text.Text, replaced, StringComparison.Ordinal))
                {
                    text.Text = replaced;
                    text.Space = SpaceProcessingModeValues.Preserve;
                    changed = true;
                }
            }

            if (changed) part.RootElement.Save();
        }

        private static bool RemoveConditionalBlock(OpenXmlPart part, string blockName)
        {
            bool changed = false;
            var paragraphs = part.RootElement.Descendants<Paragraph>().ToList();
            string startMarker = "{{" + blockName + "}}";
            string endMarker = "{{/" + blockName + "}}";

            for (int i = 0; i < paragraphs.Count; i++)
            {
                string startText = GetParagraphTextGqvl(paragraphs[i]);
                int startIndex = startText.IndexOf(startMarker, StringComparison.OrdinalIgnoreCase);
                if (startIndex < 0) continue;

                int endParagraphIndex = -1;
                int endIndex = -1;
                for (int j = i; j < paragraphs.Count; j++)
                {
                    string currentText = GetParagraphTextGqvl(paragraphs[j]);
                    endIndex = currentText.IndexOf(endMarker, StringComparison.OrdinalIgnoreCase);
                    if (endIndex >= 0)
                    {
                        endParagraphIndex = j;
                        break;
                    }
                }

                if (endParagraphIndex < 0) continue;

                if (endParagraphIndex == i)
                {
                    string pattern = Regex.Escape(startMarker) + ".*?" + Regex.Escape(endMarker);
                    string replaced = Regex.Replace(startText, pattern, string.Empty, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    if (string.IsNullOrWhiteSpace(replaced))
                        paragraphs[i].Remove();
                    else
                        WriteParagraphTextGqvl(paragraphs[i], replaced.Trim());
                }
                else
                {
                    string prefix = startText.Substring(0, startIndex).TrimEnd();
                    string endText = GetParagraphTextGqvl(paragraphs[endParagraphIndex]);
                    string suffix = endText.Substring(endIndex + endMarker.Length).TrimStart();

                    for (int k = endParagraphIndex; k > i; k--)
                        paragraphs[k].Remove();

                    string remaining = (prefix + suffix).Trim();
                    if (string.IsNullOrWhiteSpace(remaining))
                        paragraphs[i].Remove();
                    else
                        WriteParagraphTextGqvl(paragraphs[i], remaining);
                }

                changed = true;
            }

            if (changed) part.RootElement.Save();
            return changed;
        }

        private static void ReplaceBlock(OpenXmlPart part, string blockName, string value, params string[] endAliases)
        {
            if (ProcessMultiParagraphBlock(part, blockName, value ?? string.Empty, false, endAliases))
                return;

            foreach (var paragraph in part.RootElement.Descendants<Paragraph>().ToList())
            {
                string text = GetParagraphTextGqvl(paragraph);
                if (text.IndexOf("{{" + blockName + "}}", StringComparison.OrdinalIgnoreCase) < 0 &&
                    !ContainsAnyEndMarker(text, blockName, endAliases))
                {
                    continue;
                }

                string pattern = Regex.Escape("{{" + blockName + "}}") + ".*?" + BuildEndMarkerPattern(blockName, endAliases);
                string replaced = Regex.Replace(text, pattern, value ?? "", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                WriteParagraphTextGqvl(paragraph, replaced.Trim());
                SetParagraphLeft(paragraph);
            }
            part.RootElement.Save();
        }

        private static bool ProcessMultiParagraphBlock(OpenXmlPart part, string blockName, string replacement, bool keepOriginalContent, string[] endAliases)
        {
            bool changed = false;
            var paragraphs = part.RootElement.Descendants<Paragraph>().ToList();

            for (int i = 0; i < paragraphs.Count; i++)
            {
                string startText = GetParagraphTextGqvl(paragraphs[i]);
                int startIndex = startText.IndexOf("{{" + blockName + "}}", StringComparison.OrdinalIgnoreCase);
                if (startIndex < 0) continue;

                int endParagraphIndex = -1;
                Match endMatch = Match.Empty;
                for (int j = i; j < paragraphs.Count; j++)
                {
                    string currentText = GetParagraphTextGqvl(paragraphs[j]);
                    endMatch = Regex.Match(currentText, BuildEndMarkerPattern(blockName, endAliases), RegexOptions.IgnoreCase);
                    if (endMatch.Success)
                    {
                        endParagraphIndex = j;
                        break;
                    }
                }

                if (endParagraphIndex < 0) continue;

                string prefix = startText.Substring(0, startIndex);
                string suffix = string.Empty;
                var content = new List<string>();

                if (keepOriginalContent)
                {
                    string firstAfterStart = startText.Substring(startIndex + ("{{" + blockName + "}}").Length);
                    if (endParagraphIndex == i)
                    {
                        content.Add(firstAfterStart.Substring(0, endMatch.Index - (startIndex + ("{{" + blockName + "}}").Length)));
                    }
                    else
                    {
                        content.Add(firstAfterStart);
                        for (int k = i + 1; k < endParagraphIndex; k++)
                            content.Add(GetParagraphTextGqvl(paragraphs[k]));
                        string endText = GetParagraphTextGqvl(paragraphs[endParagraphIndex]);
                        content.Add(endText.Substring(0, endMatch.Index));
                    }
                }

                string endTextFull = GetParagraphTextGqvl(paragraphs[endParagraphIndex]);
                suffix = endTextFull.Substring(endMatch.Index + endMatch.Length);
                string body = keepOriginalContent ? string.Join(Environment.NewLine, content).Trim() : (replacement ?? string.Empty);

                WriteParagraphTextGqvl(paragraphs[i], (prefix + body + suffix).Trim());
                SetParagraphLeft(paragraphs[i]);

                for (int k = endParagraphIndex; k > i; k--)
                    paragraphs[k].Remove();

                changed = true;
            }

            if (changed) part.RootElement.Save();
            return changed;
        }

        private static void ReplacePlaceholdersInPart(OpenXmlPart part, Dictionary<string, string> replacements)
        {
            foreach (var textNode in part.RootElement.Descendants<Text>())
            {
                if (string.IsNullOrEmpty(textNode.Text)) continue;

                string replacedText = textNode.Text;
                foreach (var kv in replacements)
                    replacedText = ReplaceIgnoreCaseGqvl(replacedText, kv.Key, kv.Value ?? "");

                if (!string.Equals(textNode.Text, replacedText, StringComparison.Ordinal))
                {
                    textNode.Text = replacedText;
                    textNode.Space = SpaceProcessingModeValues.Preserve;
                }
            }

            foreach (var paragraph in part.RootElement.Descendants<Paragraph>().ToList())
            {
                string text = GetParagraphTextGqvl(paragraph);
                if (text.IndexOf("{{", StringComparison.OrdinalIgnoreCase) < 0) continue;
                if (text.IndexOf("{{cccdgqvl}}", StringComparison.OrdinalIgnoreCase) >= 0) continue;

                string replaced = text;
                foreach (var kv in replacements)
                    replaced = ReplaceIgnoreCaseGqvl(replaced, kv.Key, kv.Value ?? "");

                if (!string.Equals(text, replaced, StringComparison.Ordinal))
                    WriteParagraphTextGqvl(paragraph, replaced);
            }
            part.RootElement.Save();
        }

        private static void IndentThoihanVayParagraph(OpenXmlPart part)
        {
            foreach (var paragraph in part.RootElement.Descendants<Paragraph>())
            {
                string text = GetParagraphTextGqvl(paragraph);
                if (text.IndexOf("Thời hạn cho vay", StringComparison.OrdinalIgnoreCase) < 0 &&
                    text.IndexOf("hạn trả nợ cuối cùng", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                SetParagraphLeft(paragraph);
                var pPr = paragraph.GetFirstChild<ParagraphProperties>();
                var indentation = pPr.GetFirstChild<Indentation>();
                if (indentation == null)
                {
                    indentation = new Indentation();
                    pPr.Append(indentation);
                }

                // 720 twips ~= 0.5 inch, tương đương một tab mặc định trong Word.
                indentation.Left = "720";
            }
            part.RootElement.Save();
        }

        private static string GetParagraphTextGqvl(Paragraph paragraph)
        {
            return string.Concat(paragraph.Descendants<Text>().Select(t => t.Text ?? string.Empty));
        }

        private static void WriteParagraphTextGqvl(Paragraph paragraph, string text)
        {
            RunProperties runProperties = paragraph.Descendants<RunProperties>().FirstOrDefault()?.CloneNode(true) as RunProperties;
            paragraph.RemoveAllChildren<Run>();

            var run = new Run();
            if (runProperties != null) run.Append(runProperties);

            string[] lines = (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                if (i > 0) run.Append(new Break());
                AppendTextWithTabs(run, lines[i]);
            }
            paragraph.Append(run);
        }

        private static void AppendTextWithTabs(Run run, string text)
        {
            string[] parts = (text ?? string.Empty).Split('\t');
            for (int i = 0; i < parts.Length; i++)
            {
                if (i > 0) run.Append(new TabChar());
                run.Append(new Text(parts[i]) { Space = SpaceProcessingModeValues.Preserve });
            }
        }

        private static void SetParagraphLeft(Paragraph paragraph)
        {
            var pPr = paragraph.GetFirstChild<ParagraphProperties>();
            if (pPr == null)
            {
                pPr = new ParagraphProperties();
                paragraph.InsertAt(pPr, 0);
            }

            var jc = pPr.GetFirstChild<Justification>();
            if (jc == null)
            {
                jc = new Justification();
                pPr.Append(jc);
            }
            jc.Val = JustificationValues.Left;
        }

        private string DocTienBangChuGqvl(string soTienText)
        {
            long value = TinhToanGqvl.ParseMoney(soTienText);
            if (value <= 0) return string.Empty;
            string words = NumberToVietnameseWords(value);
            if (string.IsNullOrWhiteSpace(words)) return string.Empty;
            return char.ToUpper(words[0]) + words.Substring(1) + " đồng";
        }

        private static bool ContainsAnyEndMarker(string text, string blockName, string[] endAliases)
        {
            return Regex.IsMatch(text ?? string.Empty, BuildEndMarkerPattern(blockName, endAliases), RegexOptions.IgnoreCase);
        }

        private static string BuildEndMarkerPattern(string blockName, string[] endAliases)
        {
            var names = new List<string> { blockName };
            if (endAliases != null) names.AddRange(endAliases.Where(x => !string.IsNullOrWhiteSpace(x)));
            return "(?:" + string.Join("|", names.Select(x => Regex.Escape("{{/" + x + "}}"))) + ")";
        }

        private static string ReplaceIgnoreCaseGqvl(string input, string search, string replacement)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(search)) return input;
            return Regex.Replace(input, Regex.Escape(search), replacement ?? "", RegexOptions.IgnoreCase);
        }

        private void OpenGqvlFiles(IEnumerable<string> files)
        {
            foreach (var file in files)
            {
                try
                {
                    if (File.Exists(file))
                        Process.Start(new ProcessStartInfo { FileName = file, UseShellExecute = true });
                }
                catch { }
            }
        }
    }
}
