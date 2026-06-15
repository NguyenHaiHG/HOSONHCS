using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace HOSONHCS
{
    internal static class KyCuoiBlockProcessor
    {
        private const string StartMarker = "{{kycuoi_block}}";
        private const string EndMarker = "{{/kycuoi_block}}";

        public static void Apply(MainDocumentPart mainPart, string soTienText, string thoiHanVayText, string phanKyText)
        {
            if (mainPart == null || mainPart.Document == null) return;

            bool shouldShow = ShouldShowKyCuoiBlock(soTienText, thoiHanVayText, phanKyText);
            ProcessPart(mainPart, shouldShow);

            foreach (var headerPart in mainPart.HeaderParts)
                ProcessPart(headerPart, shouldShow);

            foreach (var footerPart in mainPart.FooterParts)
                ProcessPart(footerPart, shouldShow);
        }

        public static bool ShouldShowKyCuoiBlock(string soTienText, string thoiHanVayText, string phanKyText)
        {
            long soTien = ParseLong(soTienText);
            int thoiHanVay = ParseFirstInt(thoiHanVayText);
            int phanKy = ParseFirstInt(phanKyText);

            if (soTien <= 0 || thoiHanVay <= 0 || phanKy <= 0) return false;

            return thoiHanVay % phanKy != 0;
        }

        private static void ProcessPart(OpenXmlPart part, bool shouldShow)
        {
            if (part == null || part.RootElement == null) return;

            foreach (var paragraph in part.RootElement.Descendants<Paragraph>().ToList())
            {
                string text = GetParagraphText(paragraph);
                if (text.IndexOf(StartMarker, StringComparison.OrdinalIgnoreCase) < 0 &&
                    text.IndexOf(EndMarker, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                if (shouldShow)
                {
                    KeepBlockAndAddTrailingBreak(paragraph);
                }
                else
                {
                    RemoveBlock(paragraph);
                }
            }

            part.RootElement.Save();
        }

        private static void KeepBlockAndAddTrailingBreak(Paragraph paragraph)
        {
            string combined = GetParagraphText(paragraph);
            combined = Regex.Replace(combined, Regex.Escape(StartMarker), string.Empty, RegexOptions.IgnoreCase);
            combined = Regex.Replace(combined, @"\s*" + Regex.Escape(EndMarker), Environment.NewLine, RegexOptions.IgnoreCase);

            WriteParagraphWithLineBreaks(paragraph, combined);
        }

        private static void RemoveBlock(Paragraph paragraph)
        {
            string combined = GetParagraphText(paragraph);
            string blockPattern = Regex.Escape(StartMarker) + ".*?" + Regex.Escape(EndMarker);
            combined = Regex.Replace(combined, blockPattern, string.Empty, RegexOptions.IgnoreCase);

            WriteParagraphWithLineBreaks(paragraph, combined.Trim());
        }

        private static string GetParagraphText(Paragraph paragraph)
        {
            return string.Concat(paragraph.Descendants<Text>().Select(t => t.Text ?? string.Empty));
        }

        private static void WriteParagraphWithLineBreaks(Paragraph paragraph, string text)
        {
            RunProperties runProperties = paragraph.Descendants<RunProperties>().FirstOrDefault()?.CloneNode(true) as RunProperties;

            paragraph.RemoveAllChildren<Run>();

            var lines = (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var run = new Run();
            if (runProperties != null)
                run.Append(runProperties);

            for (int i = 0; i < lines.Length; i++)
            {
                if (i > 0)
                    run.Append(new Break());

                run.Append(new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve });
            }

            paragraph.Append(run);
        }

        private static long ParseLong(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return 0;
            string digits = new string(value.Where(char.IsDigit).ToArray());
            return long.TryParse(digits, out long result) ? result : 0;
        }

        private static int ParseFirstInt(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return 0;
            var match = Regex.Match(value, @"\d+");
            return match.Success && int.TryParse(match.Value, out int result) ? result : 0;
        }
    }
}
