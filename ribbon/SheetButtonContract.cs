using System;
using System.IO;

namespace ExcelDailyRibbon
{
    // These identifiers are embedded in .xlsx Shapes and must remain stable
    // across versions and Windows installations. No VBA/project path is stored.
    public static class SheetButtonContract
    {
        public const string DataCommand = "ExcelDaily_Refresh";
        public const string TimeCommand = "ExcelDaily_RefreshTimes";
        public const string DataTag = "ExcelDaily.Shape.Data.v1";
        public const string TimeTag = "ExcelDaily.Shape.Time.v1";
        public const string DataPrefix = "ED_BTN_DATA_";
        public const string TimePrefix = "ED_BTN_TIME_";

        public static string Command(string kind)
        {
            if (kind == "data") return DataCommand;
            if (kind == "time") return TimeCommand;
            throw new ArgumentException("Unknown button kind", nameof(kind));
        }

        public static string Tag(string kind)
        {
            if (kind == "data") return DataTag;
            if (kind == "time") return TimeTag;
            throw new ArgumentException("Unknown button kind", nameof(kind));
        }

        public static bool IsKind(string name, string alternateText, string kind)
        {
            string tag = Tag(kind);
            string prefix = kind == "data" ? DataPrefix : TimePrefix;
            return String.Equals(alternateText, tag, StringComparison.Ordinal) ||
                   (!String.IsNullOrEmpty(name) &&
                    name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
        }

        public static string RequireSavedWorkbook(string path)
        {
            // A copied Shape must never update an unrelated workbook by pattern.
            if (String.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path) ||
                !File.Exists(path))
                throw new InvalidOperationException(
                    "Lưu file Excel (.xlsx) vào ổ đĩa trước khi cập nhật.");
            return Path.GetFullPath(path);
        }
    }
}
