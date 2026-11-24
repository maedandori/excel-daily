using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace ExcelDailyRibbon
{
    // Fail closed on malformed configuration and broad, prefix-less patterns.
    public sealed class AutoSaveOptions
    {
        public bool Enabled { get; private set; }
        public int Minutes { get; private set; }
        public string Pattern { get; private set; }

        public static AutoSaveOptions Parse(string json)
        {
            var root = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(json);
            if (root == null) throw new FormatException("settings.json");
            object section;
            var workbook = root.TryGetValue("workbook", out section) ? section as Dictionary<string, object> : null;
            object pattern;
            string name = workbook != null && workbook.TryGetValue("filename_pattern", out pattern)
                ? pattern as string : null;
            var options = new AutoSaveOptions { Enabled = true, Minutes = 15, Pattern = name ?? "" };
            if (root.TryGetValue("auto_save", out section))
            {
                var settings = section as Dictionary<string, object>;
                if (settings == null) throw new FormatException("auto_save");
                object value;
                if (settings.TryGetValue("enabled", out value))
                {
                    if (!(value is bool)) throw new FormatException("auto_save.enabled");
                    options.Enabled = (bool)value;
                }
                if (settings.TryGetValue("interval_minutes", out value))
                {
                    if (!(value is int) || (int)value < 1 || (int)value > 120)
                        throw new FormatException("auto_save.interval_minutes");
                    options.Minutes = (int)value;
                }
            }
            return options;
        }

        public bool Matches(string name)
        {
            if (!Enabled || String.IsNullOrWhiteSpace(name) || String.IsNullOrWhiteSpace(Pattern)) return false;
            // Only literal-leading filename patterns with * and ? are accepted.
            // Unsupported PowerShell character classes are rejected rather than broadened.
            if (Pattern[0] == '*' || Pattern[0] == '?' || Pattern.IndexOfAny(new[] { '[', ']', '/', '\\', ':', '`' }) >= 0)
                return false;
            string regex = "\\A" + Regex.Escape(Pattern).Replace("\\*", ".*").Replace("\\?", ".") + "\\z";
            return Regex.IsMatch(name, regex, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100));
        }

        public bool ShouldSave(string name, bool saved, bool readOnly, bool shared, bool existingFile)
        {
            return !saved && !readOnly && !shared && existingFile && Matches(name);
        }
    }

    public sealed class AutoSaveSchedule
    {
        private string _pattern;
        private int _minutes;
        private DateTime _due;
        public bool Due(AutoSaveOptions options, DateTime now)
        {
            if (!options.Enabled) { Reset(); return false; }
            if (_minutes != options.Minutes || !String.Equals(_pattern, options.Pattern, StringComparison.Ordinal))
            {
                _minutes = options.Minutes;
                _pattern = options.Pattern;
                _due = now.AddMinutes(_minutes);
                return false;
            }
            return now >= _due;
        }
        public void Completed(DateTime now) { _due = now.AddMinutes(_minutes); }
        public void Reset() { _minutes = 0; _pattern = null; _due = DateTime.MinValue; }
    }
}
