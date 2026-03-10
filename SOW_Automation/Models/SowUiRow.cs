using System.Text.RegularExpressions;

namespace SowAutomationTool.Models
{
    public class PlaceholderEntry
    {
        public string Key { get; set; } = "";
        public string Value { get; set; } = "";
    }

    public class SowOption
    {
        public string Label { get; set; } = "";
        public string Value { get; set; } = "";
    }

    public class PlaceholderInfo
    {
        public string FullText { get; set; } = "";
        public bool IsNested { get; set; }
        public int StartIndex { get; set; }
        public int EndIndex { get; set; }
        public List<string> InnerPlaceholders { get; set; } = new();
    }

    public class VariableMapping
    {
        public string VariableName { get; set; } = "";
        public string PlaceholderKey { get; set; } = "";
    }

    public class HierarchicalTip
    {
        public string Id { get; set; } = "";       // e.g. "1", "1a", "1b"
        public string ParentId { get; set; } = "";  // e.g. "" for root, "1" for children
        public bool IsYesNo { get; set; }
        public bool HasOptions { get; set; }
        public string RawContent { get; set; } = ""; // the text after the id marker
        public List<SowOption> Options { get; set; } = new();
        public int PlaceholderIndex { get; set; } = -1; // which placeholder this maps to (-1 = none, root)
    }

    public class SowUiRow
    {
        public int RowNumber { get; set; }
        public string ClauseNumber { get; set; }
        public string SowText { get; set; }
        public string SowSummary { get; set; }
        public string EpEmResponse { get; set; }
        public string Tips { get; set; }
        public string? UserAnswer { get; set; }
        public string VariableName { get; set; } = "";
        public string ParentClauses { get; set; } = "";
        public bool IsMandatory { get; set; }

        public List<SowOption> Options { get; set; } = new();

        public List<PlaceholderEntry>? PlaceholderAnswers { get; set; }
        public string? AppendText { get; set; }

        /// <summary>
        /// Parses placeholders from SowText with support for:
        /// - Escape: *[text]* is ignored
        /// - Nesting: [outer [inner1] and [inner2]] detected with inner list
        /// - NOTE TO DRAFT exclusion
        /// </summary>
        public List<PlaceholderInfo> GetPlaceholderInfos()
        {
            if (string.IsNullOrWhiteSpace(SowText))
                return new();

            var result = new List<PlaceholderInfo>();
            var text = SowText;
            int pos = 0;

            while (pos < text.Length)
            {
                if (text[pos] != '[') { pos++; continue; }

                bool isEscaped = pos > 0 && text[pos - 1] == '*';

                int depth = 1;
                int start = pos;
                int j = pos + 1;
                bool hasInnerBracket = false;

                while (j < text.Length && depth > 0)
                {
                    if (text[j] == '[') { if (depth == 1) hasInnerBracket = true; depth++; }
                    else if (text[j] == ']') { depth--; }
                    j++;
                }

                if (depth != 0) { pos++; continue; }

                var fullText = text.Substring(start, j - start);
                bool trailingAsterisk = j < text.Length && text[j] == '*';
                isEscaped = isEscaped && trailingAsterisk;

                if (isEscaped
                    || fullText.StartsWith("[NOTE TO DRAFT", StringComparison.OrdinalIgnoreCase)
                    || fullText.StartsWith("[Optional", StringComparison.OrdinalIgnoreCase))
                {
                    pos = j;
                    continue;
                }

                var info = new PlaceholderInfo
                {
                    FullText = fullText,
                    IsNested = hasInnerBracket,
                    StartIndex = start,
                    EndIndex = j
                };

                if (hasInnerBracket)
                {
                    var innerContent = fullText.Substring(1, fullText.Length - 2);
                    foreach (Match m in Regex.Matches(innerContent, @"\[[^\[\]]+\]"))
                        info.InnerPlaceholders.Add(m.Value);
                }

                result.Add(info);
                pos = j;
            }

            return result;
        }

        public List<string> GetPlaceholders()
        {
            return GetPlaceholderInfos().Select(p => p.FullText).ToList();
        }

        public bool HasPlaceholders => GetPlaceholderInfos().Count > 0;

        public bool HasAppendPlaceholder =>
            !string.IsNullOrWhiteSpace(SowText) && SowText.Contains("*{}*");

        public bool IsTableRow =>
            !string.IsNullOrWhiteSpace(ClauseNumber)
            && ClauseNumber.Trim().StartsWith("table", StringComparison.OrdinalIgnoreCase)
            && ClauseNumber.Trim().Length > 5;

        public string TableMarkerName =>
            IsTableRow
                ? ClauseNumber.Trim().Substring(5).Trim()
                : "";

        public bool IsSectionMarker =>
            !string.IsNullOrWhiteSpace(SowText)
            && Regex.IsMatch(SowText.Trim(), @"^\*{5}.+\*{5}$");

        public string SectionMarkerName =>
            IsSectionMarker
                ? Regex.Replace(SowText.Trim(), @"^\*{5}(.*?)\*{5}$", "$1").Trim()
                : "";

        public bool HasHierarchicalTips =>
            !string.IsNullOrWhiteSpace(Tips) && Regex.IsMatch(Tips, @"\*\*\d+\s");

        /// <summary>
        /// Parses hierarchical tips like "**1 yes/no **1a Option 1 - X, Option 2 - Y **1b"
        /// into a structured list of HierarchicalTip objects.
        /// </summary>
        public List<HierarchicalTip> GetHierarchicalTips()
        {
            if (!HasHierarchicalTips) return new();

            var result = new List<HierarchicalTip>();
            var matches = Regex.Matches(Tips, @"\*\*(\d+[a-z]?)\s*");

            for (int m = 0; m < matches.Count; m++)
            {
                var id = matches[m].Groups[1].Value;
                int contentStart = matches[m].Index + matches[m].Length;
                int contentEnd = m + 1 < matches.Count ? matches[m + 1].Index : Tips.Length;
                var content = Tips.Substring(contentStart, contentEnd - contentStart).Trim();

                // Determine parent: "1a" -> parent is "1", "1" -> root
                var parentId = "";
                if (id.Length > 1 && char.IsLetter(id[^1]))
                    parentId = id.Substring(0, id.Length - 1);

                var contentLower = content.ToLower();
                bool isYesNo = contentLower.Contains("yes") || contentLower.Contains("no");

                // Parse options from content
                var optPattern = new Regex(
                    @"(?:Option|Alternative)\s*\d+\s*[:–\-]\s*(.*?)(?=(?:Option|Alternative)\s*\d+|$)",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
                var optMatches = optPattern.Matches(content);
                var options = new List<SowOption>();
                foreach (Match om in optMatches)
                {
                    var text = om.Groups[1].Value.Trim().TrimEnd('.', ' ');
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        var label = text;
                        var ci = text.IndexOf(':');
                        if (ci > 0 && ci < 40)
                        {
                            var candidate = text[..ci].Trim();
                            if (candidate.Length <= 35) label = candidate;
                        }
                        options.Add(new SowOption { Label = label, Value = text });
                    }
                }

                result.Add(new HierarchicalTip
                {
                    Id = id,
                    ParentId = parentId,
                    IsYesNo = isYesNo,
                    HasOptions = options.Count > 0,
                    RawContent = content,
                    Options = options
                });
            }

            // Assign placeholders to child tips in order
            var phInfos = GetPlaceholderInfos();
            int phIdx = 0;
            foreach (var tip in result)
            {
                if (string.IsNullOrEmpty(tip.ParentId)) continue; // root gets no placeholder
                if (phIdx < phInfos.Count)
                {
                    tip.PlaceholderIndex = phIdx;
                    phIdx++;
                }
            }

            return result;
        }

        /// <summary>
        /// Parses VariableName column into ordered variable-to-placeholder mappings.
        /// Format examples:
        ///   "PLACE,STATE" with placeholders [xxxx],[cccc] => PLACE->[xxxx], STATE->[cccc]
        ///   "TEXT,[TXT1,TXT2]" with nested [outer [yy] [cc]] => TEXT->outer, TXT1->[yy], TXT2->[cc]
        /// </summary>
        public List<VariableMapping> GetVariableMappings()
        {
            var result = new List<VariableMapping>();
            if (string.IsNullOrWhiteSpace(VariableName)) return result;

            var phInfos = GetPlaceholderInfos();
            if (phInfos.Count == 0) return result;

            var tokens = ParseVariableTokens(VariableName);
            int phIdx = 0;

            for (int t = 0; t < tokens.Count && phIdx < phInfos.Count; t++)
            {
                var token = tokens[t];
                var ph = phInfos[phIdx];

                if (token.StartsWith("[") && token.EndsWith("]") && ph.IsNested)
                {
                    var innerVars = token.Substring(1, token.Length - 2)
                        .Split(',')
                        .Select(s => s.Trim())
                        .Where(s => s.Length > 0)
                        .ToList();

                    for (int iv = 0; iv < innerVars.Count && iv < ph.InnerPlaceholders.Count; iv++)
                    {
                        result.Add(new VariableMapping
                        {
                            VariableName = innerVars[iv],
                            PlaceholderKey = ph.InnerPlaceholders[iv]
                        });
                    }
                    phIdx++;
                }
                else if (!token.StartsWith("["))
                {
                    result.Add(new VariableMapping
                    {
                        VariableName = token,
                        PlaceholderKey = ph.FullText
                    });
                    // If the next token is a bracket group and current ph is nested,
                    // don't advance phIdx so the bracket token maps inner placeholders
                    bool nextIsBracket = t + 1 < tokens.Count
                        && tokens[t + 1].StartsWith("[") && tokens[t + 1].EndsWith("]");
                    if (!(nextIsBracket && ph.IsNested))
                        phIdx++;
                }
            }

            return result;
        }

        private static List<string> ParseVariableTokens(string input)
        {
            var tokens = new List<string>();
            int pos = 0;
            while (pos < input.Length)
            {
                if (char.IsWhiteSpace(input[pos]) || input[pos] == ',') { pos++; continue; }

                if (input[pos] == '[')
                {
                    int end = input.IndexOf(']', pos);
                    if (end < 0) end = input.Length - 1;
                    tokens.Add(input.Substring(pos, end - pos + 1));
                    pos = end + 1;
                }
                else
                {
                    int end = pos;
                    while (end < input.Length && input[end] != ',' && input[end] != '[') end++;
                    tokens.Add(input.Substring(pos, end - pos).Trim());
                    pos = end;
                }
            }
            return tokens.Where(t => t.Length > 0).ToList();
        }
    }
}
