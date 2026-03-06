//bussiness_core
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SowAutomationTool.Models;
using System.Text;
using System.Text.RegularExpressions;

namespace SowAutomationTool.Services
{
    public class ProcessingService
    {
        private readonly ILogger<ProcessingService> _logger;

        public ProcessingService(ILogger<ProcessingService> logger)
        {
            _logger = logger;
        }

        #region  Parse Excel

        public List<SowUiRow> ParseExcel(byte[] excelData)
        {
            var list = new List<SowUiRow>();

            using var stream = new MemoryStream(excelData);
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheet(1);

            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var uiRow = new SowUiRow
                {
                    RowNumber = int.TryParse(row.Cell(2).GetString(), out var num) ? num : 0,
                    ClauseNumber = row.Cell(3)?.GetString()?.Trim() ?? "",
                    SowText = row.Cell(4)?.GetString()?.Trim() ?? "",
                    SowSummary = row.Cell(5)?.GetString() ?? "",
                    EpEmResponse = row.Cell(6)?.GetString() ?? "",
                    Tips = row.Cell(7)?.GetString() ?? "",
                    VariableName = row.Cell(9)?.GetString()?.Trim() ?? "",
                    ParentClauses = row.Cell(10)?.GetString()?.Trim() ?? "",
                    UserAnswer = ""
                };

                uiRow.Options = ParseOptions(uiRow.SowSummary);
                // If Tips has options with placeholders, prefer those
                var tipsText = uiRow.Tips?.Trim().Trim('"') ?? "";
                var tipsOptions = ParseOptions(tipsText);
                if (tipsOptions.Count > 0 && tipsOptions.Any(o => o.Value.Contains('[')))
                    uiRow.Options = tipsOptions;
                list.Add(uiRow);
            }

            return list;
        }

        /// <summary>
        /// Parses "Option N" / "Alternative N" patterns from the summary text.
        /// Handles formats like:
        ///   "Option 1 - text... Option 2 - text..."
        ///   "Option 1: Fixed Price: description... Option 2: T&amp;M: description..."
        ///   "Alternative 1: text... Alternative 2: text..."
        /// Returns a list with Label (short display name) and Value (full option text).
        /// </summary>
        public List<SowOption> ParseOptions(string? summary)
        {
            var options = new List<SowOption>();
            if (string.IsNullOrWhiteSpace(summary))
                return options;

            // Match "Option N" or "Alternative N" followed by separator and text
            var pattern = new Regex(
                @"(?:Option|Alternative)\s*\d+\s*[:–\-]\s*(.*?)(?=(?:Option|Alternative)\s*\d+|$)",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);

            var matches = pattern.Matches(summary);

            foreach (Match match in matches)
            {
                var fullText = match.Groups[1].Value.Trim().TrimEnd('.', ' ');
                if (string.IsNullOrWhiteSpace(fullText))
                    continue;

                // Try to extract a short label if format is "Label: rest of text"
                // e.g. "Fixed Price: EY's correction..." -> Label="Fixed Price"
                var label = fullText;
                var colonIdx = fullText.IndexOf(':');
                if (colonIdx > 0 && colonIdx < 40)
                {
                    var candidate = fullText[..colonIdx].Trim();
                    // Only use as label if it's short and doesn't look like a sentence
                    if (candidate.Length <= 35 && !candidate.Contains(' ', StringComparison.Ordinal)
                        || candidate.Split(' ').Length <= 4)
                    {
                        label = candidate;
                    }
                }

                options.Add(new SowOption
                {
                    Label = label,
                    Value = fullText
                });
            }

            return options;
        }

        #endregion

        #region  Extract Highlighted Text

        public List<string> ExtractHighlightedText(byte[] wordData)
        {
            var results = new List<string>();

            using var ms = new MemoryStream(wordData);
            using var doc = WordprocessingDocument.Open(ms, false);
            var body = doc.MainDocumentPart!.Document.Body;

            foreach (var para in body.Descendants<Paragraph>())
            {
                var sb = new StringBuilder();

                foreach (var run in para.Descendants<Run>())
                {
                    if (IsRunHighlighted(run))
                        sb.Append(run.InnerText);
                }

                var combined = sb.ToString().Trim();
                if (!string.IsNullOrWhiteSpace(combined))
                {
                    results.Add(combined);
                    _logger.LogInformation("HighlightedText: {Text}",
                        combined.Length > 150 ? combined.Substring(0, 150) + "..." : combined);
                }
                else
                {
                    var paraText = para.InnerText?.Trim() ?? "";
                    if (paraText.Contains("scope limitation", StringComparison.OrdinalIgnoreCase)
                        || paraText.Contains("Alternative 1", StringComparison.OrdinalIgnoreCase)
                        || paraText.Contains("Transfer Assistance", StringComparison.OrdinalIgnoreCase))
                    {
                        var totalRuns = para.Descendants<Run>().Count();
                        var highlightedRuns = para.Descendants<Run>().Where(IsRunHighlighted).Count();
                        _logger.LogWarning("MISSED paragraph: totalRuns={Total}, highlightedRuns={HL}, text={Text}",
                            totalRuns, highlightedRuns,
                            paraText.Length > 200 ? paraText.Substring(0, 200) + "..." : paraText);

                        // Log run-level details for first few runs
                        foreach (var run in para.Descendants<Run>().Take(5))
                        {
                            var rPr = run.RunProperties;
                            var hasHL = rPr?.Elements<Highlight>().Any() == true;
                            var shdInfo = "";
                            if (rPr != null)
                            {
                                foreach (var shd in rPr.Elements<Shading>())
                                    shdInfo += $"fill={shd.Fill?.Value} color={shd.Color?.Value} ";
                            }
                            _logger.LogWarning("  Run: text='{RunText}' highlight={HL} shading='{Shd}'",
                                run.InnerText.Length > 50 ? run.InnerText.Substring(0, 50) : run.InnerText,
                                hasHL, shdInfo);
                        }
                    }
                }
            }

            return results;
        }

        #endregion

        #region  Match Highlighted Text

        public List<SowUiRow> GetMatchedRows(
            List<SowUiRow> excelRows,
            List<string> highlightedList)
        {
            var matched = new List<SowUiRow>();

            // Section markers are always included (they span multiple pages, no highlight needed)
            foreach (var row in excelRows)
            {
                if (row.RowNumber == 0) continue;
                if (row.IsSectionMarker && !matched.Contains(row))
                    matched.Add(row);
            }

            foreach (var highlight in highlightedList)
            {
                var highlightNorm = Normalize(highlight);

                foreach (var row in excelRows)
                {
                    if (row.RowNumber == 0) continue;
                    if (row.IsSectionMarker) continue;

                    var sowNorm = Normalize(row.SowText);
                    if (string.IsNullOrWhiteSpace(sowNorm))
                        continue;

                    if (highlightNorm.Contains(sowNorm) ||
                        sowNorm.Contains(highlightNorm))
                    {
                        if (!matched.Contains(row))
                            matched.Add(row);
                    }
                }
            }

            // Log unmatched rows for debugging
            foreach (var row in excelRows)
            {
                if (row.RowNumber == 0 || row.IsSectionMarker) continue;
                if (!matched.Contains(row) && !string.IsNullOrWhiteSpace(row.SowText))
                {
                    _logger.LogWarning("Unmatched row {Row} clause={Clause}: SowText={Sow}",
                        row.RowNumber, row.ClauseNumber,
                        row.SowText.Length > 100 ? row.SowText.Substring(0, 100) + "..." : row.SowText);
                }
            }

            return matched;
        }

        #endregion

        #region 4️⃣ Generate Final Document

        public byte[] GenerateDocument(
            byte[] wordData,
            List<SowUiRow> matchedRows)
        {
            using var ms = new MemoryStream();
            ms.Write(wordData);
            ms.Position = 0;

            using (var doc = WordprocessingDocument.Open(ms, true))
            {
                var body = doc.MainDocumentPart!.Document.Body;

                RemoveSchedulesByHeading(body, matchedRows);
                RemoveSectionMarkers(body, matchedRows);
                StripTableMarkerText(body);

                var paragraphs = body.Descendants<Paragraph>().ToList();
                var appliedRowIndices = new HashSet<int>();

                // Log table row info for debugging
                var tableRowEntries = matchedRows.Where(r => r.IsTableRow).ToList();
                foreach (var tr in tableRowEntries)
                {
                    _logger.LogInformation("TableRow entry: Row={Row}, Clause={Clause}, MarkerName={Marker}, SowText={Text}, Parent={Parent}",
                        tr.RowNumber, tr.ClauseNumber, tr.TableMarkerName,
                        (tr.SowText ?? "").Length > 80 ? tr.SowText!.Substring(0, 80) + "..." : tr.SowText,
                        tr.ParentClauses);
                }

                // Log all paragraphs inside tables
                var tableParagraphs = body.Descendants<Table>()
                    .SelectMany(t => t.Descendants<Paragraph>())
                    .ToList();
                foreach (var tp in tableParagraphs)
                {
                    var tpText = tp.InnerText;
                    if (string.IsNullOrWhiteSpace(tpText)) continue;
                    var tpRegions = GetHighlightedRegions(tp);
                    var isHighlighted = tpRegions.Count > 0;
                    _logger.LogInformation("TableParagraph: InTable=true, Highlighted={Hl}, RegionCount={Rc}, Text={Text}",
                        isHighlighted, tpRegions.Count,
                        tpText.Length > 100 ? tpText.Substring(0, 100) + "..." : tpText);
                }

                foreach (var para in paragraphs)
                {
                    var regions = GetHighlightedRegions(para);
                    if (regions.Count == 0) continue;

                    var inTable = para.Ancestors<Table>().Any();

                    for (var r = regions.Count - 1; r >= 0; r--)
                    {
                        var (runs, regionText) = regions[r];
                        var regionNorm = Normalize(regionText);
                        if (string.IsNullOrWhiteSpace(regionNorm)) continue;

                        if (inTable)
                        {
                            _logger.LogInformation("HighlightedRegion in table: RegionText={Text}",
                                regionText.Length > 100 ? regionText.Substring(0, 100) + "..." : regionText);
                        }

                        SowUiRow? matchedRow = null;
                        int matchedIndex = -1;

                        for (var i = 0; i < matchedRows.Count; i++)
                        {
                            if (appliedRowIndices.Contains(i)) continue;

                            var row = matchedRows[i];
                            if (row.IsSectionMarker) continue;
                            var sowNorm = Normalize(row.SowText ?? "");

                            if (string.IsNullOrWhiteSpace(sowNorm)) continue;
                            if (!regionNorm.Contains(sowNorm) &&
                                !sowNorm.Contains(regionNorm))
                                continue;

                            matchedRow = row;
                            matchedIndex = i;
                            break;
                        }

                        if (matchedRow == null)
                        {
                            if (inTable)
                                _logger.LogWarning("Table region UNMATCHED: {Text}", regionText.Length > 100 ? regionText.Substring(0, 100) + "..." : regionText);
                            continue;
                        }

                        _logger.LogInformation("Matched: Row={Row}, Clause={Clause}, IsTable={IsTable}, InTable={InTable}",
                            matchedRow.RowNumber, matchedRow.ClauseNumber, matchedRow.IsTableRow, inTable);

                        appliedRowIndices.Add(matchedIndex);
                        ApplyAnswer(matchedRow, runs, regionText, para);

                        if (matchedRow.UserAnswer?.Trim().Equals("No", StringComparison.OrdinalIgnoreCase) == true)
                            break;
                    }
                }

                // Log unmatched table rows
                foreach (var tr in tableRowEntries)
                {
                    var idx = matchedRows.IndexOf(tr);
                    if (!appliedRowIndices.Contains(idx))
                        _logger.LogWarning("TableRow NOT matched in doc: Row={Row}, SowText={Text}",
                            tr.RowNumber, (tr.SowText ?? "").Length > 80 ? tr.SowText!.Substring(0, 80) + "..." : tr.SowText);
                }

                // Replace **variableName** references across the entire document
                ReplaceVariablesInDocument(body, matchedRows);

                // Remove any remaining *{}* append placeholders
                RemoveRemainingAppendPlaceholders(body);

                // Strip escape markers: *[text]* → [text]
                RemoveEscapeMarkersFromDocument(body);

                // Remove all remaining highlight formatting from the document
                var remainingHighlightedRuns = body.Descendants<Run>().Where(IsRunHighlighted).ToList();
                RemoveHighlightFormattingFromRuns(remainingHighlightedRuns);

                RemoveNoteToDraftFromDocument(body);
                doc.MainDocumentPart.Document.Save();
            }

            ms.Position = 0;
            return ms.ToArray();
        }

        #endregion

        #region  Correct Schedule Deletion 

        private void RemoveSchedulesByHeading(
        Body body,
        List<SowUiRow> matchedRows)
            {
                // Only remove schedules when the section marker row itself is answered "No"
                // Match only root schedule clauses like "Sch A", "Sch_A", not sub-clauses like "Sch A 1c"
                var schedulesToRemove = matchedRows
                    .Where(r =>
                        r.UserAnswer?.Equals("No", StringComparison.OrdinalIgnoreCase) == true &&
                        r.IsSectionMarker &&
                        r.ClauseNumber.StartsWith("Sch", StringComparison.OrdinalIgnoreCase))
                    .Select(r =>
                    {
                        var match = Regex.Match(r.ClauseNumber, @"^Sch[\s_]*([A-Za-z])$",
                                                RegexOptions.IgnoreCase);
                        return match.Success ? match.Groups[1].Value.ToLower() : null;
                    })
                    .Where(x => !string.IsNullOrEmpty(x))
                    .ToList();

                if (!schedulesToRemove.Any())
                    return;

                var elements = body.Elements().ToList();

                for (int i = 0; i < elements.Count; i++)
                {
                    if (elements[i] is Paragraph para)
                    {
                        var styleId = para.ParagraphProperties?
                                          .ParagraphStyleId?
                                          .Val?
                                          .Value;

                        if (styleId == "Heading1")
                        {
                            var headingText = Normalize(para.InnerText);

                            foreach (var schLetter in schedulesToRemove)
                            {
                                if (Regex.IsMatch(headingText,
                                    $@"\bschedule\s+{schLetter}\b",
                                    RegexOptions.IgnoreCase))
                                {
                                    int startIndex = i;
                                    int endIndex = i + 1;

                                    while (endIndex < elements.Count)
                                    {
                                        if (elements[endIndex] is Paragraph nextPara)
                                        {
                                            var nextStyle = nextPara.ParagraphProperties?
                                                                    .ParagraphStyleId?
                                                                    .Val?
                                                                    .Value;

                                            if (nextStyle == "Heading1")
                                            {
                                                var nextText = Normalize(nextPara.InnerText);

                                                // 🔥 STOP if next schedule heading found
                                                if (Regex.IsMatch(nextText,
                                                    @"\bschedule\s+[a-z]",
                                                    RegexOptions.IgnoreCase))
                                                {
                                                    break;
                                                }
                                            }
                                        }

                                        endIndex++;
                                    }

                                    for (int j = startIndex; j < endIndex; j++)
                                        elements[j].Remove();

                                    elements = body.Elements().ToList();
                                    i = -1;
                                    break;
                                }
                            }
                        }
                    }
                }
            }

        #endregion

        #region Section Marker Removal

        private void RemoveSectionMarkers(Body body, List<SowUiRow> matchedRows)
        {
            var sectionRows = matchedRows.Where(r => r.IsSectionMarker).ToList();
            if (sectionRows.Count == 0) return;

            foreach (var sectionRow in sectionRows)
            {
                var sectionName = sectionRow.SectionMarkerName;
                var answer = sectionRow.UserAnswer?.Trim() ?? "";
                var removeSection = answer.Equals("No", StringComparison.OrdinalIgnoreCase);

                _logger.LogInformation("SectionMarker: '{Name}' Answer={Answer} Remove={Remove}",
                    sectionName, answer, removeSection);

                var markerText = "*****" + sectionName + "*****";

                // Find the two element indices containing the start and end markers
                var elements = body.Elements().ToList();
                int startIdx = -1;
                int endIdx = -1;

                for (int i = 0; i < elements.Count; i++)
                {
                    var elText = elements[i].InnerText;
                    if (string.IsNullOrEmpty(elText)) continue;

                    var normalized = elText.Replace("\u00A0", " ");
                    if (normalized.Contains(markerText, StringComparison.OrdinalIgnoreCase))
                    {
                        if (startIdx == -1)
                            startIdx = i;
                        else
                        {
                            endIdx = i;
                            break;
                        }
                    }
                }

                if (startIdx == -1) continue;
                if (endIdx == -1) endIdx = startIdx;

                if (removeSection)
                {
                    // Remove everything between the markers.
                    // For shared paragraphs (start or end contains other markers),
                    // strip only this marker text instead of removing the whole element.
                    bool startShared = ElementHasOtherMarkers(elements[startIdx], sectionName);
                    bool endShared = endIdx != startIdx && ElementHasOtherMarkers(elements[endIdx], sectionName);

                    // Remove elements strictly between start and end (exclusive)
                    for (int i = endIdx - 1; i > startIdx; i--)
                        elements[i].Remove();

                    // Handle end element
                    if (endIdx != startIdx)
                    {
                        if (endShared)
                            StripMarkerTextFromElement(elements[endIdx], markerText);
                        else
                            elements[endIdx].Remove();
                    }

                    // Handle start element
                    if (startShared)
                        StripMarkerTextFromElement(elements[startIdx], markerText);
                    else
                        elements[startIdx].Remove();

                    // Clean up empty paragraphs with page/section breaks left behind
                    RemoveEmptyPageBreakParagraphs(body);
                }
                else
                {
                    // Keep content, strip marker text from the paragraphs
                    if (endIdx != startIdx)
                        StripMarkerTextFromElement(elements[endIdx], markerText);
                    StripMarkerTextFromElement(elements[startIdx], markerText);

                    // If stripping left an empty paragraph, remove it
                    if (endIdx != startIdx && string.IsNullOrWhiteSpace(elements[endIdx].InnerText))
                        elements[endIdx].Remove();
                    if (string.IsNullOrWhiteSpace(elements[startIdx].InnerText))
                        elements[startIdx].Remove();
                }
            }
        }

        private static bool ElementHasOtherMarkers(OpenXmlElement element, string currentSectionName)
        {
            var text = element.InnerText.Replace("\u00A0", " ");
            var matches = Regex.Matches(text, @"\*{5}(.+?)\*{5}");
            return matches.Any(m => !m.Groups[1].Value.Trim()
                .Equals(currentSectionName, StringComparison.OrdinalIgnoreCase));
        }

        private static void StripMarkerTextFromElement(OpenXmlElement element, string markerText)
        {
            foreach (var run in element.Descendants<Run>().ToList())
            {
                var textEl = run.GetFirstChild<Text>();
                if (textEl == null) continue;

                var original = textEl.Text;
                if (string.IsNullOrEmpty(original)) continue;

                var replaced = original.Replace(markerText, "", StringComparison.OrdinalIgnoreCase);
                // Also handle NBSP variant
                var nbspMarker = markerText.Replace(" ", "\u00A0");
                replaced = replaced.Replace(nbspMarker, "", StringComparison.OrdinalIgnoreCase);

                if (replaced != original)
                {
                    textEl.Text = replaced.Trim();
                    textEl.Space = SpaceProcessingModeValues.Preserve;
                }
            }
        }

        private static void StripTableMarkerText(Body body)
        {
            var markerRegex = new Regex(@"&{5}table.+?&{5}", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            foreach (var run in body.Descendants<Run>().ToList())
            {
                var textEl = run.GetFirstChild<Text>();
                if (textEl == null) continue;
                var original = textEl.Text;
                if (string.IsNullOrEmpty(original) || !original.Contains("&&&&&")) continue;

                var replaced = markerRegex.Replace(original, "").Trim();
                if (replaced != original)
                {
                    textEl.Text = replaced;
                    textEl.Space = SpaceProcessingModeValues.Preserve;
                }
            }
            // Remove paragraphs left empty after stripping markers
            foreach (var para in body.Descendants<Paragraph>().ToList())
            {
                if (!string.IsNullOrWhiteSpace(para.InnerText)) continue;
                if (para.Ancestors<TableCell>().Any()) continue;
                if (para.Descendants<Drawing>().Any() || para.Descendants<Picture>().Any()) continue;
                if (para.InnerText == "" && para.Descendants<Run>().All(r => string.IsNullOrEmpty(r.InnerText)))
                    para.Remove();
            }
        }

        private static void RemoveEmptyPageBreakParagraphs(Body body)
        {
            foreach (var para in body.Descendants<Paragraph>().ToList())
            {
                if (!string.IsNullOrWhiteSpace(para.InnerText)) continue;

                var hasPageBreak = false;
                // Check run-level page breaks
                foreach (var br in para.Descendants<Break>())
                {
                    if (br.Type?.Value == BreakValues.Page || br.Type?.Value == BreakValues.Column)
                    {
                        hasPageBreak = true;
                        break;
                    }
                }
                // Check paragraph-level page break before
                var pPr = para.ParagraphProperties;
                if (pPr != null)
                {
                    var pageBreakBefore = pPr.GetFirstChild<PageBreakBefore>();
                    if (pageBreakBefore != null) hasPageBreak = true;

                    // Check section properties with page break
                    var sectPr = pPr.GetFirstChild<SectionProperties>();
                    if (sectPr != null) hasPageBreak = true;
                }

                if (hasPageBreak)
                    para.Remove();
            }
        }

        #endregion

        #region Apply Answer

        private void ApplyAnswer(SowUiRow matchedRow, List<Run> runs, string regionText, Paragraph para)
        {
            var answer = matchedRow.UserAnswer?.Trim() ?? "";

            _logger.LogInformation(
                "Row {Row} matched. Answer={Answer}, PhCount={PhCount}, RunText={RunText}, RunCount={RunCount}",
                matchedRow.RowNumber,
                answer,
                matchedRow.PlaceholderAnswers?.Count ?? 0,
                regionText.Length > 200 ? regionText.Substring(0, 200) + "..." : regionText,
                runs.Count);

            if (answer.Equals("No", StringComparison.OrdinalIgnoreCase))
            {
                if (matchedRow.IsTableRow)
                {
                    var tableRow = para.Ancestors<TableRow>().FirstOrDefault();
                    if (tableRow != null)
                        tableRow.Remove();
                    else
                        para.Remove();
                }
                else
                {
                    para.Remove();
                }
            }
            else if (answer.Equals("N/A", StringComparison.OrdinalIgnoreCase))
            {
                ReplaceSpecificRuns(runs, "N/A");
            }
            else if (matchedRow.PlaceholderAnswers != null
                && matchedRow.PlaceholderAnswers.Any(pa => !string.IsNullOrWhiteSpace(pa.Value)))
            {
                var paList = matchedRow.PlaceholderAnswers
                    .Where(pa => !string.IsNullOrWhiteSpace(pa.Key))
                    .ToList();
                // Check for duplicate keys (positional replacement needed)
                var hasDuplicateKeys = paList.GroupBy(pa => pa.Key).Any(g => g.Count() > 1);
                if (hasDuplicateKeys)
                {
                    ReplacePositionalPlaceholders(runs, paList);
                    // If the answer is an option with placeholders, replace the outer
                    // [Alternative 1:... OR Alternative 2:...] bracket with the filled option text
                    if (!string.IsNullOrWhiteSpace(answer) && matchedRow.Options?.Count > 0)
                    {
                        ReplaceOuterAlternativeBracket(runs, answer, paList);
                    }
                }
                else
                {
                    var placeholderDict = paList
                        .GroupBy(pa => pa.Key)
                        .ToDictionary(g => g.Key, g => g.Last().Value ?? "");
                    ReplacePlaceholdersInRuns(runs, placeholderDict, matchedRow.GetPlaceholderInfos());
                }
                if (matchedRow.HasAppendPlaceholder)
                    ReplaceAppendInRuns(runs, matchedRow.AppendText);
                RemoveHighlightFormattingFromRuns(runs);
            }
            else if (answer.Equals("Yes", StringComparison.OrdinalIgnoreCase))
            {
                if (matchedRow.HasAppendPlaceholder)
                    ReplaceAppendInRuns(runs, matchedRow.AppendText);
                RemoveHighlightFormattingFromRuns(runs);
            }
            else if (!string.IsNullOrWhiteSpace(answer))
            {
                ReplaceSpecificRuns(runs, answer);
            }
            else
            {
                RemoveHighlightFormattingFromRuns(runs);
            }
        }

        #endregion

        #region Utilities

        private static bool IsRunHighlighted(Run run)
        {
            var rPr = run.RunProperties;
            if (rPr == null) return false;

            if (rPr.Elements<Highlight>().Any())
                return true;

            foreach (var shd in rPr.Elements<Shading>())
            {
                var fill = shd.Fill?.Value;
                if (string.IsNullOrEmpty(fill)) continue;

                var f = fill.Trim().ToUpperInvariant();
                if (f != "AUTO" && f != "FFFFFF")
                    return true;
            }

            return false;
        }

        private static List<(List<Run> Runs, string Text)>
            GetHighlightedRegions(Paragraph para)
        {
            var runs = para.Descendants<Run>().ToList();
            var regions = new List<(List<Run>, string)>();
            var current = new List<Run>();
            var sb = new StringBuilder();

            foreach (var run in runs)
            {
                if (IsRunHighlighted(run))
                {
                    current.Add(run);
                    sb.Append(run.InnerText);
                }
                else
                {
                    if (current.Count > 0)
                    {
                        regions.Add((new List<Run>(current), sb.ToString()));
                        current.Clear();
                        sb.Clear();
                    }
                }
            }

            if (current.Count > 0)
                regions.Add((current, sb.ToString()));

            return regions;
        }

        private static void RemoveHighlightFormattingFromRuns(List<Run> runs)
        {
            foreach (var run in runs)
            {
                run.RunProperties?.RemoveAllChildren<Highlight>();
                run.RunProperties?.RemoveAllChildren<Shading>();
            }
        }

        private void ReplacePositionalPlaceholders(List<Run> runs, List<PlaceholderEntry> entries)
        {
            var fullText = new StringBuilder();
            foreach (var run in runs)
                fullText.Append(run.InnerText);

            var text = fullText.ToString();

            _logger.LogInformation("ReplacePositional: RunText={Text}, Entries={Entries}",
                text.Length > 200 ? text.Substring(0, 200) + "..." : text,
                string.Join("; ", entries.Select(e => $"'{e.Key}'->'{e.Value}'")));

            // Replace each occurrence positionally (1st entry replaces 1st occurrence, etc.)
            foreach (var entry in entries)
            {
                if (string.IsNullOrWhiteSpace(entry.Value)) continue;

                var idx = text.IndexOf(entry.Key, StringComparison.OrdinalIgnoreCase);
                if (idx >= 0)
                {
                    text = text.Substring(0, idx) + entry.Value + text.Substring(idx + entry.Key.Length);
                }
                else
                {
                    _logger.LogWarning("ReplacePositional: Key '{Key}' NOT FOUND in text", entry.Key);
                }
            }

            if (runs.Count > 0)
            {
                var firstRun = runs[0];
                var textEl = firstRun.GetFirstChild<Text>();
                if (textEl != null)
                {
                    textEl.Text = text;
                    textEl.Space = SpaceProcessingModeValues.Preserve;
                }
                else
                {
                    firstRun.RemoveAllChildren<Text>();
                    firstRun.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
                }
                for (int i = 1; i < runs.Count; i++)
                {
                    var t = runs[i].GetFirstChild<Text>();
                    if (t != null) t.Text = "";
                }
            }
        }

        private void ReplaceOuterAlternativeBracket(List<Run> runs, string answer, List<PlaceholderEntry> entries)
        {
            // Build the filled answer: replace [placeholder] in the option text with user values
            var filledAnswer = answer;
            foreach (var entry in entries)
            {
                if (string.IsNullOrWhiteSpace(entry.Value)) continue;
                // The answer uses [Insert scope limitations], map back from round to square
                var squareKey = entry.Key;
                if (squareKey.StartsWith("(") && squareKey.EndsWith(")"))
                    squareKey = "[" + squareKey.Substring(1, squareKey.Length - 2) + "]";

                var idx = filledAnswer.IndexOf(squareKey, StringComparison.OrdinalIgnoreCase);
                if (idx >= 0)
                    filledAnswer = filledAnswer.Substring(0, idx) + entry.Value + filledAnswer.Substring(idx + squareKey.Length);
            }

            // Get full run text after positional replacement
            var sb = new StringBuilder();
            foreach (var run in runs)
                sb.Append(run.InnerText);
            var text = sb.ToString();

            // Find the outer [Alternative 1:... OR Alternative 2:...] bracket
            var altRegex = new Regex(@"\[(Alternative|Option)\s*\d+\s*[:–\-].*?\](?=[^()\[\]]*$|\s*[;\.,]|$)",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);

            // Simpler approach: find the outermost bracket that starts with [Alternative or [Option
            // and contains "OR Alternative" or "OR Option"
            int bracketStart = -1;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '[')
                {
                    var after = text.Substring(i + 1);
                    if (Regex.IsMatch(after, @"^\s*(Alternative|Option)\s*\d+", RegexOptions.IgnoreCase))
                    {
                        bracketStart = i;
                        break;
                    }
                }
            }

            if (bracketStart >= 0)
            {
                // Find the matching closing bracket
                int depth = 0;
                int bracketEnd = -1;
                for (int i = bracketStart; i < text.Length; i++)
                {
                    if (text[i] == '[') depth++;
                    else if (text[i] == ']') { depth--; if (depth == 0) { bracketEnd = i; break; } }
                }

                if (bracketEnd > bracketStart)
                {
                    _logger.LogInformation("ReplaceOuterBracket: Replacing [{Start}..{End}] with '{Filled}'",
                        bracketStart, bracketEnd, filledAnswer.Length > 100 ? filledAnswer.Substring(0, 100) + "..." : filledAnswer);

                    text = text.Substring(0, bracketStart) + filledAnswer + text.Substring(bracketEnd + 1);

                    // Write back to runs
                    if (runs.Count > 0)
                    {
                        var firstRun = runs[0];
                        var textEl = firstRun.GetFirstChild<Text>();
                        if (textEl != null)
                        {
                            textEl.Text = text;
                            textEl.Space = SpaceProcessingModeValues.Preserve;
                        }
                        else
                        {
                            firstRun.RemoveAllChildren<Text>();
                            firstRun.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
                        }
                        for (int i = 1; i < runs.Count; i++)
                        {
                            var t = runs[i].GetFirstChild<Text>();
                            if (t != null) t.Text = "";
                        }
                    }
                }
            }
        }

        private void ReplacePlaceholdersInRuns(List<Run> runs, Dictionary<string, string> placeholderAnswers, List<PlaceholderInfo>? phInfos = null)
        {
            var fullText = new StringBuilder();
            var runOffsets = new List<(Run Run, int Start, int End)>();

            foreach (var run in runs)
            {
                int start = fullText.Length;
                fullText.Append(run.InnerText);
                runOffsets.Add((run, start, fullText.Length));
            }

            var text = fullText.ToString();

            _logger.LogInformation("ReplacePlaceholders: RunText={Text}, Answers={Answers}",
                text,
                string.Join("; ", placeholderAnswers.Select(kv => $"'{kv.Key}'->'{kv.Value}'")));

            foreach (var kvp in placeholderAnswers)
            {
                var placeholder = kvp.Key;
                var replacement = kvp.Value ?? "";

                if (string.IsNullOrWhiteSpace(replacement))
                    continue;

                // Try exact replace first (case-insensitive)
                var before = text;
                text = ReplaceIgnoreCase(text, placeholder, replacement);

                // If not found, try whitespace-tolerant replace
                if (text == before)
                {
                    text = ReplaceWhitespaceTolerant(text, placeholder, replacement);
                }

                if (text == before)
                    _logger.LogWarning("ReplacePlaceholders: Key '{Key}' NOT FOUND in text", placeholder);
            }

            // For nested placeholders in custom mode: strip outer brackets
            if (phInfos != null)
            {
                foreach (var ph in phInfos)
                {
                    if (!ph.IsNested) continue;
                    // Custom mode: full key not in answers (or empty), but inner keys were replaced
                    if (placeholderAnswers.ContainsKey(ph.FullText)
                        && !string.IsNullOrWhiteSpace(placeholderAnswers[ph.FullText]))
                        continue;

                    // The outer bracket text after inner replacement: find remaining [...]
                    // that matches the outer placeholder structure (starts with [ and ends with ])
                    // We need to find the remnant of the outer placeholder in the text.
                    // After inner replacements, the outer text looks like "[power of value1 and value2]"
                    // Build the expected remnant by replacing inner placeholders in the full text
                    var remnant = ph.FullText;
                    foreach (var inner in ph.InnerPlaceholders)
                    {
                        if (placeholderAnswers.TryGetValue(inner, out var val) && !string.IsNullOrWhiteSpace(val))
                            remnant = ReplaceIgnoreCase(remnant, inner, val);
                    }
                    if (remnant.StartsWith("[") && remnant.EndsWith("]"))
                    {
                        var stripped = remnant.Substring(1, remnant.Length - 2);
                        text = ReplaceWhitespaceTolerant(text, remnant, stripped);
                    }
                }
            }

            if (runs.Count > 0)
            {
                var firstRun = runs[0];
                var textElement = firstRun.GetFirstChild<Text>();
                if (textElement != null)
                {
                    textElement.Text = text;
                    textElement.Space = SpaceProcessingModeValues.Preserve;
                }
                else
                {
                    firstRun.RemoveAllChildren<Text>();
                    firstRun.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
                }

                for (int i = 1; i < runs.Count; i++)
                {
                    var t = runs[i].GetFirstChild<Text>();
                    if (t != null) t.Text = "";
                }
            }
        }

        private void ReplaceSpecificRuns(List<Run> runs, string userText)
        {
            if (runs.Count == 0) return;

            var first = runs[0];
            var newRun = new Run(new Text(userText)
            {
                Space = SpaceProcessingModeValues.Preserve
            });

            if (first.RunProperties != null)
            {
                var props = (RunProperties)first.RunProperties.CloneNode(true);
                props.RemoveAllChildren<Highlight>();
                props.RemoveAllChildren<Shading>();
                newRun.PrependChild(props);
            }

            first.InsertBeforeSelf(newRun);

            for (int i = runs.Count - 1; i >= 0; i--)
                runs[i].Remove();
        }

        private void RemoveNoteToDraftFromDocument(Body body)
        {
            foreach (var para in body.Descendants<Paragraph>().ToList())
            {
                var runs = para.Descendants<Run>().ToList();
                if (runs.Count == 0) continue;

                var fullText = string.Join("", runs.Select(r => r.InnerText));
                if (string.IsNullOrEmpty(fullText)) continue;

                var cleaned = Regex.Replace(fullText,
                    @"\s?\[NOTE TO DRAFT[^\]]*\]", "", RegexOptions.IgnoreCase);
                cleaned = Regex.Replace(cleaned,
                    @"\s?\[Optional[^\]]*\]", "", RegexOptions.IgnoreCase);

                if (cleaned == fullText) continue;

                cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim();

                if (string.IsNullOrWhiteSpace(cleaned))
                {
                    foreach (var run in runs)
                        run.Remove();
                }
                else
                {
                    var firstRun = runs[0];
                    var textEl = firstRun.GetFirstChild<Text>();
                    if (textEl != null)
                    {
                        textEl.Text = cleaned;
                        textEl.Space = SpaceProcessingModeValues.Preserve;
                    }
                    else
                    {
                        firstRun.RemoveAllChildren<Text>();
                        firstRun.AppendChild(new Text(cleaned)
                            { Space = SpaceProcessingModeValues.Preserve });
                    }

                    for (int i = 1; i < runs.Count; i++)
                    {
                        var t = runs[i].GetFirstChild<Text>();
                        if (t != null) t.Text = "";
                    }
                }
            }
        }

        private void ReplaceAppendInRuns(List<Run> runs, string? appendText)
        {
            var value = appendText?.Trim() ?? "";
            foreach (var run in runs)
            {
                var textEl = run.GetFirstChild<Text>();
                if (textEl == null) continue;
                var original = textEl.Text;
                if (string.IsNullOrEmpty(original) || !original.Contains("*{}*")) continue;

                string cleaned;
                if (!string.IsNullOrWhiteSpace(value))
                    cleaned = original.Replace("*{}*", value);
                else
                    cleaned = Regex.Replace(original, @"\s?\*\{\}\*", "");

                textEl.Text = cleaned;
                textEl.Space = SpaceProcessingModeValues.Preserve;
            }
        }

        private void RemoveRemainingAppendPlaceholders(Body body)
        {
            foreach (var run in body.Descendants<Run>().ToList())
            {
                var textEl = run.GetFirstChild<Text>();
                if (textEl == null) continue;
                var original = textEl.Text;
                if (string.IsNullOrEmpty(original) || !original.Contains("*{}*")) continue;
                textEl.Text = Regex.Replace(original, @"\s?\*\{\}\*", "");
                textEl.Space = SpaceProcessingModeValues.Preserve;
            }
        }

        private void RemoveEscapeMarkersFromDocument(Body body)
        {
            foreach (var run in body.Descendants<Run>().ToList())
            {
                var textEl = run.GetFirstChild<Text>();
                if (textEl == null) continue;

                var original = textEl.Text;
                if (string.IsNullOrEmpty(original)) continue;

                var cleaned = Regex.Replace(original, @"\*\[([^\]]+)\]\*", "[$1]");
                if (cleaned != original)
                {
                    textEl.Text = cleaned;
                    textEl.Space = SpaceProcessingModeValues.Preserve;
                }
            }
        }

        private void ReplaceVariablesInDocument(Body body, List<SowUiRow> matchedRows)
        {
            var variableMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var row in matchedRows)
            {
                if (string.IsNullOrWhiteSpace(row.VariableName)) continue;

                var mappings = row.GetVariableMappings();
                if (mappings.Count > 0 && row.PlaceholderAnswers != null)
                {
                    var paDict = row.PlaceholderAnswers
                        .Where(pa => !string.IsNullOrWhiteSpace(pa.Key))
                        .GroupBy(pa => pa.Key)
                        .ToDictionary(g => g.Key, g => g.Last().Value ?? "");

                    foreach (var mapping in mappings)
                    {
                        if (paDict.TryGetValue(mapping.PlaceholderKey, out var val)
                            && !string.IsNullOrWhiteSpace(val))
                        {
                            variableMap[$"**{mapping.VariableName}**"] = val;
                        }
                    }
                }
                else
                {
                    var answer = row.UserAnswer?.Trim() ?? "";
                    if (!string.IsNullOrWhiteSpace(answer))
                    {
                        variableMap[$"**{row.VariableName}**"] = answer;
                    }
                }
            }

            if (variableMap.Count == 0) return;

            // Replace variables per-paragraph to handle cross-run text
            foreach (var para in body.Descendants<Paragraph>().ToList())
            {
                var runs = para.Descendants<Run>().ToList();
                if (runs.Count == 0) continue;

                var fullText = string.Join("", runs.Select(r => r.InnerText));
                if (string.IsNullOrEmpty(fullText)) continue;

                var replaced = fullText;
                foreach (var kvp in variableMap)
                    replaced = replaced.Replace(kvp.Key, kvp.Value, StringComparison.OrdinalIgnoreCase);

                if (replaced == fullText) continue;

                var firstRun = runs[0];
                var textEl = firstRun.GetFirstChild<Text>();
                if (textEl != null)
                {
                    textEl.Text = replaced;
                    textEl.Space = SpaceProcessingModeValues.Preserve;
                }
                else
                {
                    firstRun.RemoveAllChildren<Text>();
                    firstRun.AppendChild(new Text(replaced) { Space = SpaceProcessingModeValues.Preserve });
                }

                for (int ri = 1; ri < runs.Count; ri++)
                {
                    var t = runs[ri].GetFirstChild<Text>();
                    if (t != null) t.Text = "";
                }
            }
        }

        private string Normalize(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            text = text.Replace('\u00A0', ' ');
            return string.Join(" ",
                text.Split(new[] { ' ', '\r', '\n', '\t' },
                StringSplitOptions.RemoveEmptyEntries))
                .Trim()
                .ToLower();
        }

        private static string NormalizeWhitespace(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            text = text.Replace('\u00A0', ' ');
            return Regex.Replace(text, @"[\s]+", " ").Trim();
        }

        private static string ReplaceIgnoreCase(string source, string oldValue, string newValue)
        {
            return source.Replace(oldValue, newValue, StringComparison.OrdinalIgnoreCase);
        }

        private static string ReplaceWhitespaceTolerant(string source, string placeholder, string replacement)
        {
            var pattern = Regex.Escape(NormalizeWhitespace(placeholder));
            // Allow any whitespace sequence where the placeholder has a single space
            pattern = pattern.Replace(@"\ ", @"[\s\u00A0]+");
            var match = Regex.Match(source, pattern, RegexOptions.IgnoreCase);
            if (!match.Success) return source;
            return source.Substring(0, match.Index) + replacement + source.Substring(match.Index + match.Length);
        }

        #endregion
    }
}