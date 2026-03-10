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
                    // Also add the full paragraph text so matching works when the
                    // Excel SowText spans both highlighted and non-highlighted runs
                    var fullParaText = para.InnerText?.Trim() ?? "";
                    if (!string.IsNullOrWhiteSpace(fullParaText) && fullParaText != combined)
                        results.Add(fullParaText);
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
                    if (matched.Contains(row)) continue;

                    var sowNorm = Normalize(row.SowText);
                    if (string.IsNullOrWhiteSpace(sowNorm))
                        continue;

                    if (highlightNorm.Contains(sowNorm) ||
                        sowNorm.Contains(highlightNorm))
                    {
                        matched.Add(row);
                    }
                }
            }

            // Second pass: match unmatched rows using partial overlap
            // (handles cases where only a portion of the SowText is highlighted in Word,
            //  or the highlighted block is a larger paragraph containing the SowText content)
            foreach (var row in excelRows)
            {
                if (row.RowNumber == 0) continue;
                if (row.IsSectionMarker) continue;
                if (matched.Contains(row)) continue;

                var sowNorm = Normalize(row.SowText);
                if (string.IsNullOrWhiteSpace(sowNorm)) continue;

                // Extract significant words (skip short/common words and bracket content)
                var sowWords = sowNorm
                    .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                    .Where(w => w.Length > 3 && !w.StartsWith("[") && !w.StartsWith("{"))
                    .ToArray();

                // Build a 40+ char substring from the middle of the SowText for robust matching
                var sowMid = sowNorm.Length > 80
                    ? sowNorm.Substring(20, 60)
                    : (sowNorm.Length > 40 ? sowNorm.Substring(10, Math.Min(50, sowNorm.Length - 10)) : sowNorm);

                foreach (var highlight in highlightedList)
                {
                    var highlightNorm = Normalize(highlight);
                    if (string.IsNullOrWhiteSpace(highlightNorm)) continue;

                    // Strategy: check if a meaningful substring of the SowText exists in the highlight
                    if (sowMid.Length >= 20 && highlightNorm.Contains(sowMid))
                    {
                        matched.Add(row);
                        _logger.LogInformation("PartialMatch: Row={Row}, Clause={Clause} matched via mid-substring",
                            row.RowNumber, row.ClauseNumber);
                        break;
                    }

                    // Strategy: check if a meaningful substring of the highlight exists in the SowText
                    var hlMid = highlightNorm.Length > 80
                        ? highlightNorm.Substring(20, 60)
                        : (highlightNorm.Length > 40 ? highlightNorm.Substring(10, Math.Min(50, highlightNorm.Length - 10)) : highlightNorm);
                    if (hlMid.Length >= 20 && sowNorm.Contains(hlMid))
                    {
                        matched.Add(row);
                        _logger.LogInformation("PartialMatch: Row={Row}, Clause={Clause} matched via highlight mid-substring",
                            row.RowNumber, row.ClauseNumber);
                        break;
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
            List<SowUiRow> matchedRows,
            List<SowUiRow> tableDefRows)
        {
            using var ms = new MemoryStream();
            ms.Write(wordData);
            ms.Position = 0;

            using (var doc = WordprocessingDocument.Open(ms, true))
            {
                var body = doc.MainDocumentPart!.Document.Body;

                RemoveSchedulesByHeading(body, matchedRows);
                RemoveSectionMarkers(body, matchedRows);
                RemoveTableDefinitionsByParent(body, matchedRows, tableDefRows);
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

                        // Skip modification for blue-shaded regions (reserved for manual review)
                        if (runs.All(IsRunBlueShaded))
                            continue;

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

                // Remove any unfilled [placeholder] brackets left in the document
                RemoveRemainingBracketPlaceholders(body);

                // Remove any remaining *{}* append placeholders
                RemoveRemainingAppendPlaceholders(body);

                // Strip escape markers: *[text]* → [text]
                RemoveEscapeMarkersFromDocument(body);

                // Remove all remaining highlight formatting except blue shading
                var remainingHighlightedRuns = body.Descendants<Run>()
                    .Where(r => IsRunHighlighted(r) && !IsRunBlueShaded(r)).ToList();
                RemoveHighlightFormattingFromRuns(remainingHighlightedRuns);

                RemoveNoteToDraftFromDocument(body);
                StripAllRemainingMarkers(body);
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
                                        SafeRemoveElement(elements[j]);

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
                        SafeRemoveElement(elements[i]);

                    // Handle end element
                    if (endIdx != startIdx)
                    {
                        if (endShared)
                            StripMarkerTextFromElement(elements[endIdx], markerText);
                        else
                            SafeRemoveElement(elements[endIdx]);
                    }

                    // Handle start element
                    if (startShared)
                        StripMarkerTextFromElement(elements[startIdx], markerText);
                    else
                        SafeRemoveElement(elements[startIdx]);

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
                        SafeRemoveElement(elements[endIdx]);
                    if (string.IsNullOrWhiteSpace(elements[startIdx].InnerText))
                        SafeRemoveElement(elements[startIdx]);
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
                    SafeRemoveParagraph(para);
            }
        }

        private static void StripAllRemainingMarkers(Body body)
        {
            var markerRegex = new Regex(@"(\*{5}[^*]+\*{5})|(&{5}[^&]+&{5})", RegexOptions.Compiled);

            foreach (var para in body.Descendants<Paragraph>().ToList())
            {
                var fullText = para.InnerText;
                if (string.IsNullOrEmpty(fullText)) continue;
                if (!fullText.Contains("*****") && !fullText.Contains("&&&&&")) continue;

                var cleaned = markerRegex.Replace(fullText, "").Trim();
                if (string.IsNullOrWhiteSpace(cleaned))
                {
                    SafeRemoveParagraph(para);
                }
                else
                {
                    var runs = para.Descendants<Run>().ToList();
                    RemovePatternAcrossRuns(runs, markerRegex);
                }
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
                    SafeRemoveParagraph(para);
            }
        }

        #endregion

        #region Remove Table Definitions by Parent

        private void RemoveTableDefinitionsByParent(
            Body body,
            List<SowUiRow> matchedRows,
            List<SowUiRow> tableDefRows)
        {
            if (tableDefRows == null || tableDefRows.Count == 0) return;

            var noParents = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var row in matchedRows)
            {
                var answer = row.UserAnswer?.Trim() ?? "";
                if (!answer.Equals("No", StringComparison.OrdinalIgnoreCase)) continue;

                if (!string.IsNullOrWhiteSpace(row.ClauseNumber))
                    noParents.Add(row.ClauseNumber.Trim());

                if (row.IsSectionMarker && !string.IsNullOrWhiteSpace(row.SectionMarkerName))
                    noParents.Add(row.SectionMarkerName.Trim());
            }

            if (noParents.Count == 0) return;

            var rowsToRemove = tableDefRows
                .Where(tr =>
                    !string.IsNullOrWhiteSpace(tr.ParentClauses) &&
                    tr.ParentClauses.Split(',')
                        .Select(p => p.Trim())
                        .Any(p => noParents.Contains(p)))
                .ToList();

            if (rowsToRemove.Count == 0) return;

            _logger.LogInformation(
                "RemoveTableDefinitionsByParent: {Count} table rows to remove (noParents: {Parents})",
                rowsToRemove.Count, string.Join(", ", noParents));

            // Pre-index all Word definition table rows: (termNorm, defVisNorm, TableRow reference)
            var wordTableIndex = new List<(string TermNorm, string DefNorm, TableRow Row)>();
            foreach (var table in body.Descendants<Table>())
            {
                var rows = table.Descendants<TableRow>().ToList();
                if (rows.Count < 5) continue; // skip small tables
                foreach (var wordRow in rows)
                {
                    var cells = wordRow.Descendants<TableCell>().ToList();
                    if (cells.Count < 2) continue;
                    var termVis = GetVisibleText(cells[0]);
                    var defVis = GetVisibleText(cells[1]);
                    wordTableIndex.Add((Normalize(termVis), Normalize(defVis), wordRow));
                }
            }

            _logger.LogInformation("Word table index built: {Count} rows across definition tables", wordTableIndex.Count);

            foreach (var defRow in rowsToRemove)
            {
                var rawSummary = defRow.SowSummary ?? "";
                var summaryClean = rawSummary.Trim().Trim(',', ' ', '\t').Trim();
                var summaryNorm = Normalize(summaryClean);
                var sowNorm = Normalize(defRow.SowText);

                _logger.LogInformation(
                    "Looking for: SummaryRaw='{Raw}', SummaryClean='{Clean}', SummaryNorm='{Norm}', SowNorm='{Sow}'",
                    rawSummary.Length > 60 ? rawSummary.Substring(0, 60) : rawSummary,
                    summaryClean,
                    summaryNorm,
                    sowNorm.Length > 80 ? sowNorm.Substring(0, 80) : sowNorm);

                if (string.IsNullOrWhiteSpace(sowNorm) && string.IsNullOrWhiteSpace(summaryNorm))
                    continue;

                bool removed = false;

                for (int wi = wordTableIndex.Count - 1; wi >= 0; wi--)
                {
                    var (termNorm, defNorm, wordRow) = wordTableIndex[wi];
                    bool matched = false;

                    // Strategy 1: exact term name match
                    if (!string.IsNullOrWhiteSpace(summaryNorm) &&
                        !string.IsNullOrWhiteSpace(termNorm) &&
                        termNorm == summaryNorm)
                    {
                        matched = true;
                    }

                    // Strategy 2: term name contains or is contained by summary
                    if (!matched && !string.IsNullOrWhiteSpace(summaryNorm) &&
                        !string.IsNullOrWhiteSpace(termNorm) &&
                        (termNorm.Contains(summaryNorm) || summaryNorm.Contains(termNorm)))
                    {
                        // Only accept if lengths are close enough to avoid false positives
                        if (Math.Abs(termNorm.Length - summaryNorm.Length) <= 5)
                            matched = true;
                    }

                    // Strategy 3: definition text match (using contains, handles field codes)
                    if (!matched && !string.IsNullOrWhiteSpace(sowNorm) &&
                        !string.IsNullOrWhiteSpace(defNorm))
                    {
                        // Use first 60 chars of sow text for robust matching
                        var sowPrefix = sowNorm.Length > 60 ? sowNorm.Substring(0, 60) : sowNorm;
                        var defPrefix = defNorm.Length > 60 ? defNorm.Substring(0, 60) : defNorm;
                        if (defNorm.Contains(sowPrefix) || sowNorm.Contains(defPrefix))
                            matched = true;
                    }

                    if (matched)
                    {
                        _logger.LogInformation(
                            "REMOVING table row: Parent='{Parent}', Term='{Term}', WordTerm='{WordTerm}'",
                            defRow.ParentClauses, summaryClean, termNorm);
                        wordRow.Remove();
                        wordTableIndex.RemoveAt(wi);
                        removed = true;
                        break;
                    }
                }

                if (!removed)
                {
                    _logger.LogWarning(
                        "NOT FOUND in document: Parent='{Parent}', SummaryNorm='{Summary}', SowNorm='{Sow}'",
                        defRow.ParentClauses, summaryNorm,
                        sowNorm.Length > 80 ? sowNorm.Substring(0, 80) : sowNorm);
                }
            }
        }

        /// <summary>
        /// Extracts only visible text from an OpenXML element, stripping Word field
        /// instruction codes (e.g. REF, MERGEFORMAT cross-references) that pollute InnerText.
        /// </summary>
        private static string GetVisibleText(OpenXmlElement element)
        {
            var sb = new StringBuilder();
            bool skipContent = false;

            foreach (var descendant in element.Descendants())
            {
                if (descendant is FieldChar fc)
                {
                    var charType = fc.FieldCharType?.Value;
                    if (charType == FieldCharValues.Begin)
                        skipContent = true;
                    else if (charType == FieldCharValues.Separate)
                        skipContent = false;
                    else if (charType == FieldCharValues.End)
                        skipContent = false;
                }
                else if (!skipContent && descendant is Text t)
                {
                    sb.Append(t.Text);
                }
            }

            return sb.ToString();
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
                        SafeRemoveParagraph(para);
                }
                else
                {
                    SafeRemoveParagraph(para);
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
                RemoveUnfilledPlaceholders(runs, matchedRow);
                if (matchedRow.HasAppendPlaceholder)
                    ReplaceAppendInRuns(runs, matchedRow.AppendText);
                RemoveHighlightFormattingFromRuns(runs);
            }
            else if (answer.Equals("Yes", StringComparison.OrdinalIgnoreCase))
            {
                RemoveUnfilledPlaceholders(runs, matchedRow);
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

        #region Safe Removal Helpers

        /// <summary>
        /// Safely removes a paragraph, ensuring table cells keep at least one paragraph
        /// and the body's last sectPr is preserved. Prevents "unreadable content" corruption.
        /// </summary>
        private static void SafeRemoveParagraph(Paragraph para)
        {
            if (para.Parent == null) return;

            var cell = para.Ancestors<TableCell>().FirstOrDefault();
            if (cell != null)
            {
                // Table cells MUST have at least one paragraph in OpenXML
                if (cell.Elements<Paragraph>().Count() <= 1)
                {
                    // Clear content instead of removing
                    foreach (var run in para.Descendants<Run>().ToList())
                        run.Remove();
                    return;
                }
            }

            var body = para.Ancestors<Body>().FirstOrDefault();
            if (body != null)
            {
                // Preserve section properties on the last body paragraph
                var sectPr = para.Descendants<SectionProperties>().FirstOrDefault()
                          ?? para.ParagraphProperties?.Descendants<SectionProperties>().FirstOrDefault();
                if (sectPr != null)
                {
                    // Move sectPr to the previous paragraph or body
                    var prevPara = para.PreviousSibling<Paragraph>();
                    if (prevPara != null)
                    {
                        if (prevPara.ParagraphProperties == null)
                            prevPara.ParagraphProperties = new ParagraphProperties();
                        prevPara.ParagraphProperties.AppendChild(sectPr.CloneNode(true));
                    }
                    else
                    {
                        body.AppendChild(sectPr.CloneNode(true));
                    }
                }

                // Don't remove if it's the very last paragraph in the body
                if (body.Elements<Paragraph>().Count() <= 1)
                {
                    foreach (var run in para.Descendants<Run>().ToList())
                        run.Remove();
                    return;
                }
            }

            para.Remove();
        }

        /// <summary>
        /// Safely removes an element, with special handling if it's a paragraph.
        /// </summary>
        private static void SafeRemoveElement(OpenXmlElement element)
        {
            if (element is Paragraph para)
                SafeRemoveParagraph(para);
            else
                element.Remove();
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

        private static bool IsRunBlueShaded(Run run)
        {
            var rPr = run.RunProperties;
            if (rPr == null) return false;

            foreach (var hl in rPr.Elements<Highlight>())
            {
                var val = hl.Val?.Value;
                if (val == HighlightColorValues.Blue ||
                    val == HighlightColorValues.DarkBlue ||
                    val == HighlightColorValues.Cyan ||
                    val == HighlightColorValues.DarkCyan)
                    return true;
            }

            foreach (var shd in rPr.Elements<Shading>())
            {
                var fill = shd.Fill?.Value;
                if (!string.IsNullOrEmpty(fill) && IsBlueColor(fill))
                    return true;
            }

            return false;
        }

        private static bool IsBlueColor(string hexColor)
        {
            hexColor = hexColor.Trim().TrimStart('#').ToUpperInvariant();
            if (hexColor == "AUTO" || hexColor == "FFFFFF" || hexColor.Length != 6)
                return false;

            if (!int.TryParse(hexColor.Substring(0, 2), System.Globalization.NumberStyles.HexNumber, null, out int r))
                return false;
            if (!int.TryParse(hexColor.Substring(2, 2), System.Globalization.NumberStyles.HexNumber, null, out int g))
                return false;
            if (!int.TryParse(hexColor.Substring(4, 2), System.Globalization.NumberStyles.HexNumber, null, out int b))
                return false;

            // Blue family: blue channel dominates red and green
            return b > r && b > g && b >= 100;
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

        private void RemoveUnfilledPlaceholders(List<Run> runs, SowUiRow row)
        {
            if (runs.Count == 0) return;

            var phInfos = row.GetPlaceholderInfos();
            if (phInfos.Count == 0) return;

            var fullText = new StringBuilder();
            foreach (var run in runs)
                fullText.Append(run.InnerText);
            var text = fullText.ToString();

            // Remove each unfilled placeholder from the text
            foreach (var ph in phInfos)
            {
                text = ReplaceIgnoreCase(text, ph.FullText, "");
                if (ph.IsNested)
                {
                    foreach (var inner in ph.InnerPlaceholders)
                        text = ReplaceIgnoreCase(text, inner, "");
                }
            }

            // Clean up double spaces left by removal
            text = Regex.Replace(text, @"  +", " ");

            // Write back
            if (runs.Count > 0)
            {
                var textEl = runs[0].GetFirstChild<Text>();
                if (textEl != null)
                {
                    textEl.Text = text;
                    textEl.Space = SpaceProcessingModeValues.Preserve;
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
            var noteToDraftPattern = new Regex(@"\s?\[NOTE TO DRAFT[^\]]*\]", RegexOptions.IgnoreCase);
            var optionalPattern = new Regex(@"\s?\[Optional[^\]]*\]", RegexOptions.IgnoreCase);

            foreach (var para in body.Descendants<Paragraph>().ToList())
            {
                var runs = para.Descendants<Run>().ToList();
                if (runs.Count == 0) continue;

                var fullText = string.Join("", runs.Select(r => r.InnerText));
                if (string.IsNullOrEmpty(fullText)) continue;

                bool changed = RemovePatternAcrossRunsPreserveBlue(runs, noteToDraftPattern);
                changed |= RemovePatternAcrossRunsPreserveBlue(runs, optionalPattern);

                if (!changed) continue;

                var remaining = string.Join("", runs.Select(r => r.InnerText));
                if (string.IsNullOrWhiteSpace(remaining))
                {
                    foreach (var run in runs)
                        run.Remove();
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
                if (IsRunBlueShaded(run)) continue;
                var textEl = run.GetFirstChild<Text>();
                if (textEl == null) continue;
                var original = textEl.Text;
                if (string.IsNullOrEmpty(original) || !original.Contains("*{}*")) continue;
                textEl.Text = Regex.Replace(original, @"\s?\*\{\}\*", "");
                textEl.Space = SpaceProcessingModeValues.Preserve;
            }
        }

        /// <summary>
        /// Removes any unfilled [placeholder] bracket text remaining in the document.
        /// Skips escaped *[text]*, [NOTE TO DRAFT...], and [Optional...] patterns
        /// which are handled separately.
        /// </summary>
        private void RemoveRemainingBracketPlaceholders(Body body)
        {
            // Matches both nested [outer [inner] text] and simple [placeholder] brackets,
            // but excludes [NOTE TO DRAFT...] and [Optional...] which are handled elsewhere.
            var nestedPattern = new Regex(@"\s?\[(?!NOTE TO DRAFT|Optional)(?:[^\[\]]*\[[^\[\]]*\])*[^\[\]]*\]", RegexOptions.IgnoreCase);
            var simplePattern = new Regex(@"\s?\[(?!NOTE TO DRAFT|Optional)[^\[\]]+\]", RegexOptions.IgnoreCase);

            foreach (var para in body.Descendants<Paragraph>().ToList())
            {
                var runs = para.Descendants<Run>().ToList();
                if (runs.Count == 0) continue;

                var fullText = string.Join("", runs.Select(r => r.InnerText));
                if (string.IsNullOrEmpty(fullText) || !fullText.Contains('[')) continue;

                // First strip escaped *[text]* markers so they don't match
                var testText = Regex.Replace(fullText, @"\*\[[^\]]+\]\*", "");
                if (!testText.Contains('[')) continue;

                // Remove nested brackets first, then simple ones (preserving blue-shaded content)
                if (nestedPattern.IsMatch(testText))
                    RemovePatternAcrossRunsPreserveBlue(runs, nestedPattern);
                // Re-check after nested removal
                fullText = string.Join("", runs.Select(r => r.InnerText));
                testText = Regex.Replace(fullText, @"\*\[[^\]]+\]\*", "");
                if (simplePattern.IsMatch(testText))
                    RemovePatternAcrossRunsPreserveBlue(runs, simplePattern);
            }
        }

        private void RemoveEscapeMarkersFromDocument(Body body)
        {
            foreach (var run in body.Descendants<Run>().ToList())
            {
                if (IsRunBlueShaded(run)) continue;
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

            foreach (var para in body.Descendants<Paragraph>().ToList())
            {
                var runs = para.Descendants<Run>().ToList();
                if (runs.Count == 0) continue;

                var fullText = string.Join("", runs.Select(r => r.InnerText));
                if (string.IsNullOrEmpty(fullText)) continue;

                bool hasVariable = false;
                foreach (var kvp in variableMap)
                {
                    if (fullText.Contains(kvp.Key, StringComparison.OrdinalIgnoreCase))
                    {
                        hasVariable = true;
                        break;
                    }
                }
                if (!hasVariable) continue;

                foreach (var kvp in variableMap)
                {
                    while (ReplaceTextAcrossRuns(runs, kvp.Key, kvp.Value)) { }
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

        /// <summary>
        /// Removes all matches of a regex pattern from paragraph runs while
        /// preserving each run's formatting. Only the runs overlapping a match
        /// are modified; all other runs stay untouched.
        /// </summary>
        private static bool RemovePatternAcrossRuns(List<Run> runs, Regex pattern)
        {
            bool anyRemoved = false;
            bool found;
            do
            {
                found = false;
                var fullSb = new StringBuilder();
                var runStarts = new List<int>();
                foreach (var run in runs)
                {
                    runStarts.Add(fullSb.Length);
                    fullSb.Append(run.InnerText);
                }
                var fullText = fullSb.ToString();

                var match = pattern.Match(fullText);
                if (!match.Success) break;

                found = true;
                anyRemoved = true;

                int matchStart = match.Index;
                int matchEnd = match.Index + match.Length;

                for (int ri = 0; ri < runs.Count; ri++)
                {
                    int rStart = runStarts[ri];
                    int rEnd = ri + 1 < runs.Count ? runStarts[ri + 1] : fullText.Length;
                    if (rEnd <= matchStart || rStart >= matchEnd) continue;

                    var textEl = runs[ri].GetFirstChild<Text>();
                    if (textEl == null) continue;

                    int localStart = Math.Max(0, matchStart - rStart);
                    int localEnd = Math.Min(textEl.Text.Length, matchEnd - rStart);
                    textEl.Text = textEl.Text.Substring(0, localStart) + textEl.Text.Substring(localEnd);
                    textEl.Space = SpaceProcessingModeValues.Preserve;
                }
            } while (found);

            return anyRemoved;
        }

        /// <summary>
        /// Blue-shading-aware version of RemovePatternAcrossRuns.
        /// Skips any match that overlaps a blue-shaded run, preserving blue content.
        /// </summary>
        private static bool RemovePatternAcrossRunsPreserveBlue(List<Run> runs, Regex pattern)
        {
            bool anyRemoved = false;
            bool progress;
            do
            {
                progress = false;
                var fullSb = new StringBuilder();
                var runStarts = new List<int>();
                foreach (var run in runs)
                {
                    runStarts.Add(fullSb.Length);
                    fullSb.Append(run.InnerText);
                }
                var fullText = fullSb.ToString();

                foreach (Match match in pattern.Matches(fullText))
                {
                    int matchStart = match.Index;
                    int matchEnd = match.Index + match.Length;

                    bool overlapsBlue = false;
                    for (int ri = 0; ri < runs.Count; ri++)
                    {
                        int rStart = runStarts[ri];
                        int rEnd = ri + 1 < runs.Count ? runStarts[ri + 1] : fullText.Length;
                        if (rEnd <= matchStart || rStart >= matchEnd) continue;
                        if (IsRunBlueShaded(runs[ri]))
                        {
                            overlapsBlue = true;
                            break;
                        }
                    }

                    if (overlapsBlue) continue;

                    for (int ri = 0; ri < runs.Count; ri++)
                    {
                        int rStart = runStarts[ri];
                        int rEnd = ri + 1 < runs.Count ? runStarts[ri + 1] : fullText.Length;
                        if (rEnd <= matchStart || rStart >= matchEnd) continue;

                        var textEl = runs[ri].GetFirstChild<Text>();
                        if (textEl == null) continue;

                        int localStart = Math.Max(0, matchStart - rStart);
                        int localEnd = Math.Min(textEl.Text.Length, matchEnd - rStart);
                        textEl.Text = textEl.Text.Substring(0, localStart) + textEl.Text.Substring(localEnd);
                        textEl.Space = SpaceProcessingModeValues.Preserve;
                    }
                    progress = true;
                    anyRemoved = true;
                    break; // Restart since offsets changed
                }
            } while (progress);

            return anyRemoved;
        }

        /// <summary>
        /// Replaces one occurrence of oldValue across paragraph runs while
        /// preserving each run's formatting. The replacement text inherits
        /// the formatting of the run where the match starts.
        /// </summary>
        private static bool ReplaceTextAcrossRuns(List<Run> runs, string oldValue, string newValue,
            StringComparison comparison = StringComparison.OrdinalIgnoreCase)
        {
            var fullSb = new StringBuilder();
            var runStarts = new List<int>();
            foreach (var run in runs)
            {
                runStarts.Add(fullSb.Length);
                fullSb.Append(run.InnerText);
            }
            var fullText = fullSb.ToString();

            int idx = fullText.IndexOf(oldValue, comparison);
            if (idx < 0) return false;

            int matchEnd = idx + oldValue.Length;

            for (int ri = 0; ri < runs.Count; ri++)
            {
                int rStart = runStarts[ri];
                int rEnd = ri + 1 < runs.Count ? runStarts[ri + 1] : fullText.Length;
                if (rEnd <= idx || rStart >= matchEnd) continue;

                var textEl = runs[ri].GetFirstChild<Text>();
                if (textEl == null) continue;

                int localStart = Math.Max(0, idx - rStart);
                int localEnd = Math.Min(textEl.Text.Length, matchEnd - rStart);

                bool isMatchStart = (rStart <= idx && idx < rEnd);
                if (isMatchStart)
                    textEl.Text = textEl.Text.Substring(0, localStart) + newValue + textEl.Text.Substring(localEnd);
                else
                    textEl.Text = textEl.Text.Substring(0, localStart) + textEl.Text.Substring(localEnd);

                textEl.Space = SpaceProcessingModeValues.Preserve;
            }

            return true;
        }

        #endregion
    }
}