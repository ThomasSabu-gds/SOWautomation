using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Caching.Memory;
using SowAutomationTool.Models;
using SowAutomationTool.Services;
using System.Security.Claims;

namespace SowAutomationTool.Controllers
{
    [Authorize]
    public class SowController : Controller
    {
        private readonly ProcessingService _service;
        private readonly IMemoryCache _cache;

        // cache lifetime for workflow data
        private static readonly TimeSpan CacheTtl = TimeSpan.FromMinutes(30);

        public SowController(ProcessingService service, IMemoryCache cache)
        {
            _service = service;
            _cache = cache;
        }

        // ---------------------------
        // Cache key helpers (per-user)
        // ---------------------------
        private string UserScope()
        {
            // per-user scope avoids collisions between different logged-in users
            return User.FindFirstValue(ClaimTypes.NameIdentifier)
                   ?? User.Identity?.Name
                   ?? "anonymous";
        }

        private string WfKey(string workflowId, string part) => $"sow:{UserScope()}:wf:{workflowId}:{part}";
        private string TokenKey(string token) => $"sow:{UserScope()}:tok:{token}";

        // ---------------------------
        // Step 1: Upload
        // ---------------------------
        [HttpGet]
        public IActionResult Upload()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Upload(IFormFile excelFile, IFormFile wordFile)
        {
            if (excelFile == null || wordFile == null)
                return View();

            using var excelStream = new MemoryStream();
            excelFile.CopyTo(excelStream);

            using var wordStream = new MemoryStream();
            wordFile.CopyTo(wordStream);

            var excelBytes = excelStream.ToArray();
            var wordBytes = wordStream.ToArray();

            // 1️⃣ Parse Excel
            var excelRows = _service.ParseExcel(excelBytes);

            // 2️⃣ Extract highlighted text from Word
            var highlightedText = _service.ExtractHighlightedText(wordBytes);

            // 3️⃣ Match Column D (SOW text) with highlighted text
            var matchedRows = _service.GetMatchedRows(excelRows, highlightedText)
                                      .OrderBy(r => r.RowNumber)
                                      .ToList();

            // Extract table definition rows (e.g. tableA) for parent-based removal
            var tableDefRows = excelRows.Where(r => r.IsTableRow).ToList();

            // ✅ Create workflow id & cache state
            var workflowId = Guid.NewGuid().ToString("N");

            _cache.Set(WfKey(workflowId, "word"), wordBytes, CacheTtl);
            _cache.Set(WfKey(workflowId, "rows"), matchedRows, CacheTtl);
            _cache.Set(WfKey(workflowId, "tableDefs"), tableDefRows, CacheTtl);

            // ✅ Redirect to Step-2 Create (GET)
            return RedirectToAction(nameof(Create), new { id = workflowId });
        }

        // ---------------------------
        // Step 2: Create (GET shows the table)
        // ---------------------------
        [HttpGet]
        public IActionResult Create(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
                return RedirectToAction(nameof(Upload));

            if (!_cache.TryGetValue(WfKey(id, "rows"), out List<SowUiRow>? rows) || rows == null)
                return RedirectToAction(nameof(Upload));

            ViewBag.WorkflowId = id;

            // If your view file is not Create.cshtml, change "Create" to your actual name
            return View("Create", rows);
        }

        // ---------------------------
        // Step 2: Create (POST saves answers and redirects to Step-3)
        // ---------------------------
        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Create(string id, List<SowUiRow> rows)
        {
            if (string.IsNullOrWhiteSpace(id))
                return RedirectToAction(nameof(Upload));

            // Ensure workflow exists
            if (!_cache.TryGetValue(WfKey(id, "word"), out byte[]? wordBytes) || wordBytes == null)
                return RedirectToAction(nameof(Upload));

            // Merge user answers into the original cached rows (preserves computed properties)
            if (_cache.TryGetValue(WfKey(id, "rows"), out List<SowUiRow>? cachedRows) && cachedRows != null && rows != null)
            {
                var formLookup = rows.ToDictionary(r => r.RowNumber);
                foreach (var cached in cachedRows)
                {
                    if (formLookup.TryGetValue(cached.RowNumber, out var formRow))
                    {
                        cached.UserAnswer = formRow.UserAnswer;
                        cached.PlaceholderAnswers = formRow.PlaceholderAnswers;
                        cached.AppendText = formRow.AppendText;
                    }
                }
                _cache.Set(WfKey(id, "rows"), cachedRows, CacheTtl);
            }
            else
            {
                _cache.Set(WfKey(id, "rows"), rows ?? new List<SowUiRow>(), CacheTtl);
            }

            // ✅ Create a download token and map token -> workflowId
            var token = Guid.NewGuid().ToString("N");
            _cache.Set(TokenKey(token), id, CacheTtl);

            // ✅ Redirect to Step-3 generate page (GET)
            return RedirectToAction(nameof(GeneratePage), new { id = id, token = token });
        }

        // ---------------------------
        // Step 3: Generate Page (GET shows download button + answers)
        // ---------------------------
        [HttpGet]
        public IActionResult GeneratePage(string id, string token)
        {
            if (string.IsNullOrWhiteSpace(id))
                return RedirectToAction(nameof(Upload));

            if (!_cache.TryGetValue(WfKey(id, "rows"), out List<SowUiRow>? rows) || rows == null)
                return RedirectToAction(nameof(Create), new { id });

            ViewBag.WorkflowId = id;
            ViewBag.Token = token ?? "";

            return View("Generate", rows);
        }

        // ---------------------------
        // Download (GET token-based)
        // Your Step-3 button calls: /Sow/Download/{token}
        // ---------------------------
        [HttpGet]
        public IActionResult Download(string id)
        {
            // id is token
            if (string.IsNullOrWhiteSpace(id))
                return RedirectToAction(nameof(Upload));

            // token -> workflowId
            if (!_cache.TryGetValue(TokenKey(id), out string? workflowId) || string.IsNullOrWhiteSpace(workflowId))
                return BadRequest("Invalid or expired token. Go back and generate again.");

            // workflow -> word + rows
            if (!_cache.TryGetValue(WfKey(workflowId, "word"), out byte[]? wordBytes) || wordBytes == null)
                return RedirectToAction(nameof(Upload));

            if (!_cache.TryGetValue(WfKey(workflowId, "rows"), out List<SowUiRow>? rows) || rows == null)
                return RedirectToAction(nameof(Create), new { id = workflowId });

            _cache.TryGetValue(WfKey(workflowId, "tableDefs"), out List<SowUiRow>? tableDefRows);

            // Propagate "No" from parents to unanswered children for document generation
            var noParents = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in rows)
            {
                if (r.UserAnswer?.Trim().Equals("No", StringComparison.OrdinalIgnoreCase) == true)
                {
                    if (!string.IsNullOrWhiteSpace(r.ClauseNumber))
                        noParents.Add(r.ClauseNumber.Trim());
                    if (r.IsSectionMarker && !string.IsNullOrWhiteSpace(r.SectionMarkerName))
                        noParents.Add(r.SectionMarkerName.Trim());
                }
            }
            if (noParents.Count > 0)
            {
                foreach (var r in rows)
                {
                    if (!string.IsNullOrWhiteSpace(r.UserAnswer)) continue;
                    if (string.IsNullOrWhiteSpace(r.ParentClauses)) continue;
                    var parents = r.ParentClauses.Split(',').Select(p => p.Trim());
                    if (parents.Any(p => noParents.Contains(p)))
                        r.UserAnswer = "No";
                }
            }

            var finalDoc = _service.GenerateDocument(wordBytes, rows, tableDefRows ?? new List<SowUiRow>());

            return File(finalDoc,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "FinalSOW.docx");
        }
    }
}