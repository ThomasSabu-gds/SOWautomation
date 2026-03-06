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

            // ✅ Create workflow id & cache state
            var workflowId = Guid.NewGuid().ToString("N");

            _cache.Set(WfKey(workflowId, "word"), wordBytes, CacheTtl);
            _cache.Set(WfKey(workflowId, "rows"), matchedRows, CacheTtl);

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

            // Save updated answers
            _cache.Set(WfKey(id, "rows"), rows ?? new List<SowUiRow>(), CacheTtl);

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

            var finalDoc = _service.GenerateDocument(wordBytes, rows);

            return File(finalDoc,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "FinalSOW.docx");
        }
    }
}