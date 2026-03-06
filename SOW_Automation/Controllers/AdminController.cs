using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using SowAutomationTool.Data;
using SowAutomationTool.Models;

namespace SowAutomationTool.Controllers
{
    [Authorize(Roles = "Admin")]
    public class AdminController : Controller
    {
        private readonly AppDbContext _db;

        public AdminController(AppDbContext db)
        {
            _db = db;
        }

        [HttpGet]
        public async Task<IActionResult> Users()
        {
            var users = await _db.Users.OrderBy(u => u.Email).ToListAsync();
            return View(users);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> AddUser(string email, string displayName, string role)
        {
            if (string.IsNullOrWhiteSpace(email))
            {
                TempData["Error"] = "Email is required.";
                return RedirectToAction(nameof(Users));
            }

            email = email.Trim().ToLower();
            role = role == "Admin" ? "Admin" : "User";

            var exists = await _db.Users.AnyAsync(u => u.Email.ToLower() == email);
            if (exists)
            {
                TempData["Error"] = $"User '{email}' already exists.";
                return RedirectToAction(nameof(Users));
            }

            _db.Users.Add(new AppUser
            {
                Email = email,
                DisplayName = displayName?.Trim() ?? "",
                Role = role,
                IsActive = true
            });
            await _db.SaveChangesAsync();

            TempData["Success"] = $"User '{email}' added as {role}.";
            return RedirectToAction(nameof(Users));
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> ToggleActive(int id)
        {
            var user = await _db.Users.FindAsync(id);
            if (user == null)
                return NotFound();

            user.IsActive = !user.IsActive;
            await _db.SaveChangesAsync();

            TempData["Success"] = $"User '{user.Email}' is now {(user.IsActive ? "active" : "inactive")}.";
            return RedirectToAction(nameof(Users));
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteUser(int id)
        {
            var user = await _db.Users.FindAsync(id);
            if (user == null)
                return NotFound();

            // Prevent deleting yourself
            var currentEmail = User.FindFirst("preferred_username")?.Value
                            ?? User.FindFirst(System.Security.Claims.ClaimTypes.Email)?.Value
                            ?? "";
            if (user.Email.Equals(currentEmail, StringComparison.OrdinalIgnoreCase))
            {
                TempData["Error"] = "You cannot delete your own account.";
                return RedirectToAction(nameof(Users));
            }

            _db.Users.Remove(user);
            await _db.SaveChangesAsync();

            TempData["Success"] = $"User '{user.Email}' deleted.";
            return RedirectToAction(nameof(Users));
        }
    }
}
