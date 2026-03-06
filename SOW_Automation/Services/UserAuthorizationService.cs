using Microsoft.EntityFrameworkCore;
using SowAutomationTool.Data;
using SowAutomationTool.Models;

namespace SowAutomationTool.Services
{
    public class UserAuthorizationService
    {
        private readonly AppDbContext _db;

        public UserAuthorizationService(AppDbContext db)
        {
            _db = db;
        }

        public async Task<bool> IsAuthorizedAsync(string? email)
        {
            if (string.IsNullOrWhiteSpace(email))
                return false;

            return await _db.Users
                .AnyAsync(u => u.Email.ToLower() == email.ToLower() && u.IsActive);
        }

        public async Task<AppUser?> GetUserAsync(string? email)
        {
            if (string.IsNullOrWhiteSpace(email))
                return null;

            return await _db.Users
                .FirstOrDefaultAsync(u => u.Email.ToLower() == email.ToLower() && u.IsActive);
        }
    }
}
