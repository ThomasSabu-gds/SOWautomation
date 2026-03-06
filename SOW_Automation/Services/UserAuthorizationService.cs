using Microsoft.EntityFrameworkCore;
using SowAutomationTool.Data;

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
    }
}
