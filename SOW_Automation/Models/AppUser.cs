
using System.ComponentModel.DataAnnotations;

namespace SowAutomationTool.Models
{
    public class AppUser
    {
        [Key]
        public int Id { get; set; }

        [Required]
        [MaxLength(256)]
        public string Email { get; set; } = string.Empty;

        [MaxLength(256)]
        public string DisplayName { get; set; } = string.Empty;

        [MaxLength(50)]
        public string Role { get; set; } = "User";

        public bool IsActive { get; set; } = true;

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }
}
