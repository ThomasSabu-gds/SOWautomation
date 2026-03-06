using Microsoft.AspNetCore.Authentication.OpenIdConnect;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc.Authorization;
using Microsoft.EntityFrameworkCore;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.UI;
using SowAutomationTool.Data;
using SowAutomationTool.Middleware;
using SowAutomationTool.Models;
using SowAutomationTool.Services;

var builder = WebApplication.CreateBuilder(args);

// ✅ Add MemoryCache (needed for workflow/token cache)
builder.Services.AddMemoryCache();

// Auth (Azure AD / Entra ID)
builder.Services.AddAuthentication(OpenIdConnectDefaults.AuthenticationScheme)
    .AddMicrosoftIdentityWebApp(builder.Configuration.GetSection("AzureAd"));

// Ensure ClaimTypes.Role is used for role checks (matches what the middleware injects)
builder.Services.Configure<OpenIdConnectOptions>(OpenIdConnectDefaults.AuthenticationScheme, options =>
{
    options.TokenValidationParameters.RoleClaimType = System.Security.Claims.ClaimTypes.Role;
});

// Global Authorization Policy (Require Authenticated User)
builder.Services.AddControllersWithViews(options =>
{
    var policy = new AuthorizationPolicyBuilder()
        .RequireAuthenticatedUser()
        .Build();
    options.Filters.Add(new AuthorizeFilter(policy));
});

builder.Services.AddRazorPages()
    .AddMicrosoftIdentityUI();

// DB Context
builder.Services.AddDbContext<AppDbContext>(options =>
    options.UseSqlite(builder.Configuration.GetConnectionString("UsersDb")));

// DI Services
builder.Services.AddSingleton<TemplateProvider>();
builder.Services.AddScoped<ProcessingService>();
builder.Services.AddScoped<UserAuthorizationService>();

var app = builder.Build();

// Seed DB
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
    db.Database.EnsureCreated();

    var seedUsers = new[]
    {
        new { Email = builder.Configuration["SeedUser:Email"] ?? "admin@ey.com",
              Name = builder.Configuration["SeedUser:DisplayName"] ?? "Default Admin" },
        new { Email = "christo.kl@gds.ey.com", Name = "Christo" },
        new { Email = "Aravind.Sathishan@gds.ey.com", Name = "Aravind" }
    };

    foreach (var seed in seedUsers)
    {
        if (!db.Users.Any(u => u.Email.ToLower() == seed.Email.ToLower()))
        {
            db.Users.Add(new AppUser
            {
                Email = seed.Email,
                DisplayName = seed.Name,
                Role = "Admin",
                IsActive = true
            });
        }
    }
    db.SaveChanges();
}

// Pipeline
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthentication();

// Inject app role claims BEFORE authorization checks them
app.UseMiddleware<UserAuthorizationMiddleware>();

app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.MapRazorPages();

app.Run();