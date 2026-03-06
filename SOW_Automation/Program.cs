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
builder.Services.AddScoped<ProcessingService>();
builder.Services.AddScoped<UserAuthorizationService>();

var app = builder.Build();

// Seed DB
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
    db.Database.EnsureCreated();

    if (!db.Users.Any())
    {
        var seedEmail = builder.Configuration["SeedUser:Email"] ?? "admin@ey.com";
        var seedName = builder.Configuration["SeedUser:DisplayName"] ?? "Default Admin";
        var seedRole = builder.Configuration["SeedUser:Role"] ?? "Admin";

        db.Users.Add(new AppUser
        {
            Email = seedEmail,
            DisplayName = seedName,
            Role = seedRole,
            IsActive = true
        });

        db.SaveChanges();
    }
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
app.UseAuthorization();

// ✅ Your custom middleware AFTER auth (so it can read user/claims)
app.UseMiddleware<UserAuthorizationMiddleware>();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.MapRazorPages();

app.Run();