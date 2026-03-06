using System.Security.Claims;
using SowAutomationTool.Services;

namespace SowAutomationTool.Middleware
{
    public class UserAuthorizationMiddleware
    {
        private readonly RequestDelegate _next;

        private static readonly HashSet<string> BypassPaths = new(StringComparer.OrdinalIgnoreCase)
        {
            "/MicrosoftIdentity/Account/SignIn",
            "/MicrosoftIdentity/Account/SignOut",
            "/MicrosoftIdentity/Account/SignedOut",
            "/signin-oidc",
            "/signout-oidc",
            "/signout-callback-oidc",
            "/Account/AccessDenied"
        };

        public UserAuthorizationMiddleware(RequestDelegate next)
        {
            _next = next;
        }

        public async Task InvokeAsync(HttpContext context, UserAuthorizationService authService)
        {
            var path = context.Request.Path.Value ?? "";

            if (BypassPaths.Contains(path) ||
                path.StartsWith("/css", StringComparison.OrdinalIgnoreCase) ||
                path.StartsWith("/js", StringComparison.OrdinalIgnoreCase) ||
                path.StartsWith("/lib", StringComparison.OrdinalIgnoreCase) ||
                path.StartsWith("/favicon", StringComparison.OrdinalIgnoreCase))
            {
                await _next(context);
                return;
            }

            if (context.User.Identity?.IsAuthenticated == true)
            {
                var email = context.User.FindFirst("preferred_username")?.Value
                         ?? context.User.FindFirst(ClaimTypes.Email)?.Value
                         ?? context.User.FindFirst("email")?.Value;

                var appUser = await authService.GetUserAsync(email);
                if (appUser == null)
                {
                    context.Response.Redirect("/Account/AccessDenied");
                    return;
                }

                // Inject the app role as a claim so controllers can check it
                var identity = context.User.Identity as ClaimsIdentity;
                if (identity != null && !context.User.IsInRole(appUser.Role))
                {
                    identity.AddClaim(new Claim(ClaimTypes.Role, appUser.Role));
                }
            }

            await _next(context);
        }
    }
}
