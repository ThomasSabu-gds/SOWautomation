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
                         ?? context.User.FindFirst(System.Security.Claims.ClaimTypes.Email)?.Value
                         ?? context.User.FindFirst("email")?.Value;

                if (!await authService.IsAuthorizedAsync(email))
                {
                    context.Response.Redirect("/Account/AccessDenied");
                    return;
                }
            }

            await _next(context);
        }
    }
}
