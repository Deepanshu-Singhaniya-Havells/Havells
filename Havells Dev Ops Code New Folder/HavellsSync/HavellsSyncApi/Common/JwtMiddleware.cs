using Microsoft.AspNetCore.DataProtection;
using Microsoft.Extensions.Options;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using HavellsSync_Busines.IService;
using Microsoft.Crm.Sdk.Messages;
using HavellsSync_ModelData.Common;

namespace HavellsSyncApi.Common
{
    public class JwtMiddleware
    {
        private readonly RequestDelegate _next;

        public JwtMiddleware(RequestDelegate next, IOptions<AppSettings> appSettings)
        {
            _next = next;
        }

        public async Task Invoke(HttpContext context, IAuthentication AuthService)
        {
            try
            {
                var username = ""; var password = ""; var LoginUserId = "";
                ValidateSession objSession = new ValidateSession();
                if (context.Request.Headers.ContainsKey("Authorization"))
                {
                    var authHeader = AuthenticationHeaderValue.Parse(context.Request.Headers["Authorization"]);
                    var credentialBytes = Convert.FromBase64String(authHeader.Parameter);
                    var credentials = Encoding.UTF8.GetString(credentialBytes).Split(':');
                    username = credentials[0];
                    password = credentials[1];
                    LoginUserId = context.Request.Headers["LoginUserId"];
                    objSession.SessionId = context.Request.Headers["AccessToken"];
                }
                else
                {
                    LoginUserId = context.Request.Headers["LoginUserId"];
                    objSession.SessionId = context.Request.Headers["AccessToken"];
                }
               // authenticate credentials with user service and attach user to http context
                if (!string.IsNullOrEmpty(objSession.SessionId) && !string.IsNullOrEmpty(LoginUserId))
                {
                    context.Items["AuthInfo"] = AuthService.ValidateSessionDetails(objSession, LoginUserId);
                }
            }
            catch
            {
                // do nothing if invalid auth header
                // user is not attached to context so request won't have access to secure routes
            }
            await _next(context);
        }
    }
}
