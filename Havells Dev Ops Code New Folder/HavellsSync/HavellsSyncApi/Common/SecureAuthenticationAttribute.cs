using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.ICommon;
using HavellsSync_ModelData.Product;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.DataProtection.KeyManagement;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.Crm.Sdk.Messages;
using Newtonsoft.Json;
using System.ComponentModel.Design;
using System.Configuration;
using System.Net;
using System.Net.Mime;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using static System.Net.Mime.MediaTypeNames;

namespace HavellsSyncApi.Common
{
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Method)]
    public class CustomAuthorizeAttribute : Attribute, IAuthorizationFilter
    {
        public void OnAuthorization(AuthorizationFilterContext context)
        {
            if (string.IsNullOrEmpty(context.HttpContext.Request.Headers["LoginUserId"]) || string.IsNullOrEmpty(context.HttpContext.Request.Headers["AccessToken"]))
            {
                context.HttpContext.Response.StatusCode = StatusCodes.Status401Unauthorized;
                context.Result = new UnauthorizedObjectResult("+iXfS0L9p+XND5stK1a3qc+e62lOf7JUuK8eM6pt5SHmXwmdQXcyClO0DKAK50VXCWS/qhD+IusHgnei/yO1JyUJQCBbTvRQoPXh3S3x4Qo=");
            }
            else
            {
                //skip authorization if action is decorated with [AllowAnonymous] attribute
                var allowAnonymous = context.ActionDescriptor.EndpointMetadata.OfType<AllowAnonymousAttribute>().Any();
                if (allowAnonymous)
                    return;
                validatesessionResponse? objdata = (validatesessionResponse?)context.HttpContext.Items["AuthInfo"];
                if (objdata.StatusCode == 400)
                {
                    // not logged in - return 401 unauthorized
                    //context.Result = new JsonResult(objdata.Message) { StatusCode = StatusCodes.Status200OK };
                    context.HttpContext.Response.StatusCode = StatusCodes.Status401Unauthorized;
                    context.Result = new UnauthorizedObjectResult(objdata.Message);
                    // set 'WWW-Authenticate' header to trigger login popup in browsers
                    context.HttpContext.Response.Headers["WWW-Authenticate"] = "Basic realm=\"\", charset=\"UTF-8\"";
                }
            }
        }
    }
}
