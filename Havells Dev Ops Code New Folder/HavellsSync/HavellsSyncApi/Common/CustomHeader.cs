using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc.ApiExplorer;
using Microsoft.AspNetCore.Mvc.Controllers;
using Microsoft.OpenApi.Any;
using Microsoft.OpenApi.Models;
using Newtonsoft.Json;
using Swashbuckle.AspNetCore.SwaggerGen;
using System.Collections.Generic;
using System.Reflection;
using System.Security.Policy;

namespace HavellsSyncApi.Common
{
    public class CustomHeader : IOperationFilter
    {
        public void Apply(OpenApiOperation operation, OperationFilterContext context)
        {

            if (operation.Parameters == null)
                operation.Parameters = new List<OpenApiParameter>();
            string[] ControllerName = { "Authenticate", "EasyReward", "Epos", "Consumer", "GrievanceHandling" };
            if (!ControllerName.Contains((context.ApiDescription.ActionDescriptor as ControllerActionDescriptor).ControllerName))
            {
                operation.Parameters.Add(new OpenApiParameter()
                {
                    Name = "LoginUserId",
                    Description = "Login User Id",
                    In = ParameterLocation.Header,
                    Schema = new OpenApiSchema() { Type = "string" },
                    Required = true,
                    //Example = new OpenApiString("9819292121")
                });

                operation.Parameters.Add(new OpenApiParameter()
                {
                    Name = "AccessToken",
                    Description = "Access Token",
                    In = ParameterLocation.Header,
                    Schema = new OpenApiSchema() { Type = "string" },
                    Required = true,
                    //Example = new OpenApiString("lrpwF2T/Kz1a3vJurVaDAXu6YEZae9gUxj1EHEWIdVu36/JLhKAohojyeCH5praa")
                });
                if (context.ApiDescription.ActionDescriptor is ControllerActionDescriptor descriptor)
                {
                    // If [AllowAnonymous] is not applied or [Authorize] or Custom Authorization filter is applied on either the endpoint or the controller
                    if (!context.ApiDescription.CustomAttributes().Any((a) => a is AllowAnonymousAttribute)
                        && (context.ApiDescription.CustomAttributes().Any((a) => a is AuthorizeAttribute)
                            || descriptor.ControllerTypeInfo.GetCustomAttributes<AuthorizeAttribute>() != null))
                    {
                        if (operation.Security == null)
                            operation.Security = new List<OpenApiSecurityRequirement>();

                        operation.Security.Add(
                            new OpenApiSecurityRequirement{
                {
                    new OpenApiSecurityScheme
                    {
                        Name = "",
                        In = ParameterLocation.Header,
                        BearerFormat = "Key",

                        Reference = new OpenApiReference
                        {
                            Type = ReferenceType.SecurityScheme,
                            Id = "Key"
                        }
                    },
                    new string[]{ }
                }
                        });
                    }
                }
            }
        }
    }
    public class SwaggerIgnoreFilter : ISchemaFilter
    {
        public void Apply(OpenApiSchema schema, SchemaFilterContext context)
        {
            if (schema?.Properties == null)
            {
                return;
            }

            const BindingFlags bindingFlags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;
            var memberList = context.Type
                .GetFields(bindingFlags).Cast<MemberInfo>()
                .Concat(context.Type.GetProperties(bindingFlags));

            var excludedList = memberList
                .Select(m => m.GetCustomAttribute<JsonPropertyAttribute>()?.PropertyName ?? m.Name);

            foreach (var excludedName in excludedList)
            {
                if (schema.Properties.ContainsKey(excludedName))
                    schema.Properties.Remove(excludedName);
            }
        }

        //public static string ToCamelCase(this string str)
        //{
        //    if (!string.IsNullOrEmpty(str) && str.Length > 1)
        //    {
        //        return char.ToLowerInvariant(str[0]) + str.Substring(1);
        //    }
        //    return str.ToLowerInvariant();
        //}
    }
}