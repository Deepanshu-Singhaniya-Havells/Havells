using HavellsSync_ModelData.Common;
using HavellsSyncApi.Common;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Options;
using Newtonsoft.Json.Serialization;
using System.Net;
using Newtonsoft.Json;
using Microsoft.AspNetCore.Http.Json;
//using static HavellsSyncApi.Common.JwtMiddleware;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddCors();
builder.Services.AddControllers();

// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(options =>
{
    options.OperationFilter<CustomHeader>();
    options.SchemaFilter<SwaggerIgnoreFilter>();
});

builder.Services.DataRegisterServices();
builder.Services.ServiceRegisterServices();
builder.Services.RegisterCommonServices();
builder.Services.AddControllers().AddJsonOptions(options => { options.JsonSerializerOptions.PropertyNamingPolicy = null; }); // Model poperty don't serialize with CamelCase. 
//builder.Services.AddScoped<IAuthorization, AuthorizationUser>();
builder.Services.AddMvcCore().AddAuthorization();

var app = builder.Build();
app.UseCors(x => x.AllowAnyOrigin().AllowAnyMethod().AllowAnyHeader());

// custom jwt auth middleware
app.UseMiddleware<JwtMiddleware>();
// Configure the HTTP request pipeline.
if (app.Environment.IsProduction() || app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}
app.UseStaticFiles(new StaticFileOptions()
{
    FileProvider = new PhysicalFileProvider(Path.Combine(Directory.GetCurrentDirectory(), "MyStaticFiles")),
    RequestPath = new PathString("/StaticFiles")
});
app.MapControllerRoute("Default", "{controller = Home}/{action = Index}/{id?}");
app.UseAuthentication();
app.UseAuthorization();
app.MapControllers();
app.Run();
