using Havells.D365.Services.Abstract;
using Havells.D365.Services.Concrete;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace Havells.D365.API.Installer
{
    public class MvcInstaller : IInstaller
    {
        public void InstallServices(IServiceCollection services, IConfiguration configuration)
        {
            services.AddControllersWithViews();
            services.AddScoped<IWorkOrderRepository, WorkOrdersRepository>();
            services.AddScoped<IIncidentRepository, IncidentRepository>();
            services.AddScoped<IWorkOrderProductRepository, WorkOrderProductRepository>();
            services.AddScoped<IWOServiceRepository, WOServiceRepository>();
        }
    }
}
