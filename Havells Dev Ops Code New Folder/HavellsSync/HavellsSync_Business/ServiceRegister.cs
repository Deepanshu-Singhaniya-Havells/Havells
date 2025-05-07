using HavellsSync_Busines.IService;
using HavellsSync_Busines.Service;
using HavellsSync_Business.IService;
using HavellsSync_Business.Service;
using Microsoft.Extensions.DependencyInjection;
public static class ServiceRegister
{
    public static IServiceCollection ServiceRegisterServices(this IServiceCollection serviceCollection)
    {
        serviceCollection.AddScoped<IAMC, AMC>();
        serviceCollection.AddScoped<IProduct, Product>();
        serviceCollection.AddScoped<IAuthentication, Authentication>();
        serviceCollection.AddScoped<IEasyReward, EasyReward>();
        serviceCollection.AddScoped<IEpos, Epos>();
        serviceCollection.AddScoped<IWhatsApp, WhatsApp>();
        serviceCollection.AddScoped<IConsumer, Consumer>();
        serviceCollection.AddScoped<IServiceAlaCarte, ServiceAlaCarte>();
        serviceCollection.AddScoped<IGrievanceHandling, GrievanceHandling>();
        return serviceCollection;
    }
}
