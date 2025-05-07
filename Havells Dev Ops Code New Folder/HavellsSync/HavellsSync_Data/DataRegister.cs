using Microsoft.Extensions.DependencyInjection;
using HavellsSync_Data.IManager;
using HavellsSync_Data.Manager;
using HavellsSync_Data.IServiceAlaCarte;
public static class DataRegister
{
    public static IServiceCollection DataRegisterServices(this IServiceCollection serviceCollection)
    {
        serviceCollection.AddTransient<ICrmService, CrmService>();
        serviceCollection.AddScoped<IProductManager, ProductManager>();
        serviceCollection.AddScoped<IAuthenticationManager, AuthenticationManager>();
        serviceCollection.AddScoped<IAMCManager, AMCManager>();
        serviceCollection.AddScoped<IEposManager, EposManager>();
        serviceCollection.AddScoped<IEasyRewardManager, EasyRewardManager>();
        serviceCollection.AddScoped<IWhatsAppManager, WhatsAppManager>();
        serviceCollection.AddScoped<IConsumerManager, ConsumerManager>();
        serviceCollection.AddScoped<IServiceAlaCarteManager, ServiceAlaCarteManager>();
        serviceCollection.AddScoped<IGrievanceHandlingManager, GrievanceHandlingManager>();
        return serviceCollection;
    }
}
