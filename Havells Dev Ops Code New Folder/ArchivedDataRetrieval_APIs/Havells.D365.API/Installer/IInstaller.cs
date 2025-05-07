using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace Havells.D365.API.Installer
{
    public interface IInstaller
    {
        void InstallServices(IServiceCollection serviceCollection,IConfiguration configuration);
    }
}
