using HavellsSync_ModelData.ICommon;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace HavellsSync_ModelData.Common
{
    public static class RegisterCommon
    {
        public static IServiceCollection RegisterCommonServices(this IServiceCollection serviceCollection)
        {
            serviceCollection.AddScoped<IAES256, AES256>();
            serviceCollection.AddScoped<ICustomLog, CustomLog>();
            return serviceCollection;
        }
    }
}
