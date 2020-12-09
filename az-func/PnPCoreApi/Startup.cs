using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using PnPCoreApi;

[assembly: FunctionsStartup(typeof(Startup))]

namespace PnPCoreApi
{
    public class Startup : FunctionsStartup
    {
        public override void Configure(IFunctionsHostBuilder builder)
        {
            var config = builder.GetContext().Configuration;
            var appInfo = new AppInfo();
            config.Bind(appInfo);

            builder.Services.AddSingleton(appInfo);
            builder.Services.AddPnPCore(opts => { 
                // optional configuration here
            });
        }
    }
}