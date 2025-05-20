using Microsoft.Azure.Functions.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using System;

[assembly: FunctionsStartup(typeof(Brief_Builder.Startup))]

namespace Brief_Builder
{
    public class Startup : FunctionsStartup
    {
        public override void ConfigureAppConfiguration(IFunctionsConfigurationBuilder builder)
        {
            builder.ConfigurationBuilder.AddEnvironmentVariables().Build();
        }

        public override void Configure(IFunctionsHostBuilder builder)
        {
            string URL = Environment.GetEnvironmentVariable("Dataverse_URL");
            string clientID = Environment.GetEnvironmentVariable("ClientID");
            string secret = Environment.GetEnvironmentVariable("Secret");

            string AuthType = "ClientSecret";
            string conn = $"Url = {URL};AuthType = {AuthType};SkipDiscovery=true;Secret={secret};ClientId={clientID};";

            builder.Services.AddSingleton<IOrganizationService>(s =>
            {
                return new ServiceClient(conn);
            });
        }
    }
}