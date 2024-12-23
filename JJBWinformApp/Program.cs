using Common;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace JJBWinformApp
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();

            ServiceCollection services = new ServiceCollection();
            ConfigureServices(services);

            var serviceProvider = services.BuildServiceProvider();

            var formMain = serviceProvider.GetRequiredService<MainForm>();

            Application.Run(formMain);
        }

        private static void ConfigureServices(ServiceCollection services)
        {
            //services.AddSingleton<DbFactory>();
            services.AddSingleton<IDialogService, MessageBoxDialogService>();
            services.AddSingleton<TableService>();
            services.AddScoped<MainForm>();
            services.AddScoped<TableCompareForm>();

            //register configuration
            IConfigurationBuilder cfgBuilder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json")
                .AddJsonFile($"appsettings.{Environment.GetEnvironmentVariable("DOTNET_ENVIRONMENT")}.json", optional: true, reloadOnChange: false);
            IConfiguration configuration = cfgBuilder.Build();
            services.AddSingleton<IConfiguration>(configuration);

            services.Configure<RowIndexOptions>(configuration.GetSection("RowIndexOptions"));
        }
    }
}