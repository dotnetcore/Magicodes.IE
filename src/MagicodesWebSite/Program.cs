using Magicodes.ExporterAndImporter.Builder;
using Magicodes.ExporterAndImporter.Filters;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using System;
using MagicodesWebSite.Extensions;
using Microsoft.OpenApi.Models;

namespace MagicodesWebSite
{
    public class Program
    {
        public const string MiddlewareScenario = "middlewarescenario";
        public const string FilterScenario = "filterscenario";
        public static void Main(string[] args)
        {
            CreateHostBuilder(args).Build().Run();
        }

        public static IWebHostBuilder CreateHostBuilder(string[] args)
        {
            string scenario;
            if (args.Length == 0)
            {
                Console.WriteLine("Choose a sample to run:");
                Console.WriteLine($"1. MiddlewareScenario");
                Console.WriteLine($"2. FilterScenario");
                Console.WriteLine();

                scenario = Console.ReadLine();
            }
            else
            {
                scenario = args[0];
            }

            Type startupType;
            switch (scenario)
            {
                case "1":
                case MiddlewareScenario:
                    startupType = typeof(StartupMiddlewareTest);
                    break;

                case "2":
                case FilterScenario:
                    startupType = typeof(StartupFilterTest);
                    break;
                default:
                    throw new InvalidOperationException();
            }
            return new WebHostBuilder()
                .UseKestrel()
                .UseIISIntegration()
                .ConfigureLogging(b =>
                {
                    b.AddConsole();
                    b.SetMinimumLevel(LogLevel.Critical);
                })
                .UseContentRoot(Environment.CurrentDirectory)
                .UseStartup(startupType);
        }

        class StartupMiddlewareTest
        {
            public void ConfigureServices(IServiceCollection services)
            {
                services.AddControllers();
                services.AddSwaggerGen(c =>
                {
                    c.SwaggerDoc("v1", new OpenApiInfo { Title = "Magicodes WebSite API", Version = "v1" });
                    c.OperationFilter<AddRequiredHeaderParameter>();
                });
            }
            public void Configure(IApplicationBuilder app)
            {
                app.UseRouting();
                app.UseMagiCodesIE();

                app.UseSwagger();
                app.UseSwaggerUI();
                app.UseEndpoints(endpoints =>
                {
                    endpoints.MapControllers();
                });
            }
        }

        class StartupFilterTest
        {
            public void ConfigureServices(IServiceCollection services)
            {
                services.AddControllers(options => options.Filters.Add(typeof(MagicodesFilter)));
                services.AddRazorPages();
                services.AddSwaggerGen(c =>
                {
                    c.SwaggerDoc("v1", new OpenApiInfo { Title = "Magicodes WebSite API", Version = "v1" });
                    c.OperationFilter<AddRequiredHeaderParameter>();
                });
            }
            public void Configure(IApplicationBuilder app)
            {
                app.UseRouting();
                app.UseSwagger();
                app.UseSwaggerUI();
                app.UseEndpoints(endpoints =>
                {
                    endpoints.MapRazorPages();
                    endpoints.MapControllers();
                });
            }
        }

    }
}
