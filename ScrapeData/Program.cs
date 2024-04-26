// See https://aka.ms/new-console-template for more information
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using ScrapeData;

Console.WriteLine("Hello, World!");

var hostBuilder = Host.CreateDefaultBuilder(args)
            .ConfigureAppConfiguration(opt =>
            {
                opt.AddUserSecrets<Program>();
            })
            .ConfigureServices((hostContext, services) =>
            {
                services.AddSingleton<Configuration>(opt =>
                {
                    return new Configuration
                    {
                        WebsiteUrl = hostContext.Configuration.GetValue<string>("WebsiteUrl"),
                        Term = hostContext.Configuration.GetValue<string>("FormatString"),
                        ScrapeValue = hostContext.Configuration.GetValue<string>("ScrapeValue"),
                        Directory = hostContext.Configuration.GetValue<string>("Directory"),
                        EmailPassword = hostContext.Configuration.GetValue<string>("EmailPassword"),
                        EmailAddress = hostContext.Configuration.GetValue<string>("EmailAddress"),
                    };
                });

                services.AddScoped<IScrapeWorker, ScrapeWorker>();
            });

var host = hostBuilder.Build();

Console.WriteLine("Start Worker");
Console.WriteLine("--------------------");
var worker = host.Services.GetRequiredService<IScrapeWorker>();
worker.Run();
Console.WriteLine("--------------------");
Console.WriteLine("Worker Completed");
