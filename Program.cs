//using Microsoft.Azure.Functions.Worker;

//var host = new HostBuilder().ConfigureFunctionsWebApplication().ConfigureServices
//    (
//        services =>
//        {
//            services.AddApplicationInsightsTelemetryWorkerService();
//            services.ConfigureFunctionsApplicationInsights();
//        }
//    ).Build();

//host.Run();

var host = new HostBuilder().ConfigureFunctionsWebApplication().Build();
host.Run();