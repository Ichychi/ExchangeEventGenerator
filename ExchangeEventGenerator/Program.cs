namespace ExchangeEventGenerator;

internal static class Program {
    public static void Main(string[] args) {
        //service worker template with dependency injection
        var host = Host.CreateDefaultBuilder()
            .ConfigureServices((context, services) => {
                services.AddSingleton(new EventGenerator(context.Configuration));
                services.AddHostedService<Worker>();
            }).Build();

        host.Run();
    }
}

