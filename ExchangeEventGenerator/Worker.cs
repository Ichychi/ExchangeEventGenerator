namespace ExchangeEventGenerator;

//TODO rename this class and create new background services for other functionality (e.g. deleting events, sending mail, ...)
public class Worker : BackgroundService {
    private readonly ILogger<Worker> _logger;
    private readonly EventGenerator _generator;
    private readonly int _interval;

    public Worker(ILogger<Worker> logger, EventGenerator generator, IConfiguration configuration) {
        _logger = logger;
        _generator = generator;
        _interval = configuration.GetSection("Settings").GetValue<int>("WorkerIntervalInMinutes");
    }

    //TODO make use of logger throughout the program
    protected override async Task ExecuteAsync(CancellationToken stoppingToken) {
        while (!stoppingToken.IsCancellationRequested){
            //_logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
            _generator.Generate();
            await Task.Delay(1000*60*_interval, stoppingToken);
        }
    }
}