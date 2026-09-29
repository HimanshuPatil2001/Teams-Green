using Hangfire;
using TeamsGreen.Data;
using TeamsGreen.Hangfire;
using TeamsGreen.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllers();

// Add CORS services to allow frontend to communicate with backend
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll", policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyMethod()
              .AllowAnyHeader();
    });
});

// Configure Microsoft SQL Server connection and register KeepAliveRepository
var connectionString = builder.Configuration.GetConnectionString("DefaultConnection") 
                       ?? "Server=localhost;Database=TeamsKeepAliveDb;Trusted_Connection=True;TrustServerCertificate=True;";
builder.Services.AddScoped<KeepAliveRepository>(provider => 
    new KeepAliveRepository(connectionString));

// Register MSGraphPresenceService with HttpClient factory support
builder.Services.AddHttpClient<IMSGraphPresenceService, MSGraphPresenceService>();

// Register Hangfire background job helper
builder.Services.AddScoped<KeepAliveBackgroundJob>();

// Configure Hangfire with SQL Server Storage
builder.Services.AddHangfire(configuration => configuration
    .SetDataCompatibilityLevel(CompatibilityLevel.Version_180)
    .UseSimpleAssemblyNameTypeSerializer()
    .UseRecommendedSerializerSettings()
    .UseSqlServerStorage(connectionString));

// Add the Hangfire background processing server
builder.Services.AddHangfireServer();

// Swagger and API Explorer
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(options =>
    {
        options.SwaggerDoc("v1", new Microsoft.OpenApi.Models.OpenApiInfo
        {
            Title = "Teams Status Keep-Alive API",
            Version = "v1",
            Description = "A production-grade, multi-tenant API to programmatically keep Microsoft Teams presence status active."
        });
    });

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

if (!app.Environment.IsDevelopment())
{
    app.UseHttpsRedirection();
}

app.UseCors("AllowAll");

app.UseAuthorization();

// Setup Hangfire Dashboard with custom authorization filter
app.UseHangfireDashboard("/hangfire", new DashboardOptions
{
    Authorization = new[] { new HangfireDashboardAuthorizationFilter() }
});

app.MapControllers();

app.MapGet("/", async context =>
{
    context.Response.ContentType = "text/html";
    await context.Response.WriteAsync(@"
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset='utf-8' />
            <title>Teams Stay Green - Backend Service</title>
            <style>
                body {
                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                    background-color: #0f172a;
                    color: #f8fafc;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    height: 100vh;
                    margin: 0;
                }
                .container {
                    text-align: center;
                    background: rgba(30, 41, 59, 0.7);
                    padding: 40px;
                    border-radius: 12px;
                    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
                    border: 1px solid rgba(255, 255, 255, 0.08);
                    max-width: 500px;
                }
                h1 {
                    margin-bottom: 10px;
                    font-size: 28px;
                    background: linear-gradient(135deg, #fff 0%, #cbd5e1 100%);
                    -webkit-background-clip: text;
                    -webkit-text-fill-color: transparent;
                }
                p {
                    color: #94a3b8;
                    font-size: 16px;
                    margin-bottom: 30px;
                }
                .btn {
                    display: inline-block;
                    padding: 12px 24px;
                    font-size: 16px;
                    font-weight: 600;
                    text-decoration: none;
                    border-radius: 8px;
                    color: #fff;
                    background: #6366f1;
                    transition: background 0.2s;
                }
                .btn:hover {
                    background: #4f46e5;
                }
            </style>
        </head>
        <body>
            <div class='container'>
                <h1>Teams Stay Green</h1>
                <p>The backend service is running successfully.</p>
                <a href='/hangfire' class='btn'>📊 Open Web UI Dashboard</a>
            </div>
        </body>
        </html>
    ");
});

// Schedule the Keep-Alive recurring job (Every 4 minutes)
var recurringJobManager = app.Services.GetRequiredService<IRecurringJobManager>();
recurringJobManager.AddOrUpdate<KeepAliveBackgroundJob>(
    "keep-presence-alive",
    job => job.RunAsync(),
    "*/4 * * * *"
);

app.Run();
