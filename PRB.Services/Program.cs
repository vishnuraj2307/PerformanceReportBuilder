using Microsoft.EntityFrameworkCore;
using PRB.Repository;
using PRB.Repository.Automation_Repository;
using PRB.Repository.DataContext;
using PRB.Repository.Repository;
using Serilog;
using System.Configuration;

var builder = WebApplication.CreateBuilder(args);

var configuration = new ConfigurationBuilder()
 .SetBasePath(Directory.GetCurrentDirectory())
 .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
 .AddEnvironmentVariables()
 .Build();
builder.Services.AddCors(corsoptions =>
{
    corsoptions.AddPolicy("Mypolicy",
    (policyoptions) =>
    {
        policyoptions.AllowAnyHeader().AllowAnyOrigin().AllowAnyMethod();
    });
});

builder.Services.AddCors(option =>
{
    option.AddPolicy(name: "AllowOrgin", builder =>
    {
        builder.WithOrigins("https://localhost:4200").AllowAnyHeader().AllowAnyMethod();
    });
}
    );

//Read Configuration from appSettings    
var config = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
//Initialize Logger    
Log.Logger = new LoggerConfiguration().MinimumLevel.Verbose().WriteTo.File("D:\\PRB.Services\\PRB.Services\\Logs\\PRB_Logs.log", rollingInterval:RollingInterval.Day).CreateLogger();





// Add services to the container.
builder.Services.AddDbContext<PRB_DB_Context>(Options =>
 Options.UseSqlServer(builder.Configuration.GetConnectionString("MyDBConnection")));

//var appSettings = builder.Configuration.Get<PRB.Services.AppSettings>();




builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddScoped<IPRB_Repository, PRBRepository>();
builder.Services.AddScoped<IAutomation_Repository, Automation_Repository>();
builder.Services.AddScoped<IRuleExecutor, RuleExecutor>();


//builder.Services.AddScoped<IAutomation_Repository, Automation_Repository>();


var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();
app.UseCors("Mypolicy");
app.UseAuthorization();


app.MapControllers();

try
{
    Log.Information("Application Running...");
    app.Run();
}
catch(Exception ex)
{
    Log.Fatal("The Application Failed to start correctly. ",ex);
}
finally
{
    Log.CloseAndFlush();
}
