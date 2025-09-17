using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.Identity.Web;
using Microsoft.Graph;
using TicketRouter.Api.Services;


var builder = WebApplication.CreateBuilder(args);

// AuthN/AuthZ + OBO Graph
builder.Services
    .AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddMicrosoftIdentityWebApi(builder.Configuration.GetSection("AzureAd"))
    .EnableTokenAcquisitionToCallDownstreamApi()
    .AddMicrosoftGraph(builder.Configuration.GetSection("Graph"))
    .AddInMemoryTokenCaches();

builder.Services.AddAuthorization();
builder.Services.AddControllers();
builder.Services.AddScoped<GraphMailService>();

// CORS (allow Outlook origins + localhost)
builder.Services.AddCors(options =>
{
    options.AddPolicy("Outlook",
        policy => policy
            .WithOrigins("https://outlook.office.com", 
                         "https://outlook.office365.com",
                         "https://*.outlook.com", 
                         "https://ticketrouterapi-ecd7h3b6a6cme3g8.canadacentral-01.azurewebsites.net", 
                         "https://mahlab.net" ,
                         "https://localhost:3000")
            .AllowAnyHeader()
            .AllowAnyMethod()
            .AllowCredentials());
});

var app = builder.Build();

app.UseHttpsRedirection();
app.UseStaticFiles();        // serves wwwroot (client html/js/icons)
app.UseCors("Outlook");
app.UseAuthentication();
app.UseAuthorization();
app.MapControllers();
app.Run();