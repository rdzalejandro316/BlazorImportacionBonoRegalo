using importacionBono.Components;
using importacionBono.Components.Services;
using Syncfusion.Blazor;

Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("MzQ4ODU2NUAzMjM3MmUzMDJlMzBYV3lRQkRBOEFDaEI1WlZNOHAzRWJCNk1lUkpVNytMcERGM0V4QkVDSklnPQ==");

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents().AddCircuitOptions(options => {  options.DetailedErrors = true; });

builder.Services.AddSyncfusionBlazor();
builder.Services.AddSingleton<ExcelService>();
builder.Services.AddSingleton<DataService>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error", createScopeForErrors: true);
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();

app.UseStaticFiles();
app.UseAntiforgery();

app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
