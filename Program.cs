using Microsoft.AspNetCore.Authentication.OpenIdConnect;
//using Microsoft.AspNetCore.Authorization;
//using Microsoft.AspNetCore.Mvc.Authorization;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.UI;
//using System.IdentityModel.Tokens.Jwt;
//using Microsoft.Graph;
using Microsoft.Identity.Web.TokenCacheProviders.InMemory;
/*using Microsoft.Net.Http.Headers;
using Microsoft.Kiota.Abstractions;
using Microsoft.Graph.Authentication;
using Microsoft.Kiota.Abstractions.Authentication;*/

var builder = WebApplication.CreateBuilder(new WebApplicationOptions {
    ContentRootPath = Directory.GetCurrentDirectory(),
    WebRootPath = "static"
});

// Sign-in users with the Microsoft identity platform
builder.Services.AddAuthentication(OpenIdConnectDefaults.AuthenticationScheme).AddMicrosoftIdentityWebApp(builder.Configuration.GetSection("AzureAD")).EnableTokenAcquisitionToCallDownstreamApi().AddMicrosoftGraph(builder.Configuration.GetSection("Graph")).AddInMemoryTokenCaches();

builder.Services.AddAuthorization();

builder.Services.AddControllersWithViews().AddMicrosoftIdentityUI();

var app = builder.Build();

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthentication();

app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();