
using Microsoft.SharePoint.Client;
using Microsoft.Identity.Client;
using Microsoft.Extensions.Configuration;


internal class Program
{
    private static IConfiguration configuration;
    private static PublicClientApplicationOptions appconfig;

    private static string sharepointUrl = "https://thekidneyfoundationofcanada.sharepoint.com/on";

    private static string[] scopes;
    private static IPublicClientApplication application;

    private static async Task Main(string[] args)
    {
        var builder = new ConfigurationBuilder()
                            .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                            .AddJsonFile("appsettings.json")
                            .AddUserSecrets<Program>();

        configuration = builder.Build();
        Console.WriteLine(builder.Sources.Count());
        appconfig = configuration.Get<PublicClientApplicationOptions>();

        scopes = new[] { "User.Read AllSites.FullControl MyFiles.Read Sites.Search.All"};


        var result = await SignInUserAndGetTokenUsingMSAL(appconfig, scopes);

        var Client = new ClientContext(sharepointUrl);

        Client.ExecutingWebRequest += async (sender, e) =>
        {

            e.WebRequestExecutor.RequestHeaders.Add("bearer", result);

        };

        var web = Client.Web;
        Client.Load(web);

        Client.ExecuteQuery();
        var site = Client.Site;

        Client.Load(site);

        var query = new ChangeQuery(true, true);

        var changes = site.GetChanges(query);

        Client.Load(changes);

        Client.ExecuteQuery();

        foreach(var change in changes)
        {
            System.Console.WriteLine("{0}, {1}", change.ChangeType, change.Time);

        }

        
    }
    private async static Task<string?> SignInUserAndGetTokenUsingMSAL(PublicClientApplicationOptions configuration, string[] scopes)
    {
        string authority = string.Concat(configuration.Instance, configuration.TenantId);

        application = PublicClientApplicationBuilder.Create(configuration.ClientId)
                                    .WithAuthority(authority)
                                    .WithDefaultRedirectUri()
                                    .Build();
        AuthenticationResult result = null;

        var accounts = await application.GetAccountsAsync();
        try
        {
            result = await application.AcquireTokenSilent(scopes, accounts.FirstOrDefault()).ExecuteAsync();
        }
        catch (MsalUiRequiredException)
        {
            result = await application.AcquireTokenInteractive(scopes).ExecuteAsync();
        }

        return result.AccessToken;
    }
}