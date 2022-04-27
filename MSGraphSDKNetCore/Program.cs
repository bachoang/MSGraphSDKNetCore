using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.Graph;

namespace MSGraphSDKNetCore
{
    class Program
    {
        static void Main(string[] args)
        {
            var scopes = new[] { "https://graph.microsoft.com/.default" };
            var tenantId = "[put in your Directory (tenant) ID from the Overview blade of App Registration]";
            // Value from app registration
            var clientId = "[put in your Application (client) ID from the Overview blade of App Registration]";
            var clientSecret = "[put in your secret value (not secret ID) from certificates & secret blade]";
          
            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // https://docs.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            // client credentials grant flow
            var clientSecretCredential = new ClientSecretCredential(
                tenantId, clientId, clientSecret, options);

            // For auth code grant flow use the following instead of
            // using Azure.Identity;
            /*
            var options = new InteractiveBrowserCredentialOptions
            {
                TenantId = tenantId,
                ClientId = clientId,
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                // MUST be http://localhost or http://localhost:PORT
                // See https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki/System-Browser-on-.Net-Core
                RedirectUri = new Uri("http://localhost"),
            };

            // https://docs.microsoft.com/dotnet/api/azure.identity.interactivebrowsercredential
            var interactiveCredential = new InteractiveBrowserCredential(options);
            */

            // auth code grant flow
            // var graphClient = new GraphServiceClient(interactiveCredential, scopes);

            // client credentials flow:
            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);
            CreateGroup(graphClient).Wait();
        }
        static async Task CreateGroup(GraphServiceClient graphClient)
        {
            try
            {

                var group = new Group
                {
                    Description = "xxx",
                    GroupTypes = new List<String>()
                    {
                        "Unified"
                    },
                    MailEnabled = true,
                    MailNickname = "libraryyyy",
                    SecurityEnabled = false,
                    // Intead of setting DisplayName via AdditionalData I can also set it via the following DisplayName property
                    // DisplayName = "TestGroup",
                    // if ResourceBehaviorOptions property is available, use group.ResourceBehaviorOptions instead of
                    AdditionalData = new Dictionary<string, object> { { "resourceBehaviorOptions", new List<string>() { "WelcomeEmailDisabled" } }, { "DisplayName", "TestGroup"  } }

                };

                Console.WriteLine("creating group...");
                
                var createdGroup = await graphClient.Groups
                    .Request()
                    .AddAsync(group);
                

                // Let's read our created group back
                Console.WriteLine("reading group...");
                 var mygroup = await graphClient.Groups[createdGroup.Id]
                    .Request()
                    .GetAsync();

                List<string> behavioroptions = new List<string>();
                Object obj;

                // try to query the value at the ResourceBehaviorOptions property first.  If the property does not existk
                // then we will look for it at the AdditionalData property
                if (mygroup.GetType().GetProperty("ResourceBehaviorOptions") != null)
                { 
                    behavioroptions = (List<string>) mygroup.GetType().GetProperty("ResourceBehaviorOptions").GetValue(mygroup,null);
                }
                else if (mygroup.AdditionalData.ContainsKey("resourceBehaviorOptions"))
                { 
                    mygroup.AdditionalData.TryGetValue("resourceBehaviorOptions", out obj);
                    Console.WriteLine($"behaviorOptions: {obj.ToString()}");
                }
                else
                { 
                    Console.WriteLine("Can't query group's ResourceBehaviorOptions...");
                }
                // Console.WriteLine($"behaviorOptions: {behavioroptions[0]}");

                Console.WriteLine("deleteing group...");
                
                await graphClient.Groups[createdGroup.Id]
                    .Request()
                    .DeleteAsync();           

            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception occurred");
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
