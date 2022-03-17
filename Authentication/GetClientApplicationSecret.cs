using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net.Http.Headers;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Impower.Office365.Authentication
{
    [DisplayName("Get Graph Client By Application Secret")]
    public class GetClientApplicationSecret : CodeActivity
    {
        [DisplayName("Tenant ID")]
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> TenantID { get; set; }
        
        [DisplayName("Application ID")]
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> ApplicationID { get; set; }

        [DisplayName("Application Secret")]
        [Category("Input")]
        [RequiredArgument]
        public InArgument<SecureString> ApplicationSecret { get; set; }

        [DisplayName("Graph Client")]
        [Category("Output")]
        public OutArgument<GraphServiceClient> GraphClient { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            //https://stackoverflow.com/a/56154921/3892177
            var scopes = new string[] { "https://graph.microsoft.com/.default" };
            var tenantID = context.GetValue(TenantID);
            var applicationID = context.GetValue(ApplicationID);
            var applicationSecret = context.GetValue(ApplicationSecret);
            var confidentialClient = ConfidentialClientApplicationBuilder
                .Create(applicationID)
                .WithAuthority($"https://login.microsoftonline.com/{tenantID}/v2.0")
                .WithClientSecret(new System.Net.NetworkCredential(string.Empty, applicationSecret).Password)
                .Build();

            GraphServiceClient graphServiceClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        var authResult = await confidentialClient
                            .AcquireTokenForClient(scopes)
                            .ExecuteAsync();

                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                    }
                )
            );
            context.SetValue(GraphClient, graphServiceClient);
        }
    }
}
