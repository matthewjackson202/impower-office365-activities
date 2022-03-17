using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using UiPath.MicrosoftOffice365.Activities;
using UiPath.MicrosoftOffice365.Activities.Files;
using UiPath.MicrosoftOffice365.Activities.Files.Contracts;

namespace Impower.Office365.Authentication
{
    [DisplayName("Get Graph Client From Office365 Scope")]
    public class GetClientFromScope : GraphDriveClientActivity
    {
        [RequiredArgument]
        [DisplayName("Graph Client")]
        public OutArgument<GraphServiceClient> GraphClient { get; set; }
        protected override Task<Action<AsyncCodeActivityContext>> ExecuteAsync(AsyncCodeActivityContext context, CancellationToken token)
        {
            var client = Extensions.GetClientFromScope(context);
            if(client == null) { throw new Exception("Could not acquire a GraphServiceClient from the current context."); }
            return Task.FromResult((Action<AsyncCodeActivityContext>)(ctx =>
            {
                ctx.SetValue(GraphClient, client);
            }));
        }

    }
}
