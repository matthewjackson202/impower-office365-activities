using Microsoft.Graph;
using System;
using System.Activities;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;
namespace Impower.Office365.Sharepoint
{
    public abstract class SharepointSiteActivity : Office365Activity
    {
        internal Site site;
        [RequiredArgument]
        [Category("Connection")]
        [DisplayName("Sharepoint URL")]
        public InArgument<string> WebURL { get; set; }
        internal string webUrl;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            webUrl = context.GetValue(WebURL);
            
        }
        protected override async Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            try
            {
                this.site = await client.GetSharepointSite(token, webUrl);
            }
            catch(Exception e)
            {
                throw new Exception("Error Occured While Retrieving Site From URL", e);
            }
        }
    }
}
