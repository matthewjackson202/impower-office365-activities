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
        [Category("Connection")]
        [DisplayName("Sharepoint URL")]
        public InArgument<string> WebURL { get; set; }
        internal string SiteId => SiteValue.Id;
        internal string WebUrlValue;
        internal Site SiteValue;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            WebUrlValue = context.GetValue(WebURL);
        }
        protected override async Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            try
            {
                SiteValue = await client.GetSharepointSiteFromUrl(token, WebUrlValue);
            }
            catch(Exception e)
            {
                throw new Exception("Error Occured While Retrieving Site From URL", e);
            }
        }
    }
}
