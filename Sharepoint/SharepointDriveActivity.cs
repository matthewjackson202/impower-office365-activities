using Microsoft.Graph;
using System;
using System.Activities;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
namespace Impower.Office365.Sharepoint
{
    public abstract class SharepointDriveActivity : SharepointSiteActivity
    {
        [RequiredArgument]
        [Category("Connection")]
        [DisplayName("Sharepoint Drive")]
        public InArgument<string> DriveName { get; set; }
        internal string driveName;
        internal Drive drive;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            driveName = context.GetValue(DriveName);
        }
        protected override async Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            await base.Initialize(client, context, token);
            try
            {
                this.drive = await client.GetSharepointDrive(token, this.site.Id, driveName);
            }
            catch(Exception e)
            {
                throw new Exception("Error Occured While Retrieving Drive By Name", e);
            }
        }
    }
}
