using Microsoft.Graph;
using System.Activities;
using System.ComponentModel;
using System.Threading;
using System.Threading.Tasks;

namespace Impower.Office365.Sharepoint
{
    public abstract class SharepointDriveItemActivity : SharepointDriveActivity
    {
        [Category("Input")]
        [DisplayName("DriveItem ID")]
        [RequiredArgument]
        public InArgument<string> DriveItemID { get; set; }
        internal string DriveItemIdValue;
        internal DriveItem DriveItemValue;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            DriveItemIdValue = context.GetValue(DriveItemID);
        }
        protected override async Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            await base.Initialize(client, context, token);
            DriveItemValue = await client.GetSharepointDriveItem(token, SiteId, DriveId, DriveItemIdValue);
            
        }
    }
}
