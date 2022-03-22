using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Impower.Office365.Sharepoint {
    [DisplayName("Get Sharing Link For DriveItem")]
    public class GetSharingLinkForSharepointDriveItem : SharepointDriveItemActivity
    {
        [Category("Output")]
        [DisplayName("Sharing Link")]
        public OutArgument<string> SharingLink { get; set; }
        [Category("Output")]
        [DisplayName("Drive Item")]
        public OutArgument<DriveItem> DriveItem { get; set; }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken cancellationToken, GraphServiceClient client)
        {
            Permission permission = await client.ShareDriveItem(cancellationToken, DriveItemIdValue, SiteId, DriveId);
            return ctx =>
            {
                ctx.SetValue(DriveItem, DriveItemValue);
                ctx.SetValue(SharingLink, permission.Link.WebUrl);
            };

        }
    }
}
