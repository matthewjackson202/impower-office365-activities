using Impower.Office365.Sharepoint.Models;
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
        [Category("Input")]
        [DisplayName("Link Type")]
        public InArgument<LinkType> LinkType { get; set; }
        internal LinkType LinkTypeValue;
        [Category("Output")]
        [DisplayName("Sharing Link")]
        public OutArgument<string> SharingLink { get; set; }
        [Category("Output")]
        [DisplayName("Drive Item")]
        public OutArgument<DriveItem> DriveItem { get; set; }
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            LinkTypeValue = context.GetValue(LinkType);
        }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken cancellationToken, GraphServiceClient client)
        {
            Permission permission = await client.ShareDriveItem(cancellationToken, DriveItemIdValue, SiteId, DriveId, LinkTypeValue);
            return ctx =>
            {
                ctx.SetValue(DriveItem, DriveItemValue);
                ctx.SetValue(SharingLink, permission.Link.WebUrl);
            };

        }
    }
}
