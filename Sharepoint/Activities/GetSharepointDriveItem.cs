using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
namespace Impower.Office365.Sharepoint
{
    [DisplayName("Get DriveItem")]
    public class GetSharepointDriveItem : SharepointDriveItemActivity
    {
        [Category("Output")]
        public OutArgument<ListItem> ListItem { get; set; }
        [Category("Output")]
        public OutArgument<DriveItem> DriveItem { get; set; }

        [Category("Output")]
        public OutArgument<Dictionary<string,object>> Fields { get; set; }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(
          CancellationToken token,
          GraphServiceClient client
        )
        {
            var driveItem = await client.GetSharepointDriveItem(
                token,
                this.site.Id,
                this.drive.Id,
                driveItemId
            );
            return (Action<AsyncCodeActivityContext>)(ctx =>
            {
                Fields.Set(ctx, new Dictionary<string, object>());
                DriveItem.Set(ctx, driveItem);
                if (driveItem.ListItem != null)
                {
                    ListItem.Set(ctx, driveItem.ListItem);
                    if (driveItem.ListItem.AdditionalData != null)
                    {
                        Fields.Set(ctx, driveItem.ListItem.Fields.AdditionalData);
                    }
                }
            });
        }
    }
}
