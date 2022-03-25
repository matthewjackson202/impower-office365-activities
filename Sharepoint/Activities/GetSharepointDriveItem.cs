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
        [Category("Output")]
        public OutArgument<ItemReference> Reference { get; set; }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(
          CancellationToken token,
          GraphServiceClient client
        )
        {
            return ctx =>
            {
                Fields.Set(ctx, new Dictionary<string, object>());
                DriveItem.Set(ctx, DriveItemValue);
                if(DriveItemValue.ParentReference != null)
                {
                    Reference.Set(ctx, DriveItemValue.ParentReference);
                }
                if (DriveItemValue.ListItem != null)
                {
                    ListItem.Set(ctx, DriveItemValue.ListItem);
                    if (DriveItemValue.ListItem.AdditionalData != null)
                    {
                        Fields.Set(ctx, DriveItemValue.ListItem.Fields.AdditionalData);
                    }
                }
            };
        }
    }
}
