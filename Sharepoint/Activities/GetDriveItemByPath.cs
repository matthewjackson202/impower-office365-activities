using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
namespace Impower.Office365.Sharepoint
{
    [DisplayName("Get DriveItem By Path")]
    public class GetDriveItemByPath : SharepointDriveActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> Path { get; set; }

        [Category("Output")]
        [DisplayName("Results")]
        public OutArgument<List<DriveItem>> DriveItems { get; set; }
        [Category("Output")]
        [DisplayName("First")]
        public OutArgument<DriveItem> DriveItem { get; set; }

        internal string path;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            path = context.GetValue(Path);
        }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(
          CancellationToken token,
          GraphServiceClient client
        )
        {
            var items = await client.GetSharepointDriveItemsByPath(token, this.site.Id, this.drive.Id, path);
            return (Action<AsyncCodeActivityContext>)(ctx =>
            {
                this.DriveItems.Set(ctx, items);
                if (items.Any())
                {
                    this.DriveItem.Set(ctx, items.First());
                }
            });

        }
    }
}
