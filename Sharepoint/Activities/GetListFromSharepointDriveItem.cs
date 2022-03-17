using Microsoft.Graph;
using System;
using System.Activities;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
namespace Impower.Office365.Sharepoint
{
    [DisplayName("Get List From DriveItem")]
    public class GetListFromSharepointDriveItem : SharepointDriveItemActivity
    {
        [Category("Output")]
        public OutArgument<List> List { get; set; }
        [Category("Output")]
        public OutArgument<string[]> Fields { get; set; }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            var driveItem = await client.GetSharepointDriveItem(token, site.Id, drive.Id, driveItemId);
            if(driveItem.ListItem == null)
            {
                //Can this even happen?
                throw new Exception("The specified DriveItem did not have a ListItem associated with it.");
            }
            var list = await client.GetSharepointList(token, site.Id, driveItem.ListItem.ParentReference.Id);
            var fields = list.Columns.Select(column => column.Name).ToArray();
            return (Action<AsyncCodeActivityContext>)(ctx =>
            {
                ctx.SetValue(List, list);
                ctx.SetValue(Fields, fields);
            });
        }
    }
}
