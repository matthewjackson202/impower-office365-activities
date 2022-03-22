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
            if(DriveItemValue.ListItem == null)
            {
                throw new Exception("The specified DriveItem did not have a ListItem associated with it.");
            }
            List list = await client.GetSharepointList(token, SiteValue.Id, DriveItemValue.ListItem.ParentReference.Id);
            string[] fields = list.Columns.Select(column => column.Name).ToArray();
            return ctx =>
            {
                ctx.SetValue(List, list);
                ctx.SetValue(Fields, fields);
            };
        }
    }
}
