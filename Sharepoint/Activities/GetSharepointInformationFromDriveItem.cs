using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Impower.Office365.Sharepoint
{
    [DisplayName("Get Sharepoint Information From Drive Item")]
    public class GetSharepointInformationFromDriveItem : Office365Activity
    {
        [RequiredArgument]
        [Category("Input")]
        public InArgument<DriveItem> DriveItem { get; set; }
        [Category("Output")]
        public OutArgument<Drive> Drive { get; set; }
        [Category("Output")]
        public OutArgument<Site> Site { get; set; }
        [Category("Output")]
        public OutArgument<string> DriveName { get; set; }
        [Category("Output")]
        public OutArgument<string> SiteURL { get; set; }
        private DriveItem DriveItemValue { get; set; }
        private Site SiteValue;
        private Drive DriveValue;
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            SiteValue = await client.AttemptToRetreiveSiteFromDriveItem(token, DriveItemValue);
            DriveValue = await client.AttemptToRetrieveDriveFromDriveItem(token, DriveItemValue, SiteValue.Id);

            return ctx =>
            {
                ctx.SetValue(Drive, DriveValue);
                ctx.SetValue(Site, SiteValue);
                ctx.SetValue(DriveName, DriveValue.Name);
                ctx.SetValue(SiteURL, SiteValue.WebUrl);
            };
            

        }
        protected override Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            return Task.CompletedTask;
        }
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            DriveItemValue = DriveItem.Get(context);
        }
    }
}
