using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
namespace Impower.Office365.Sharepoint
{
    [DisplayName("Update DriveItem Fields")]
    public class UpdateDriveItemFields : SharepointDriveActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [DisplayName("DriveItem ID")]
        public InArgument<string> DriveItemID { get; set; }
        [Category("Input")]
        [RequiredArgument]
        public InArgument<Dictionary<string, object>> Fields { get; set; }
        [Category("Output")]
        [DisplayName("Updated Fields")]
        public OutArgument<Dictionary<string, object>> UpdatedFields { get; set; }
        internal Dictionary<string, object> FieldValue;
        internal string DriveItemIdValue;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            DriveItemIdValue = context.GetValue(DriveItemID);
            FieldValue = context.GetValue(Fields);
        }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            FieldValueSet fieldValueSet = new FieldValueSet
            {
                AdditionalData = FieldValue
            };
            FieldValueSet result = await client.UpdateSharepointDriveItemFields(token, SiteId, DriveId, DriveItemIdValue, fieldValueSet);
            return ctx =>
            {
                ctx.SetValue(UpdatedFields, result.AdditionalData);
            };
        }
    }
}
