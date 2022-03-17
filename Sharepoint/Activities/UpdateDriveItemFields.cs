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
    public class UpdateDriveItemFields : SharepointDriveItemActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<Dictionary<string, object>> Fields { get; set; }
        [Category("Output")]
        [DisplayName("Updated Fields")]
        public OutArgument<Dictionary<string, object>> UpdatedFields { get; set; }
        internal Dictionary<string, object> fields;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            fields = context.GetValue(Fields);
        }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            var fieldValueSet = new FieldValueSet
            {
                AdditionalData = fields
            };
            var result = await client.UpdateSharepointDriveItemFields(token, site.Id, drive.Id, driveItemId, fieldValueSet);
            return (Action<AsyncCodeActivityContext>)(ctx =>
            {
                ctx.SetValue(UpdatedFields, result.AdditionalData);
            });
        }
    }
}
