using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Linq;

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
        [Category("Input")]
        [RequiredArgument]
        [DisplayName("Use Display Names")]
        [Description("Allows referencing columns by their display name. If set, keys will be matching first against the internal name and then against the display name, as a fallback.")]
        public InArgument<bool> UseDisplayNames { get; set; }
        [Category("Output")]
        [DisplayName("Updated Fields")]
        public OutArgument<Dictionary<string, object>> UpdatedFields { get; set; }
        private Dictionary<string, object> FieldsValue;
        private string DriveItemIdValue;
        private bool UseDisplayNamesValue; 
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            DriveItemIdValue = context.GetValue(DriveItemID);
            FieldsValue = context.GetValue(Fields);
            UseDisplayNamesValue = context.GetValue(UseDisplayNames);
        }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(CancellationToken token, GraphServiceClient client)
        {
            if (UseDisplayNamesValue)
            {
                var driveItem = await client.GetSharepointDriveItem(token, SiteId, DriveId, DriveItemIdValue);
                var list = await client.GetSharepointList(token, SiteId, driveItem.ListItem.ParentReference.Id);
                //TODO - this could be cleaned up.
                //This will throw if one of the display names resolves to a name that already exists in the dictionary.
                var NewFieldsValue = new Dictionary<string, object>();
                var WritableColumns = list.Columns.Where(column => column.ReadOnly ?? false);
                foreach (var kvp in FieldsValue)
                {
                    var MatchingColumns = WritableColumns.Where(column => column.Name.Equals(kvp.Key));
                    if(MatchingColumns.Any())
                    {
                        NewFieldsValue.Add(kvp.Key, kvp.Value);
                        break;
                    }
                    MatchingColumns = WritableColumns.Where(column => column.DisplayName.Equals(kvp.Key));
                    if(MatchingColumns.Any())
                    {
                        var MatchingColumn = MatchingColumns.First();
                        NewFieldsValue.Add(MatchingColumn.Name, kvp.Value);
                        break;
                    }
                    throw new Exception($"Could not find a field matching '{kvp.Key}' in the target list. Available fields are: {String.Join(",", WritableColumns.Select(column => column.Name))}");
                }
                FieldsValue = NewFieldsValue;
            }

            FieldValueSet fieldValueSet = new FieldValueSet
            {
                AdditionalData = FieldsValue
            };
            FieldValueSet result = await client.UpdateSharepointDriveItemFields(token, SiteId, DriveId, DriveItemIdValue, fieldValueSet);
            return ctx =>
            {
                ctx.SetValue(UpdatedFields, result.AdditionalData);
            };
        }
    }
}
