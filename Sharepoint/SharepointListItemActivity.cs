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
    public abstract class SharepointListItemActivity : SharepointSiteActivity
    {
        [RequiredArgument]
        [DisplayName("List ID")]
        public InArgument<string> ListID { get; set; }
        [RequiredArgument]
        [DisplayName("ListItem ID")]
        public InArgument<string> ListItemID { get; set; }

        internal string ListIdValue;
        internal string ListItemIdValue;
        internal List ListValue;
        internal ListItem ListItemValue;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            ListIdValue = context.GetValue(ListID);
            ListItemIdValue = context.GetValue(ListItemID);
        }

        protected override async Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            await base.Initialize(client, context, token);

            try
            {
                ListValue = await client.GetSharepointList(token, SiteValue.Id, ListIdValue);
            }
            catch(Exception e)
            {
                throw new Exception("An Error Occured While Trying To Retrieve The Specified List.",e);
            }
            try
            {
                ListItemValue = await client.GetSharepointListItem(token, SiteValue.Id, ListIdValue, ListItemIdValue);
            }
            catch(Exception e)
            {
                throw new Exception("An Error Occured While Trying To Retrieve The Specified ListItem",e);
            }

        }
    }
}
