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

        internal string listId;
        internal string listItemId;
        internal List list;
        internal ListItem listItem;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            base.ReadContext(context);
            listId = context.GetValue(ListID);
            listItemId = context.GetValue(ListItemID);
        }

        protected override async Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token)
        {
            await base.Initialize(client, context, token);

            try
            {
                list = await client.GetSharepointList(token, site.Id, listId);
            }
            catch(Exception e)
            {
                throw new Exception("An Error Occured While Trying To Retrieve The Specified List.",e);
            }
            try
            {
                listItem = await client.GetSharepointListItem(token, site.Id, listId, listItemId);
            }
            catch(Exception e)
            {
                throw new Exception("An Error Occured While Trying To Retrieve The Specified ListItem",e);
            }

        }
    }
}
