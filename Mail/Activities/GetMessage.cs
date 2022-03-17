using Microsoft.Graph;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Impower.Office365;
using System.ComponentModel;

namespace Impower.Office365.Mail
{
    [DisplayName("Get Message By ID")]
    public class GetMessage : Office365Activity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("MessageID of the target message")]
        public InArgument<string> MessageID { get; set; }
        [Category("Input")]
        [Description("Email address of user the email is associated with")]
        [RequiredArgument]
        public InArgument<string> Email { get; set; }
        [Category("Output")]
        [Description("Message object retrieved")]
        public OutArgument<Message> Message { get; set; }


        internal string messageID;
        internal string email;
        protected override void ReadContext(AsyncCodeActivityContext context)
        {
            this.messageID = context.GetValue(MessageID);
            this.email = context.GetValue(Email);
        }
        protected override Task Initialize(GraphServiceClient client, AsyncCodeActivityContext context, CancellationToken token){ return Task.CompletedTask; }
        protected override async Task<Action<AsyncCodeActivityContext>> ExecuteAsyncWithClient(
          CancellationToken cancellationToken,
          GraphServiceClient client
        )
        {
            var message = await client.GetMessage(cancellationToken, messageID, email);
            return (Action<AsyncCodeActivityContext>)(ctx =>
            {
                this.Message.Set(ctx, message);
            }
            );
        }
    }
}
