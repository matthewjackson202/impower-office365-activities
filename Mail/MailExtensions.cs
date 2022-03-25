using Microsoft.Graph;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Impower.Office365.Mail
{
    public static class MailExtensions
    {
        public static async Task<Stream> GetMessageAsEML(
            this GraphServiceClient client,
            CancellationToken token,
            string messageID,
            string email
        )
        {
            return await client.Users[email].Messages[messageID].Content.Request().GetAsync(token);
        }
        public static async Task<Message> GetMessage(
            this GraphServiceClient client,
            CancellationToken token,
            string messageID,
            string email,
            bool getAttachments = true
        )
        {
            var request = client.Users[email].Messages[messageID].Request();
            if (getAttachments)
            {
                request = request.Expand(message => message.Attachments);
            }
            return await request.GetAsync(token);
        }
    }
}
