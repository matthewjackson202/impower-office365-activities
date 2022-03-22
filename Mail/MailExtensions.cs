using Microsoft.Graph;
using System.Threading;
using System.Threading.Tasks;

namespace Impower.Office365.Mail
{
    public static class MailExtensions
    {
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
