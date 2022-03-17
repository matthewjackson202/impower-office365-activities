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
            string email
        )
        {
            return await client.Users[email].Messages[messageID].Request().GetAsync(token);
        }
    }
}
