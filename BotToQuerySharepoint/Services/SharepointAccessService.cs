using System.Threading.Tasks;
using BotToQuerySharepoint.Models;

namespace BotToQuerySharepoint.Services
{
    public class SharepointAccessService
    {        

        public async Task SendRequestAccessEmail(SharepointModel model, string userName)
        {
            var service = new SharepointService();
            string body =
                $"Hello,\r\nCould you please grant {model.AccessRights} access to {userName} for sharepoint site {model.Url}?\r\nThank you.";

            await service.SendEmail(model.Token, model.Owner.EmailAddress, body, "Sharepoint access request");
        }
    }
}