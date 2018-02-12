using System;
using System.Configuration;
using System.Threading.Tasks;
using AuthBot;
using AuthBot.Models;
using BotToQuerySharepoint.Forms;
using BotToQuerySharepoint.Models;
using BotToQuerySharepoint.Services;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.FormFlow;
using Microsoft.Bot.Builder.Luis;
using Microsoft.Bot.Builder.Luis.Models;

namespace BotToQuerySharepoint.Dialogs
{
    [Serializable]
    [LuisModel("Key-with-dash", "secret")]
    public class SpDialog : LuisDialog<object>
    {
        [LuisIntent("sp-access")]
        public virtual async Task ProcessSpAccessRequest(IDialogContext context, LuisResult luisResult)
        {
            var accessToken =
                await context.GetAccessToken(
                    ConfigurationManager.AppSettings[
                        "ActiveDirectory.ResourceId"]); //assuming graph api is the resource
            SpFormHandler formHandler = new SpFormHandler();
            IFormDialog<SharepointModel> formDialog = formHandler.GetFormDialog(luisResult.Entities, accessToken);

            context.Call(formDialog, OnFormComplete);
        }

        private async Task OnFormComplete(IDialogContext context, IAwaitable<SharepointModel> awaitableResult)
        {
            SharepointModel spModel = await awaitableResult;

            var service = new SharepointService();
            spModel.Owner = await service.GetOwnerNameForSite(spModel.Token, spModel.SiteId);

            context.PrivateConversationData.SetValue("spModel", spModel);
            PromptDialog.Confirm(
                context, 
                SpAction, 
                $"The owner of the site is {spModel.Owner.Name}. Do you want me to send an email to the owner or create a ticket?", 
                options: new []{"Email", "Ticket"},
                patterns: new[]{new[] { "Email"}, new[] { "Ticket" }}
            );
        }

        private async Task SpAction(IDialogContext context, IAwaitable<bool> awaitableResult)
        {
            bool answer = await awaitableResult;
            context.PrivateConversationData.TryGetValue("spModel", out SharepointModel spModel);
            context.PrivateConversationData.RemoveValue("spModel");
            AuthResult authResult;
            context.UserData.TryGetValue(ContextConstants.AuthResultKey, out authResult);
            var sharepointAccessService = new SharepointAccessService();
            if (answer)
            {
                await sharepointAccessService.SendRequestAccessEmail(spModel, authResult.UserName);
                await context.PostAsync($"I have sent an email to {spModel.Owner.Name} requesting access for you.");
            }
            else
            {
                //here was the creation of a ticket automatically
                await context.PostAsync("I have created a ticket requesting access for you.");
            }

            context.Done(true);
        }
    }
}
