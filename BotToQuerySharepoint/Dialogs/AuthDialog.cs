using System;
using System.Configuration;
using System.Threading;
using System.Threading.Tasks;
using AuthBot;
using AuthBot.Dialogs;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;

namespace BotToQuerySharepoint.Dialogs
{
    [Serializable]
    public class AuthDialog : IDialog<object>
    {
        public async Task StartAsync(IDialogContext context)
        {
            if (context.Activity.ChannelId != "email")
                await context.PostAsync("Oh, it appears you are not logged in");
            context.Wait(MessageReceivedAsync);
        }

        public virtual async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> item)
        {
            var message = await item;

            //endpoint v1
            var appSetting = ConfigurationManager.AppSettings["ActiveDirectory.ResourceId"];
            if (string.IsNullOrEmpty(await context.GetAccessToken(appSetting)))
            {
                var azureAuthDialog = new AzureAuthDialog(appSetting);
                await context.Forward(azureAuthDialog, this.ResumeAfterAuth, message, CancellationToken.None);
            }
            else
            {
                await context.PostAsync("A problem occured and you cannot log in");
                context.Done(true);
            }
        }

        private async Task ResumeAfterAuth(IDialogContext context, IAwaitable<string> result)
        {
            var message = await result;

            await context.PostAsync(message);
            
            context.Done(true);
        }
    }
}
