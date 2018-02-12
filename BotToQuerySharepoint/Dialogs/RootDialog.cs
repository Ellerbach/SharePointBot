using System;
using System.Configuration;
using System.Threading;
using System.Threading.Tasks;
using AuthBot;
using AuthBot.Models;
using BotToQuerySharepoint.Services;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;

namespace BotToQuerySharepoint.Dialogs
{
    [Serializable]
    public class RootDialog : IDialog<object>
    {
        public async Task StartAsync(IDialogContext context)
        {
            context.Wait(MessageReceivedAsync);
        }
        
        public virtual async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> item)
        {
            var message = await item;

            if (!await IsUserLoggedIn(context))
            {
                await context.Forward(new AuthDialog(), ResumeAfterAuth, message, CancellationToken.None);
            }
            else if (message.Text.ToLower().Contains("token"))
            {
                await TokenSample(context);
            }
            else if (message.Text.ToLower().Contains("logout"))
            {
                await context.Logout();
                context.Wait(this.MessageReceivedAsync);
            }
            else if (message.ChannelId != "email")
            {
                if (message.Text.ToLower().Contains("sp") || message.Text.ToLower().Contains("sharepoint"))
                {
                    //if (message.ChannelId != "email")
                    //{
                    //    await context.Forward(new SpDialog(), Resume, message, CancellationToken.None);
                    //}
                    //else
                        await context.Forward(new SpDialog(), Resume, message, CancellationToken.None);
                }
                else if (message.Text.ToLower().Contains("hi"))
                {
                    await context.PostAsync(
                        "Hi there, what would you like to do today? (You can ask for access to a Sharepoint site or ask me any Endava related question and I will try to answer it)");
                }
                else
                {
                    await context.Forward(new KbDialog(), Resume, message, CancellationToken.None);
                }
            }
            else
            {
                // Here was the creation of a ticket directly in the service ticket database

                await context.PostAsync("A ticket has been succesfully created for you. Thank you.");
            }
        }

        private async Task ResumeAfterAuth(IDialogContext context, IAwaitable<object> result)
        {
            context.Wait(MessageReceivedAsync);
            if (context.Activity.ChannelId != "email")
            {
                await context.PostAsync($"Type logout if you want to log out");
                await context.PostAsync("Hi there, what would you like to do today? (ask a question or type sharepoint for Sharepoint issues)");
            }
            else
            {
                await context.PostAsync($"Hi there,\nType logout if you want to log out in a blank email without any signature.\nPlease send me a mail with your question: (type: sp for Sharepoint and question for QNA)");
            }
        }

        private async Task Resume(IDialogContext context, IAwaitable<object> result)
        {
            await context.PostAsync("What else would you like to do today? (ask a question or type sharepoint for Sharepoint issues)");
            context.Wait(MessageReceivedAsync);
        }

        private async Task<bool> IsUserLoggedIn(IDialogContext context)
        {
            var accessToken = await context.GetAccessToken(ConfigurationManager.AppSettings["ActiveDirectory.ResourceId"]);
            return !string.IsNullOrEmpty(accessToken);
        }

        public async Task TokenSample(IDialogContext context)
        {
            //endpoint v1
            var accessToken = await context.GetAccessToken(ConfigurationManager.AppSettings["ActiveDirectory.ResourceId"]);

            if (string.IsNullOrEmpty(accessToken))
            {
                return;
            }

            await context.PostAsync($"Your access token is: {accessToken}");

            context.Wait(MessageReceivedAsync);
        }
    }
}
