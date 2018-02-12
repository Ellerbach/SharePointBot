using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using Newtonsoft.Json;

namespace BotToQuerySharepoint.Dialogs
{
    [Serializable]
    public class KbDialog : IDialog<object>
    {
        public async Task StartAsync(IDialogContext context)
        {
            //await context.PostAsync("What do you want to find out?");
            context.Wait(MessageReceivedAsync);
        }

        private async Task WaitForMessage(IDialogContext context, IAwaitable<IMessageActivity> item)
        {
            context.Wait(MessageReceivedAsync);
        }

        public virtual async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> item)
        {
            var message = await item;

            //if (message.Text == "no")
            //{
            //    context.Done(true);
            //}
            //else if (message.Text == "yes")
            //{
            //    await context.PostAsync("What do you want to find out?");
            //    context.Wait(MessageReceivedAsync);
            //}
            //else
            {
                string kbId = ConfigurationManager.AppSettings["QnaMaker.KbId"];
                string kbKey = ConfigurationManager.AppSettings["QnaMaker.KbKey"];
                string qnaUrl =
                    $"https://westus.api.cognitive.microsoft.com/qnamaker/v2.0/knowledgebases/{kbId}/generateAnswer";

                HttpClient client = new HttpClient();
                string strtosend = message.Text;
                if (message.ChannelId == "email")
                {
                    var str = strtosend.Split('\n');
                    int maxidx = str.Length;
                    if (maxidx > 3)
                        maxidx = 4;
                    for (int i = 0; i < maxidx; i++)
                        strtosend += str[i] + " ";
                }

                var json = new
                {
                    question = strtosend,
                    top = 3
                };
                var content = new StringContent(JsonConvert.SerializeObject(json), Encoding.UTF8, "application/json");
                content.Headers.Add("Ocp-Apim-Subscription-Key", kbKey);
                HttpResponseMessage response = await client.PostAsync(qnaUrl, content);
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    QnaResponse qnaResponse =
                        JsonConvert.DeserializeObject<QnaResponse>(await response.Content.ReadAsStringAsync());

                    if (qnaResponse.answers.Count == 0)
                    {
                        await context.PostAsync(
                            "I couldn't find any information on the topic. \nAnything else from the QNA section I can help you with ?");
                        context.Done(true);
                    }
                    else if ((qnaResponse.answers.Count == 1) || (message.ChannelId == "email"))
                    {
                        await context.PostAsync(qnaResponse.answers.First().answer);
                        context.Done(true);
                    }
                    else
                    {
                        qnaResponse.answers.Add(new Answer()
                        {
                            answer = "(exit)",
                            questions = new string[] {"None of the above."}
                        });

                        PromptDialog.Choice<Answer>(context, OnSelectedAnswer,
                            qnaResponse.answers,
                            "Is this what you were looking for ?",
                            descriptions: qnaResponse.answers.Select(x => x.questions.First()),
                            promptStyle: PromptStyle.Auto
                        );
                    }
                }
                else
                {
                    await context.PostAsync(
                        "Something went terribly wrong. You will have to figure this out by human interraction.");
                    context.Done(true);
                }
            }
        }

        private async Task OnSelectedAnswer(IDialogContext context, IAwaitable<Answer> result)
        {
            Answer answer = await result;
            if (answer.answer!="(exit)")
            {
                await context.PostAsync(answer.answer);
            }
            else
            {
                await context.PostAsync("Sorry I couldn't find what you were looking for.");
            }

            context.Done(true);
        }

        [Serializable]
        class QnaResponse
        {
            public List<Answer> answers { get; set; }
        }

        [Serializable]
        class Answer
        {
            public string answer { get; set; }
            public string[] questions { get; set; }
            public string score { get; set; }

            public override string ToString()
            {
                return questions.First();
            }
        }
    }
}
