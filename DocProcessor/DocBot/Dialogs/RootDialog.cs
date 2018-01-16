using System;
using System.Threading.Tasks;
using DocBot.Dtos;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Newtonsoft.Json.Linq;

namespace DocBot.Dialogs
{
    [Serializable]
    public class RootDialog : IDialog<object>
    {
        public Task StartAsync(IDialogContext context)
        {
            context.Wait(MessageReceivedAsync);

            return Task.CompletedTask;
        }

        private async Task MessageReceivedAsync(IDialogContext context, IAwaitable<object> result)
        {
            var activity = await result as Activity;

            // Typing
            var typingActivity = ((Activity)context.Activity).CreateReply();
            typingActivity.Type = ActivityTypes.Typing;
            await context.PostAsync(typingActivity);

            string searchText = activity.Text;
            QnAResponse response = await Utils.ServiceProxies.GetQnAResponse(searchText);

            if (response.answers != null && response.answers.Length > 0 && response.answers[0].score >= 50)
            {
                searchText = response.answers[0].answer;
            }

            // Now do Azure Search
            string searchResult = Utils.ServiceProxies.SearchMarkdown(searchText);
            string resultText = "I unfortunately didn't find anything for you!";
            string resultSpeech = "I unfortunately didn't find anything for you!";

            if (!String.IsNullOrEmpty(searchResult))
            {
                JObject jsonResult = JObject.Parse(searchResult);
                if (jsonResult["value"][0] != null)
                {
                    resultText = jsonResult["value"][0]["content"].Value<string>();
                    resultSpeech = jsonResult["value"][0]["textcontent"].Value<string>();
                }
            }

            Activity reply = activity.CreateReply(resultText);
            reply.Speak = resultSpeech;

            await context.PostAsync(reply);
            context.Wait(MessageReceivedAsync);
        }
    }
}