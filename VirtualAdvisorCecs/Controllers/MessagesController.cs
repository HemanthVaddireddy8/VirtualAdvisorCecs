using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Newtonsoft.Json.Linq;

namespace VirtualAdvisorCecs
{
    [BotAuthentication]
    public class MessagesController : ApiController
    {
        /// <summary>
        /// POST: api/Messages
        /// Receive a message from a user and reply to it
        /// </summary>
        public async Task<HttpResponseMessage> Post([FromBody]Activity activity)
        {
            if (activity.GetActivityType() == ActivityTypes.Message)
            {
                await Conversation.SendAsync(activity, () => new LuisIntents.LUISIntents());
            }
            else if (activity.GetActivityType() == ActivityTypes.ConversationUpdate)
            {
                {
                    IConversationUpdateActivity update = activity;
                    ConnectorClient connector = new ConnectorClient(new Uri(activity.ServiceUrl));

                    if (update.MembersAdded != null && update.MembersAdded.Any())
                    {
                        foreach (var newMember in update.MembersAdded)
                        {
                            if (newMember.Id != activity.Recipient.Id)
                            {
                                var reply3 = activity.CreateReply();
                                reply3.Text = "Hello I'm Kloppo, your academic assistant.";
                                await connector.Conversations.ReplyToActivityAsync(reply3);

                                var reply2 = activity.CreateReply();
                                reply2.Text = "I can book your appointments with the academic advisors and answer your questions.";
                                await connector.Conversations.ReplyToActivityAsync(reply2);

                                var reply = activity.CreateReply();
                                reply.Text = "What would you like me to do?";
                                await connector.Conversations.ReplyToActivityAsync(reply);

                                //await Conversation.SendAsync(activity, () => new Dialogs.RootDialog());
                            }
                        }
                    }
                }
            }
            else
            {
                HandleSystemMessage(activity);
            }
            var response = Request.CreateResponse(HttpStatusCode.OK);
            return response;
        }

        private Activity HandleSystemMessage(Activity message)
        {
            string messageType = message.GetActivityType();
            if (messageType == ActivityTypes.DeleteUserData)
            {
                // Implement user deletion here
                // If we handle user deletion, return a real message
            }
            else if (messageType == ActivityTypes.ConversationUpdate)
            {
                // Handle conversation state changes, like members being added and removed
                // Use Activity.MembersAdded and Activity.MembersRemoved and Activity.Action for info
                // Not available in all channels
            }
            else if (messageType == ActivityTypes.ContactRelationUpdate)
            {
                // Handle add/remove from contact lists
                // Activity.From + Activity.Action represent what happened
            }
            else if (messageType == ActivityTypes.Typing)
            {
                // Handle knowing that the user is typing
            }
            else if (messageType == ActivityTypes.Ping)
            {
            }

            return null;
        }
    }
}