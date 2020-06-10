using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Calendar.v3;
using Google.Apis.Services;

using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Luis;
using Microsoft.Bot.Builder.Luis.Models;
using Microsoft.Bot.Connector;

using AdaptiveCards;
using VirtualAdvisorCecs.Appointments;

namespace VirtualAdvisorCecs.LuisIntents
{
    [LuisModel("1f6c92be-a1f0-426f-a51a-a21b00de3dab", "2d040fb5cf3d402da36444d7c97f4ed6")]

    [Serializable]
    public class LUISIntents : LuisDialog<object>
    {
        private static string ApplicationName = ConfigurationManager.AppSettings["ApplicationName"].ToString();
        private static string ClientId = ConfigurationManager.AppSettings["ClientId"].ToString();
        private static string ClientSecret = ConfigurationManager.AppSettings["ClientSecret"].ToString();
        private static string RedirectURL = ConfigurationManager.AppSettings["RedirectURL"].ToString();
        private static string connString = ConfigurationManager.ConnectionStrings["ApplicationConnectionString"].ConnectionString;

        private static ClientSecrets GoogleClientSecrets = new ClientSecrets()
        {
            ClientId = ClientId,
            ClientSecret = ClientSecret
        };

        private static string UserName = string.Empty;
        private static string StudentEmailID = string.Empty;
        private static string AdvisorEmailID = string.Empty;
        private static string Program = string.Empty;
        private static DateTime appointmentDateTime = new DateTime();

        //public LUISIntents(Activity activity)
        //{
        //    userName = activity.From.Name;
        //    msgReceivedDate = activity.Timestamp.ToString();// ?? DateTime.Now;
        //}
        [LuisIntent("")]
        [LuisIntent("None")]
        public async Task None(IDialogContext context, LuisResult result)
        {
            var activity = context.Activity as IMessageActivity;
            string message = String.Empty;
            if (activity.Text != null && activity.Value == null)
            {
                message = "Sorry, I did not understand your question. Could you be more specific, to help us understand and assist you in a better way.";
            }
            else if (activity.Text == null & activity.Value != null)
            {

            }
            await context.PostAsync(message);

            context.Wait(this.MessageReceived);
        }
        [LuisIntent("Greeting.Welcome")]
        public async Task WelcomeGreetings(IDialogContext context, LuisResult result)
        {
            string replyMessage = GetWelcomeMessage();
            await context.PostAsync(replyMessage);

            context.Wait(this.MessageReceived);
        }
        [LuisIntent("Greeting.Farewell")]
        public async Task FarewellGreetings(IDialogContext context, LuisResult result)
        {
            string replyMessage = "I'm glad I could help." + " \U0001F642";//GetWelcomeMessage();
            await context.PostAsync(replyMessage);

            context.Wait(this.MessageReceived);
        }
        [LuisIntent("AdvisoryOfficeLocation")]
        public async Task AdvisorOfficeLocation(IDialogContext context, LuisResult result)
        {
            string reply = string.Empty;
            var listReplies = new List<string>();
            listReplies = GetAdvisingOfficeAddress().Split(',').ToList();
            foreach (var strReply in listReplies)
            {
                reply = strReply;
                await context.PostAsync(reply);
                Thread.Sleep(5000);
            }
            context.Wait(this.MessageReceived);
        }
        [LuisIntent("AdmissionChange")]
        public async Task AdmissionChange(IDialogContext context, LuisResult result)
        {
            string reply = string.Empty;
            var listReplies = new List<string>();
            listReplies = GetAdmissionChangeInfo().Split('#').ToList();
            foreach (var strReply in listReplies)
            {
                reply = strReply;
                await context.PostAsync(reply);
                Thread.Sleep(5000);
            }

            //PromptDialog.Confirm(context, ShareAdmissionChangeForm, "Do you want us to share the Admission change form?");
            Thread.Sleep(5000);
            PromptDialog.Text(
                context: context,
                resume: ShareAdmissionChangeForm,
                prompt: "Do you want us to share the Admission change form?",
                retry: "Sorry, could you please try again."
                );
        }
        public async Task ShareAdmissionChangeForm(IDialogContext context, IAwaitable<string> confirmCode)
        {
            var replyMessage = context.MakeMessage();
            //var isRequired = await confirmCode;
            var confirmCodes = new List<string>() { "yes", "yup", "yeah", "ya", "ok", "okay", "k", "sure", "alright", "right", "ofcourse" };
            var isRequired = false;
            isRequired = confirmCodes.Any(confirmCode.ToString().ToLower().Contains);

            if (isRequired)
            {
                replyMessage.Text = "Check the attached Admission Change Form for more information.";
                List<Microsoft.Bot.Connector.Attachment> attachments = new List<Microsoft.Bot.Connector.Attachment>();
                Microsoft.Bot.Connector.Attachment attachment = new Microsoft.Bot.Connector.Attachment
                {
                    ContentType = "application/pdf",
                    ContentUrl = HttpContext.Current.Server.MapPath("~/Documents/Admission_Change_Form.pdf"),
                    Name = "Admission Change Form"
                };
                attachments.Add(attachment);
                replyMessage.Attachments = attachments;
            }
            else
            {
                replyMessage.Text = "Okay. Can I assist you with anything else?";
            }
            await context.PostAsync(replyMessage);
            context.Wait(this.MessageReceived);
        }
        #region Co-Op
        [LuisIntent("Co-Op.Overview")]
        public async Task CoOpOverview(IDialogContext context, LuisResult result)
        {
            string reply = string.Empty;
            var listReplies = new List<string>();
            listReplies = GetCoOpOverview().Split('#').ToList();
            foreach (var strReply in listReplies)
            {
                reply = strReply;
                await context.PostAsync(reply);
                Thread.Sleep(5000);
            }
            Thread.Sleep(5000);
            PromptDialog.Text(
                context: context,
                resume: ShareCOOPReferenceDoc,
                prompt: "Do you want us to share the CO-OP simplicity reference document?",
                retry: "Sorry, could you please try again."
                );
        }
        public async Task ShareCOOPReferenceDoc(IDialogContext context, IAwaitable<string> confirmCode)
        {
            var replyMessage = context.MakeMessage();
            //var isRequired = await confirmCode;
            var confirmCodes = new List<string>() { "yes", "yup", "yeah", "ya", "ok", "okay", "k", "sure", "alright", "right", "ofcourse" };
            var isRequired = false;
            isRequired = confirmCodes.Any(confirmCode.ToString().ToLower().Contains);

            if (isRequired)
            {
                replyMessage.Text = "Check the attached CO-OP Simplicity Reference Document for more information.";
                List<Microsoft.Bot.Connector.Attachment> attachments = new List<Microsoft.Bot.Connector.Attachment>();
                Microsoft.Bot.Connector.Attachment attachment = new Microsoft.Bot.Connector.Attachment
                {
                    ContentType = "application/pdf",
                    ContentUrl = HttpContext.Current.Server.MapPath("~/Documents/CoopRefDoc.pdf"),
                    Name = "CO-OP Simplicity Reference Document"
                };
                attachments.Add(attachment);
                replyMessage.Attachments = attachments;
            }
            else
            {
                replyMessage.Text = "Okay. Can I assist you with anything else?";
            }
            await context.PostAsync(replyMessage);

            context.Wait(this.MessageReceived);
        }
        [LuisIntent("Co-Op.Office")]
        public async Task CoOpOfficeDetails(IDialogContext context, LuisResult result)
        {
            var dt = GetCoOpOffice();
            var address = string.Empty;
            var phoneNumber = string.Empty;
            var EmailID = string.Empty;
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dRow in dt.Rows)
                {
                    address = dRow["AddressLine"].ToString();
                    phoneNumber = dRow["PhoneNumber"].ToString();
                    EmailID = dRow["EmailID"].ToString();
                }
            }

            var reply = context.MakeMessage();
            reply.Text = "The office of CO-OP education is at " + address;
            await context.PostAsync(reply);

            Thread.Sleep(5000);

            reply.Text = "You can also reach us at " + phoneNumber + " or drop us an email at " + EmailID;
            await context.PostAsync(reply);

            context.Wait(this.MessageReceived);
        }

        [LuisIntent("Co-Op.RegistrationOptions")]
        public async Task GetCoopRegistrationOptions(IDialogContext context, LuisResult result)
        {
            var reply = "There are four options for the students looking to register for a Co-Op -";
            await context.PostAsync(reply);
            Thread.Sleep(3000);

            PromptDialog.Choice(
                    context: context,
                    options: new List<string> {"Degree co-op", "Additive co-op",
                        "Self Report co-op", "International / F1 Students co-op"},
                    promptStyle: PromptStyle.PerLine,
                    resume: GetCoopRegistrationOptionsPartTwo,
                    prompt: "So, which option from below are you looking at, to register for Co-op program?",
                    retry: "Sorry, could you please try again."
                    );
        }
        public async Task GetCoopRegistrationOptionsPartTwo(IDialogContext context, IAwaitable<string> strRegOption)
        {
            var RegOption = await strRegOption;
            var degreeCodes = new List<string>() { "degree", "d", "degree coop credit", "first", "1", "top", "degree co-op credit" };
            var additiveCodes = new List<string>() { "additive", "a", "additive coop credit", "second", "2", "sec", "additive co-op credit", "add" };
            var selfReport = new List<string>() { "self report", "self", "self report your internship", "third", "3", "three" };
            var F1Students = new List<string>() { "international", "f1", "fourth", "last", "4", "int", "f1 students" };

            if (degreeCodes.Any(RegOption.ToString().ToLower().Equals))
            {
                var degreeOption = "Would you like to know in-detail about any of the below options of a Degree co-op program?";
                await context.PostAsync(degreeOption);
                Thread.Sleep(3000);

                PromptDialog.Choice(
                    context: context,
                    options: new List<string> {"1. Overview", "2. Benefits",
                        "3. Student Eligibility"},
                    promptStyle: PromptStyle.PerLine,
                    resume: showDegreeRegOption,
                    prompt: "Please respond with either option number or the option name -",
                    retry: "Sorry, could you please try again."
                    );
            }
            else if (additiveCodes.Any(RegOption.ToString().ToLower().Equals))
            {
                var additiveOption = "Would you like to know in-detail about any of the below options of a Additive co-op program?";
                await context.PostAsync(additiveOption);
                Thread.Sleep(3000);

                PromptDialog.Choice(
                    context: context,
                    options: new List<string> {"1. Overview", "2. Benefits",
                        "3. Student Eligibility"},
                    promptStyle: PromptStyle.PerLine,
                    resume: showAdditiveRegOption,
                    prompt: "Please respond with either option number or the option name -",
                    retry: "Sorry, could you please try again."
                    );
            }
            else if (selfReport.Any(RegOption.ToString().ToLower().Equals))
            {
                var SelfReportOption = "Would you like to know in-detail about any of the below options of a Self Report co-op program?";
                await context.PostAsync(SelfReportOption);
                Thread.Sleep(3000);

                PromptDialog.Choice(
                    context: context,
                    options: new List<string> {"1. Overview", "2. Benefits",
                        "3. Student Eligibility"},
                    promptStyle: PromptStyle.PerLine,
                    resume: showSelfReportRegOption,
                    prompt: "Please respond with either option number or the option name -",
                    retry: "Sorry, could you please try again."
                    );
            }
            else if (F1Students.Any(RegOption.ToString().ToLower().Equals))
            {
                var F1StudentsOption = "Would you like to know in-detail about any of the below options of a International / F1 Students co-op program?";
                await context.PostAsync(F1StudentsOption);
                Thread.Sleep(3000);

                PromptDialog.Choice(
                    context: context,
                    options: new List<string> {"1. Overview", "2. Benefits",
                        "3. Student Eligibility"},
                    promptStyle: PromptStyle.PerLine,
                    resume: showF1StudentsRegOption,
                    prompt: "Please respond with either option number or the option name -",
                    retry: "Sorry, could you please try again."
                    );
            }
        }

        public async Task showDegreeRegOption(IDialogContext context, IAwaitable<string> strOption)
        {
            var option = await strOption;
            if (option.ToLower().Equals("1") || option.ToLower().Equals("overview")
                || option.ToLower().Equals("first") || option.ToLower().Equals("top"))
            {
                var reply = GetCoOpOverview("Degree");
                await context.PostAsync(reply);
                context.Wait(this.MessageReceived);
            }
            else if (option.ToLower().Equals("2") || option.ToLower().Equals("benefits")
              || option.ToLower().Equals("second") || option.ToLower().Equals("two"))
            {
                var reply = GetCoOpBenefits("Degree");
                await context.PostAsync(reply);
                context.Wait(this.MessageReceived);
            }
            else if (option.ToLower().Equals("3") || option.ToLower().Equals("student eligibility")
              || option.ToLower().Equals("third") || option.ToLower().Equals("eligibility") || option.ToLower().Equals("student"))
            {
                var reply = GetStudentEligibility("Degree");
                context.Wait(this.MessageReceived);
            }
        }
        public async Task showAdditiveRegOption(IDialogContext context, IAwaitable<string> strOption)
        {
            var option = await strOption;
            if (option.ToLower().Equals("1") || option.ToLower().Equals("overview")
                || option.ToLower().Equals("first") || option.ToLower().Equals("top"))
            {
                var reply = GetCoOpOverview("Additive");
                await context.PostAsync(reply);
                context.Wait(this.MessageReceived);
            }
            else if (option.ToLower().Equals("2") || option.ToLower().Equals("benefits")
              || option.ToLower().Equals("second") || option.ToLower().Equals("two"))
            {
                var reply = GetCoOpBenefits("Additive");
                await context.PostAsync(reply);
                context.Wait(this.MessageReceived);
            }
            else if (option.ToLower().Equals("3") || option.ToLower().Equals("student eligibility")
              || option.ToLower().Equals("third") || option.ToLower().Equals("eligibility") || option.ToLower().Equals("student"))
            {
                var reply = GetStudentEligibility("Additive");
                context.Wait(this.MessageReceived);
            }
        }

        public async Task showSelfReportRegOption(IDialogContext context, IAwaitable<string> strOption)
        {
            var option = await strOption;
            if (option.ToLower().Equals("1") || option.ToLower().Equals("overview")
                || option.ToLower().Equals("first") || option.ToLower().Equals("top"))
            {
                var reply = GetCoOpOverview("SelfReport");
                await context.PostAsync(reply);
                context.Wait(this.MessageReceived);
            }
            else if (option.ToLower().Equals("2") || option.ToLower().Equals("benefits")
              || option.ToLower().Equals("second") || option.ToLower().Equals("two"))
            {
                var reply = GetCoOpBenefits("SelfReport");
                await context.PostAsync(reply);
                context.Wait(this.MessageReceived);
            }
            else if (option.ToLower().Equals("3") || option.ToLower().Equals("student eligibility")
              || option.ToLower().Equals("third") || option.ToLower().Equals("eligibility") || option.ToLower().Equals("student"))
            {
                var reply = GetStudentEligibility("SelfReport");
                context.Wait(this.MessageReceived);
            }
        }

        public async Task showF1StudentsRegOption(IDialogContext context, IAwaitable<string> strOption)
        {
            var option = await strOption;
            if (option.ToLower().Equals("1") || option.ToLower().Equals("overview")
                || option.ToLower().Equals("first") || option.ToLower().Equals("top"))
            {
                var reply = GetCoOpOverview("F1Students");
                await context.PostAsync(reply);
                context.Wait(this.MessageReceived);
            }
            else if (option.ToLower().Equals("2") || option.ToLower().Equals("benefits")
              || option.ToLower().Equals("second") || option.ToLower().Equals("two"))
            {
                var reply = GetCoOpBenefits("F1Students");
                await context.PostAsync(reply);
                context.Wait(this.MessageReceived);
            }
            else if (option.ToLower().Equals("3") || option.ToLower().Equals("student eligibility")
              || option.ToLower().Equals("third") || option.ToLower().Equals("eligibility") || option.ToLower().Equals("student"))
            {
                var reply = GetStudentEligibility("F1Students");
                context.Wait(this.MessageReceived);
            }
        }

        [LuisIntent("Co-Op.StudentEligibility")]
        public async Task GetCoopStudentEligibility(IDialogContext context, LuisResult result)
        {
            string reply1 = "Sure, I can provide you the student eligibility criteria for Co-op program.";
            await context.PostAsync(reply1);
            Thread.Sleep(5000);

            PromptDialog.Choice(
                    context: context,
                    options: new List<string> {"Degree Co-op Credit", "Additive Co-op Credit",
                        "Self Report your Internship", "International F1 Student Registration for Co-op"},
                    promptStyle: PromptStyle.PerLine,
                    resume: ShowCoopStudentCriteria,
                    prompt: "So, which option from below are you looking at, to register for Co-op program?",
                    retry: "Sorry, could you please try again."
                    );
        }
        private List<string> GetCoopRegistrationOptions()
        {
            var listRegOptions = new List<string>();
            listRegOptions.Add("Degree Co-op Credit");
            listRegOptions.Add("Additive Co-op Credit");
            listRegOptions.Add("Self Report your Internship");
            listRegOptions.Add("International Registration for Co-op");
            return listRegOptions;
        }
        public async Task ShowCoopStudentCriteria(IDialogContext context, IAwaitable<string> strRegOption)
        {
            var strSelectedOption = await strRegOption;
            var degreeCodes = new List<string>() { "degree", "d", "degree coop credit", "first", "1", "top", "degree co-op credit" };
            var additiveCodes = new List<string>() { "additive", "a", "additive coop credit", "second", "2", "sec", "additive co-op credit", "add" };
            var selfReport = new List<string>() { "self report", "self", "self report your internship", "third", "3", "three" };
            var F1Students = new List<string>() { "international", "f1", "fourth", "last", "4", "int", "f1 students" };

            var overview = string.Empty;
            var Requirements = string.Empty;
            if (degreeCodes.Any(strSelectedOption.ToString().ToLower().Equals))
            {
                var dt = GetStudentEligibility("Degree");
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dRow in dt.Rows)
                    {
                        overview = dRow["Overview"].ToString();
                        Requirements = dRow["Requirements"].ToString();
                    }
                }
                await context.PostAsync("Thank you. For a Degree Co-op credit program -");
                Thread.Sleep(3000);

                var listOverview = new List<string>();
                var listRequirements = new List<string>();
                listOverview = overview.Split('#').ToList();
                listRequirements = Requirements.Split('#').ToList();

                foreach (var strOverview in listOverview)
                {
                    await context.PostAsync(strOverview);
                    Thread.Sleep(3000);
                }
                foreach (var strRequirement in listRequirements)
                {
                    await context.PostAsync(strRequirement);
                    Thread.Sleep(3000);
                }
            }
            else if (additiveCodes.Any(strSelectedOption.ToString().ToLower().Equals))
            {
                var dt = GetStudentEligibility("Additive");
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dRow in dt.Rows)
                    {
                        overview = dRow["Overview"].ToString();
                        Requirements = dRow["Requirements"].ToString();
                    }
                }
                await context.PostAsync("Thank you. For a Additive Co-op credit program -");
                Thread.Sleep(3000);

                var listOverview = new List<string>();
                var listRequirements = new List<string>();
                listOverview = overview.Split('#').ToList();
                listRequirements = Requirements.Split('#').ToList();

                foreach (var strOverview in listOverview)
                {
                    await context.PostAsync(strOverview);
                    Thread.Sleep(3000);
                }
                foreach (var strRequirement in listRequirements)
                {
                    await context.PostAsync(strRequirement);
                    Thread.Sleep(3000);
                }
            }
            else if (selfReport.Any(strSelectedOption.ToString().ToLower().Equals))
            {
                var dt = GetStudentEligibility("SelfReport");
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dRow in dt.Rows)
                    {
                        overview = dRow["Overview"].ToString();
                        Requirements = dRow["Requirements"].ToString();
                    }
                }
                await context.PostAsync("Thank you. If you wish to Self Report your internship program -");
                Thread.Sleep(3000);

                var listOverview = new List<string>();
                var listRequirements = new List<string>();
                listOverview = overview.Split('#').ToList();
                listRequirements = Requirements.Split('#').ToList();

                foreach (var strOverview in listOverview)
                {
                    await context.PostAsync(strOverview);
                    Thread.Sleep(3000);
                }
                foreach (var strRequirement in listRequirements)
                {
                    await context.PostAsync(strRequirement);
                    Thread.Sleep(3000);
                }

                var reply = @"[Check the Self Report Internship/Co-op Verification form here](https://docs.google.com/forms/d/e/1FAIpQLSe_zyMYR3S-ToVg1hSFWqV8WXo7lgAz1C17DnB3PvVVOft8eQ/viewform)";
                await context.PostAsync(reply);
                Thread.Sleep(3000);
            }
            else if (F1Students.Any(strSelectedOption.ToString().ToLower().Equals))
            {
                var dt = GetStudentEligibility("F1Students");
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dRow in dt.Rows)
                    {
                        overview = dRow["Overview"].ToString();
                        Requirements = dRow["Requirements"].ToString();
                    }
                }
                await context.PostAsync("Thank you. For an international student to enroll for a Co-op program -");
                Thread.Sleep(3000);

                var listOverview = new List<string>();
                var listRequirements = new List<string>();
                listOverview = overview.Split('#').ToList();
                listRequirements = Requirements.Split('#').ToList();

                foreach (var strOverview in listOverview)
                {
                    await context.PostAsync(strOverview);
                    Thread.Sleep(3000);
                }
                foreach (var strRequirement in listRequirements)
                {
                    await context.PostAsync(strRequirement);
                    Thread.Sleep(3000);
                }
            }
            context.Wait(this.MessageReceived);
        }

        [LuisIntent("Co-Op.Apply")]
        public async Task CoopRegistration(IDialogContext context, LuisResult result)
        {
            var reply1 = "I can help you by taking you through various steps of applying for a co-op";
            await context.PostAsync(reply1);
            Thread.Sleep(5000);

            var reply2 = "So, the first step is to login into Simplicity with your Unique Name and Password";
            await context.PostAsync(reply2);
            Thread.Sleep(5000);

            var reply3 = "You can login into Simplicity at - " + @"[UM Dearborn - Simplicity](https://umichdearborn-csn.simplicity.com/)";
            await context.PostAsync(reply3);
            Thread.Sleep(5000);

            var reply4 = "After you login, update your profile, upload your resume and all the other relevant documents";
            await context.PostAsync(reply4);
            Thread.Sleep(5000);

            var reply5 = "After you build a strong profile, search for jobs and we suggest you to apply for as many as you can to increase your chances.";
            await context.PostAsync(reply5);
            Thread.Sleep(5000);

            var reply6 = "Once you are hired, make sure to report us about your employement.";
            await context.PostAsync(reply6);
            Thread.Sleep(5000);

            var reply7 = "If you need any assistance or like to know about studeb=nt services, you could always stop by at our office";
            await context.PostAsync(reply7);
            context.Wait(this.MessageReceived);
        }

        [LuisIntent("Co-Op.CreditHours")]
        public async Task GetCreditHours(IDialogContext context, LuisResult result)
        {
            var reply = "Sure, I could provide you the credit count for a Co-Op program.";
            await context.PostAsync(reply);
            Thread.Sleep(3000);

            PromptDialog.Choice(
                    context: context,
                    options: new List<string> {"1. Degree", "2. Additive",
                        "3. Self Report", "4. International / F1 Students"},
                    promptStyle: PromptStyle.PerLine,
                    resume: ShowCreditHours,
                    prompt: "Which registration option's credit hour count do you need?",
                    retry: "Sorry, could you please try again."
                    );
        }
        public async Task ShowCreditHours(IDialogContext context, IAwaitable<string> strRegOption)
        {
            var strSelectedOption = await strRegOption;
            var degreeCodes = new List<string>() { "1. Degree", "degree", "d", "degree coop credit", "first", "1", "top", "degree co-op credit" };
            var additiveCodes = new List<string>() { "2. Additive", "additive", "a", "additive coop credit", "second", "2", "sec", "additive co-op credit", "add" };
            var selfReport = new List<string>() { "3. Self Report", "self report", "self", "self report your internship", "third", "3", "three" };
            var F1Students = new List<string>() { "4. International / F1 Students", "international", "f1", "fourth", "last", "4", "int", "f1 students" };

            if (degreeCodes.Any(strSelectedOption.ToString().ToLower().Equals))
            {
                var strDegreeCreditHours = "For a Degree program, Co-Op carries 1 credit hour.";
                await context.PostAsync(strDegreeCreditHours);
                Thread.Sleep(2000);

                var strDegreeCourse = "Course is ENGR - 399. It will be considered as full-time.";
                await context.PostAsync(strDegreeCourse);
                Thread.Sleep(2000);

                var strDegreeApproval = "However, your require your department's approval prior to registration.";
                await context.PostAsync(strDegreeApproval);
            }
            else if (additiveCodes.Any(strSelectedOption.ToString().ToLower().Equals))
            {
                var strDegreeCreditHours = "For a Additive program, Co-Op carries 1 credit hour.";
                await context.PostAsync(strDegreeCreditHours);
                Thread.Sleep(2000);

                var strDegreeCourse = "Courses - CIS, ECE, IMSE & ME: 299, 399, 499 (One credit for each section). It will be considered as full-time.";
                await context.PostAsync(strDegreeCourse);
                Thread.Sleep(2000);

                var strDegreeApproval = "However, your require your department's approval prior to registration.";
                await context.PostAsync(strDegreeApproval);
            }
            else if (selfReport.Any(strSelectedOption.ToString().ToLower().Equals))
            {
                var strDegreeCreditHours = "For a Self Report program, Co-Op carries 1 credit hour.";
                await context.PostAsync(strDegreeCreditHours);
                Thread.Sleep(2000);

                var strDegreeCourse = "Course is ENGR - 399. It will be considered as full-time.";
                await context.PostAsync(strDegreeCourse);
                Thread.Sleep(2000);

                var strDegreeApproval = "However, your require your department's approval prior to registration.";
                await context.PostAsync(strDegreeApproval);
            }
            else if (F1Students.Any(strSelectedOption.ToString().ToLower().Equals))
            {
                var strDegreeCreditHours = "For a International/F1 Students program, Co-Op carries 1 credit hour.";
                await context.PostAsync(strDegreeCreditHours);
                Thread.Sleep(2000);

                var strDegreeCourse = "Course is ENGR - 399. It will be considered as full-time.";
                await context.PostAsync(strDegreeCourse);
                Thread.Sleep(2000);

                var strDegreeApproval = "However, your require your department's approval prior to registration.";
                await context.PostAsync(strDegreeApproval);
            }
        }

        [LuisIntent("Co-Op.DeptApproval")]
        public async Task GetDeptApprovalForCoop(IDialogContext context, LuisResult result)
        {
            var reply = "Yes, you would need your concerned department's approval before you register for any Co-Op program.";
            await context.PostAsync(reply);
            context.Wait(this.MessageReceived);
        }

        #endregion


        #region Co-Op Student Services

        [LuisIntent("Co-Op.Student.Assistance")]
        public async Task GetCoOpStudentServices(IDialogContext context, LuisResult result)
        {
            var reply1 = context.MakeMessage();
            reply1.Text = "Yes, we provide various services to assist you with your Co-Op and Job Search";
            await context.PostAsync(reply1);
            Thread.Sleep(2000);

            var reply2 = context.MakeMessage();
            reply2.Text = "We provide services such as -";
            await context.PostAsync(reply2);
            Thread.Sleep(5000);

            string services = GetServices();
            var reply3 = context.MakeMessage();
            reply3.Text = services;
            await context.PostAsync(reply3);
            Thread.Sleep(10000);

            var reply4 = context.MakeMessage();
            reply4.Text = "Stop into our office or contact us for assistance and/or take advantage of these available resources.";
            await context.PostAsync(reply4);

            context.Wait(this.MessageReceived);
        }
        [LuisIntent("Co-Op.JobLinks")]
        public async Task GetJobLinks(IDialogContext context, LuisResult result)
        {
            var reply1 = context.MakeMessage();
            reply1.Text = "Sure, I could provide you some job links.";
            await context.PostAsync(reply1);
            Thread.Sleep(3000);

            var reply2 = context.MakeMessage();
            reply2.Text = "May I know the stream you are looking at for a job, from below options?";
            await context.PostAsync(reply2);
            Thread.Sleep(3000);

            PromptDialog.Choice(
                    context: context,
                    options: new List<string> {"1. Automation Alley", "2. Mechanical Engineering",
                        "3. Computer Science", "4. International", "5. Electrical & Computer Engineering"},
                    promptStyle: PromptStyle.PerLine,
                    resume: ShowJobLinks,
                    prompt: "Please respond with the option number.",
                    retry: "Sorry, could you please try again."
                    );
        }
        public async Task ShowJobLinks(IDialogContext context, IAwaitable<string> strRegOption)
        {
            var streamOption = await strRegOption;
            var dt = new DataTable();

            var reply = "Thank you. Let me find some job links that might be helpful for you in your job search.";
            await context.PostAsync(reply);
            Thread.Sleep(5000);

            if (streamOption.Equals("1. Automation Alley"))
            {
                dt = GetJobLinks("Automation Alley");
                foreach (DataRow dRow in dt.Rows)
                {
                    var linkName = dRow["linkName"].ToString();
                    var linkURL = dRow["linkURL"].ToString();

                    var replyMessage = string.Format(@"[{0}]({1})", linkName, linkURL);
                    await context.PostAsync(replyMessage);
                    Thread.Sleep(3000);
                }
            }
            else if (streamOption.Equals("2. Mechanical Engineering"))
            {
                dt = GetJobLinks("Mechanical Engineering");
                foreach (DataRow dRow in dt.Rows)
                {
                    var linkName = dRow["linkName"].ToString();
                    var linkURL = dRow["linkURL"].ToString();

                    var replyMessage = string.Format(@"[@Link](@url)", linkName, linkURL);
                    await context.PostAsync(replyMessage);
                    Thread.Sleep(3000);
                }
            }
            else if (streamOption.Equals("3. Computer Science"))
            {
                dt = GetJobLinks("Computer Science");
                foreach (DataRow dRow in dt.Rows)
                {
                    var linkName = dRow["linkName"].ToString();
                    var linkURL = dRow["linkURL"].ToString();

                    var replyMessage = string.Format(@"[{0}]({1})", linkName, linkURL);
                    await context.PostAsync(replyMessage);
                    Thread.Sleep(3000);
                }
            }
            else if (streamOption.Equals("4. International"))
            {
                dt = GetJobLinks("International");
                foreach (DataRow dRow in dt.Rows)
                {
                    var linkName = dRow["linkName"].ToString();
                    var linkURL = dRow["linkURL"].ToString();

                    var replyMessage = string.Format(@"[{0}]({1})", linkName, linkURL);
                    await context.PostAsync(replyMessage);
                    Thread.Sleep(3000);
                }
            }
            else if (streamOption.Equals("5. Electrical & Computer Engineering"))
            {
                dt = GetJobLinks("Electrical & Computer Engineering");
                foreach (DataRow dRow in dt.Rows)
                {
                    var linkName = dRow["linkName"].ToString();
                    var linkURL = dRow["linkURL"].ToString();

                    var replyMessage = string.Format(@"[{0}]({1})", linkName, linkURL);
                    await context.PostAsync(replyMessage);
                    Thread.Sleep(3000);
                }
            }
            Thread.Sleep(5000);

            PromptDialog.Text(
                context: context,
                resume: ShowStudentServices,
                prompt: "We also provide some services to students to help them in their job pursuit. Would you like to know aout them?",
                retry: "Sorry, could you please try again."
                );
        }
        public async Task ShowStudentServices(IDialogContext context, IAwaitable<string> strResponse)
        {
            var response = await strResponse;
            var confirmCodes = new List<string>() { "yes", "yup", "yeah", "ya", "ok", "okay", "k", "sure", "alright", "right", "ofcourse" };
            var isRequired = false;
            isRequired = confirmCodes.Any(response.ToString().ToLower().Contains);

            if (isRequired)
            {
                string services = GetServices();
                var reply3 = context.MakeMessage();
                reply3.Text = services;
                await context.PostAsync(reply3);
                Thread.Sleep(10000);

                var reply4 = context.MakeMessage();
                reply4.Text = "Stop into our office or contact us for assistance and/or take advantage of these available resources.";
                await context.PostAsync(reply4);
            }
            else
            {
                var reply5 = context.MakeMessage();
                reply5.Text = "Stop into our office or contact us for assistance and/or take advantage of these available resources.";
                await context.PostAsync(reply5);
            }
            context.Wait(this.MessageReceived);
        }

        #endregion

        #region Co-Op.Degree

        [LuisIntent("Co-Op.Degree.Overview")]
        public async Task GetCoopDegreeOverview(IDialogContext context, LuisResult result)
        {
            string reply = "Sure, I can get you some information about a Degree co-op credit program.";
            await context.PostAsync(reply);
            Thread.Sleep(3000);

            var strCoopOverview = GetCoOpOverview("Degree");
            await context.PostAsync(strCoopOverview);
            context.Wait(this.MessageReceived);
        }

        [LuisIntent("Co-Op.Degree.Requirements")]
        public async Task GetCoOpDegreeRequirements(IDialogContext context, LuisResult result)
        {
            string reply = "Sure, I can get you the requirements for a Degree Co-Op program.";
            await context.PostAsync(reply);
            Thread.Sleep(5000);

            var listReplies = new List<string>();
            listReplies = GetCoOpRequirements("Degree").Split('#').ToList();
            foreach (var strReply in listReplies)
            {
                reply = strReply;
                await context.PostAsync(reply);
                Thread.Sleep(5000);
            }
            context.Wait(this.MessageReceived);
        }

        [LuisIntent("Co-Op.Degree.Benefits")]
        public async Task GetCoOpDegreeBenefits(IDialogContext context, LuisResult result)
        {
            string reply = "Sure, there are a lot of benefits, the student can experience by registering for a Co-op program";
            await context.PostAsync(reply);
            Thread.Sleep(5000);

            var listReplies = new List<string>();
            listReplies = GetCoOpBenefits("Degree").Split('#').ToList();
            foreach (var strReply in listReplies)
            {
                reply = strReply;
                await context.PostAsync(reply);
                Thread.Sleep(5000);
            }
            context.Wait(this.MessageReceived);
        }

        [LuisIntent("Co-Op.Degree.PreRequisites")]
        public async Task GetCoOpDegreePreRequisites(IDialogContext context, LuisResult result)
        {
            string reply = "There are a few PreRequisites you would've to complete, before registering for a Co-Op program.";
            await context.PostAsync(reply);
            Thread.Sleep(3000);

            string reply2 = "Below are the PreRequisites you would have to complete -";
            await context.PostAsync(reply2);
            Thread.Sleep(3000);

            var strPreRequisites = GetCoOpPreRequisites("Degree");
            await context.PostAsync(strPreRequisites);
            context.Wait(this.MessageReceived);
        }

        #endregion

        #region Co-Op.Additive

        [LuisIntent("Co-Op.Additive.Overview")]
        public async Task GetCoopAdditiveOverview(IDialogContext context, LuisResult result)
        {
            string reply = "Sure, I can get you some information about a Additive co-op credit program.";
            await context.PostAsync(reply);
            Thread.Sleep(3000);

            var strCoopOverview = GetCoOpOverview("Additive");
            await context.PostAsync(strCoopOverview);
            context.Wait(this.MessageReceived);
        }

        [LuisIntent("Co-Op.Additive.Requirements")]
        public async Task GetCoOpAdditiveRequirements(IDialogContext context, LuisResult result)
        {
            string reply = "Sure, I can get you the requirements for a Degree Co-Op program.";
            await context.PostAsync(reply);
            Thread.Sleep(5000);

            var listReplies = new List<string>();
            listReplies = GetCoOpRequirements("Additive").Split('#').ToList();
            foreach (var strReply in listReplies)
            {
                reply = strReply;
                await context.PostAsync(reply);
                Thread.Sleep(5000);
            }
        }

        [LuisIntent("Co-Op.Additive.Benefits")]
        public async Task GetCoOpAdditiveBenefits(IDialogContext context, LuisResult result)
        {
            string reply = "Sure, there are a lot of benefits, the student can experience by registering for a Co-op program";
            await context.PostAsync(reply);
            Thread.Sleep(5000);

            var listReplies = new List<string>();
            listReplies = GetCoOpBenefits("Additive").Split('#').ToList();
            foreach (var strReply in listReplies)
            {
                reply = strReply;
                await context.PostAsync(reply);
                Thread.Sleep(5000);
            }
        }

        [LuisIntent("Co-Op.Additive.PreRequisites")]
        public async Task GetCoOpAdditivePreRequisites(IDialogContext context, LuisResult result)
        {
            string reply = "There are a few PreRequisites you would've to complete, before registering for a Co-Op program.";
            await context.PostAsync(reply);
            Thread.Sleep(3000);

            string reply2 = "Below are the PreRequisites you would have to complete -";
            await context.PostAsync(reply2);
            Thread.Sleep(3000);

            var strPreRequisites = GetCoOpPreRequisites("Additive");
            await context.PostAsync(strPreRequisites);
            context.Wait(this.MessageReceived);
        }

        #endregion

        #region Co-Op.Self-Report

        [LuisIntent("Co-Op.SelfReport.Overview")]
        public async Task GetCoopSelfReporteOverview(IDialogContext context, LuisResult result)
        {
            string reply = "Sure, I can get you some information about a Self Report co-op credit program.";
            await context.PostAsync(reply);
            Thread.Sleep(3000);

            var strCoopOverview = GetCoOpOverview("SelfReport");
            await context.PostAsync(strCoopOverview);
            context.Wait(this.MessageReceived);
        }

        [LuisIntent("Co-Op.SelfReport.Requirements")]
        public async Task GetCoOpSelfReportRequirements(IDialogContext context, LuisResult result)
        {
            string reply = "Sure, I can get you the requirements for a Self-Report Co-Op program.";
            await context.PostAsync(reply);
            Thread.Sleep(5000);

            var listReplies = new List<string>();
            listReplies = GetCoOpRequirements("SelfReport").Split('#').ToList();
            foreach (var strReply in listReplies)
            {
                reply = strReply;
                await context.PostAsync(reply);
                Thread.Sleep(5000);
            }
            context.Wait(this.MessageReceived);
        }

        [LuisIntent("Co-Op.SelfReport.Benefits")]
        public async Task GetCoOpSelfReportBenefits(IDialogContext context, LuisResult result)
        {
            string reply = "Sure, there are a lot of benefits, the student can experience by registering for a Co-op program";
            await context.PostAsync(reply);
            Thread.Sleep(5000);

            var listReplies = new List<string>();
            listReplies = GetCoOpBenefits("SelfReport").Split('#').ToList();
            foreach (var strReply in listReplies)
            {
                reply = strReply;
                await context.PostAsync(reply);
                Thread.Sleep(5000);
            }
            context.Wait(this.MessageReceived);
        }

        [LuisIntent("Co-Op.SelfReport.PreRequisites")]
        public async Task GetCoOpSelfReportPreRequisites(IDialogContext context, LuisResult result)
        {
            string reply = "There are a few PreRequisites you would've to complete, before registering for a Co-Op program.";
            await context.PostAsync(reply);
            Thread.Sleep(3000);

            string reply2 = "Below are the PreRequisites you would have to complete -";
            await context.PostAsync(reply2);
            Thread.Sleep(3000);

            var strPreRequisites = GetCoOpPreRequisites("SelfReport");
            await context.PostAsync(strPreRequisites);
            context.Wait(this.MessageReceived);
        }

        #endregion

        #region Co-Op.International F1 Students

        [LuisIntent("Co-Op.F1Students.Overview")]
        public async Task GetCoopF1StudentsOverview(IDialogContext context, LuisResult result)
        {
            string reply = "Sure, I can get you some information about a International / F1 Students Co-Op credit program";
            await context.PostAsync(reply);
            Thread.Sleep(3000);

            var strCoopOverview = GetCoOpOverview("F1Students");
            await context.PostAsync(strCoopOverview);
            context.Wait(this.MessageReceived);
        }

        [LuisIntent("Co-Op.F1Students.Requirements")]
        public async Task GetCoOpF1StudentsRequirements(IDialogContext context, LuisResult result)
        {
            string reply = "Sure, I can get you the requirements for an International / F1 Students Co-Op program.";
            await context.PostAsync(reply);
            Thread.Sleep(5000);

            var listReplies = new List<string>();
            listReplies = GetCoOpRequirements("F1Students").Split('#').ToList();
            foreach (var strReply in listReplies)
            {
                reply = strReply;
                await context.PostAsync(reply);
                Thread.Sleep(5000);
            }
            context.Wait(this.MessageReceived);
        }

        [LuisIntent("Co-Op.F1Students.Benefits")]
        public async Task GetCoOpF1StudentsBenefits(IDialogContext context, LuisResult result)
        {
            string reply = "Sure, there are a lot of benefits, the student can experience by registering for a Co-op program";
            await context.PostAsync(reply);
            Thread.Sleep(5000);

            var listReplies = new List<string>();
            listReplies = GetCoOpBenefits("F1Students").Split('#').ToList();
            foreach (var strReply in listReplies)
            {
                reply = strReply;
                await context.PostAsync(reply);
                Thread.Sleep(5000);
            }
            context.Wait(this.MessageReceived);
        }

        [LuisIntent("Co-Op.F1Students.PreRequisites")]
        public async Task GetCoOpF1StudentsPreRequisites(IDialogContext context, LuisResult result)
        {
            string reply = "There are a few PreRequisites you would've to complete, before registering for a Co-Op program.";
            await context.PostAsync(reply);
            Thread.Sleep(3000);

            string reply2 = "Below are the PreRequisites you would have to complete -";
            await context.PostAsync(reply2);
            Thread.Sleep(3000);

            var strPreRequisites = GetCoOpPreRequisites("F1Students");
            await context.PostAsync(strPreRequisites);
            context.Wait(this.MessageReceived);
        }

        #endregion

        #region Appointment Module
        [LuisIntent("NeedAppointment")]
        public async Task NeedAppointment(IDialogContext context, LuisResult result)
        {
            string reply1 = "Sure, I can book an appointment for you with the respective student advisor.";
            await context.PostAsync(reply1);
            Thread.Sleep(5000);

            string reply2 = "However, I require a few details to book your appointment.";
            await context.PostAsync(reply2);
            Thread.Sleep(5000);

            PromptDialog.Text(
                context: context,
                resume: ReceivedNameGetEmailID,
                prompt: "Could you please provide your name?",
                retry: "Sorry, could you please try again."
                );
        }
        public async Task ReceivedNameGetEmailID(IDialogContext context, IAwaitable<string> strUserName)
        {
            UserName = await strUserName; ;
            await context.PostAsync("Thank you " + UserName + ".");
            Thread.Sleep(2000);

            PromptDialog.Text(
                context: context,
                resume: ReceivedEmailIDGetProgram,
                prompt: "Could you please provide your Email ID?",
                retry: "Sorry, could you please try again."
                );
        }
        public async Task ReceivedEmailIDGetProgram(IDialogContext context, IAwaitable<string> strEmailID)
        {
            StudentEmailID = await strEmailID;
            await context.PostAsync("Thank you " + UserName + ", for sharing your Email Address.");
            Thread.Sleep(5000);

            await context.PostAsync("Please check the below list and respond with the code of the department, whose advisor you wish to meet.");
            Thread.Sleep(5000);

            PromptDialog.Choice(
                context: context,
                options: GetMajors(),
                promptStyle: PromptStyle.PerLine,
                resume: ReceivedProgramGetDate,
                prompt: "Programs - ",//Please respond with the code of the department.",
                retry: "Sorry, could you please try again."
                );
        }
        public async Task ReceivedProgramGetDate(IDialogContext context, IAwaitable<string> strProgram)
        {
            Program = await strProgram;
            Thread.Sleep(5000);
            PromptDialog.Text(
                context: context,
                resume: ConfirmDetails,
                prompt: "Please provide your preffered appointment date in (MM/DD/YYYY) format.",
                retry: "Sorry, could you please try again."
                );
        }
        public async Task ConfirmDetails(IDialogContext context, IAwaitable<string> strDateTime)
        {
            appointmentDateTime = Convert.ToDateTime(await strDateTime);

            await context.PostAsync("Thank you. Before I show you the available slots, could you check below and confirm if I had gotten everything right?");
            Thread.Sleep(5000);
            var department = Program.Substring(Program.IndexOf('-') + 2);

            //SMTP Approach
            //sendOutlookInvitationViaICSFile();

            PromptDialog.Text(
                context: context,
                resume: ShowTimeSlots,
                prompt: "Name - " + UserName + "\n" + "Email ID - " + StudentEmailID + "\n" + "\n" +
                "Department - " + department + "\n" + "Appointment Date - " + appointmentDateTime.ToShortDateString() + "\n\n" +
                "Please confirm with a yes or no",
                retry: "Sorry, could you please try again."
                );

            //await context.PostAsync(appointmentDateTime.ToShortDateString());
            //context.Wait(this.MessageReceived);

        }
        public async Task ShowTimeSlots(IDialogContext context, IAwaitable<string> confirmCode)
        {
            var confirmCodes = new List<string>() { "yes", "yup", "yeah", "ya", "ok", "okay", "k", "sure", "alright", "right", "ofcourse" };
            var isConfirmed = false;
            isConfirmed = confirmCodes.Any(confirmCode.ToString().ToLower().Contains);

            if (isConfirmed)
            {
                await context.PostAsync("Thank you " + UserName + "." +
                    "\n\nPlease check the below available time slots and respond with the slot number.");

                string addEventError = string.Empty;
                string refreshToken = string.Empty;
                string credentialError = string.Empty;
                var advisorEmailID = string.Empty;
                var selectedDate = appointmentDateTime;
                var listAvailableTimeSlots = new List<DateTime>();

                AdvisorEmailID = advisorEmailID = GetAdvisorInfo();

                var credential = GetUserCredential(advisorEmailID, out credentialError);
                if (credential != null && string.IsNullOrWhiteSpace(credentialError))
                {
                    //Save RefreshToken into Database 
                    refreshToken = credential.Token.RefreshToken;
                }

                listAvailableTimeSlots = GetCalenderEvents(refreshToken, advisorEmailID, "My Calendar Event", selectedDate, selectedDate.AddHours(24), out addEventError);
                //listTimeSlotChoices = GetTimeSlotChoices(listAvailableTimeSlots);

                var listTimeSlotChoices = new List<string>();
                var i = 0;
                foreach (var item in listAvailableTimeSlots)
                {
                    i++;
                    listTimeSlotChoices.Add(i.ToString() + " - " + item.ToString());
                }
                i = 0;

                PromptDialog.Choice(
                    context: context,
                    options: listTimeSlotChoices,
                    promptStyle: PromptStyle.PerLine,
                    resume: BookAppointment,
                    prompt: "Available timeslots on " + appointmentDateTime.ToShortDateString() + " - ",
                    retry: "Sorry, could you please try again."
                    );

            }
        }

        public async Task BookAppointment(IDialogContext context, IAwaitable<string> strTimeSlot)
        {
            var rawTimeSlot = await strTimeSlot;
            var TimeSlot = rawTimeSlot.Substring(rawTimeSlot.IndexOf('-') + 2);

            var finalTimeSlot = new Appointments.Calendar().FixAppointment(TimeSlot, appointmentDateTime.ToString(), StudentEmailID, AdvisorEmailID);

            var reply = context.MakeMessage();

            var advisor = GetAdvisorDetails();
            var heroCard = new HeroCard()
            {
                Title = advisor.AdvisorName,
                Subtitle = "Academic Advisor",
                Text = "Contact " + advisor.AdvisorName + " at " + advisor.AdvisorEmail,
                Images = new List<CardImage> { new CardImage(advisor.AdvisorImage) }
            };

            reply.Attachments = new List<Microsoft.Bot.Connector.Attachment> { heroCard.ToAttachment() };

            reply.Text = "Dear " + UserName + ", your appointment is confirmed on " + appointmentDateTime.ToShortDateString()
                + " at " + finalTimeSlot + " with - ";
            await context.PostAsync(reply);

            var reply2 = "Thank you for your interest in U of M - Dearborn.";
            await context.PostAsync(reply2);
            Thread.Sleep(2000);
            var reply3 = context.MakeMessage();
            reply3.Text = "Would you like me to assist you with anything else?";
            context.Wait(this.MessageReceived);
        }

        private List<DateTime> GetCalenderEvents(string refreshToken, string advisorEmailID, string summary, DateTime? start, DateTime? end, out string error)
        {
            string eventId = string.Empty;
            error = string.Empty;
            string serviceError;
            var listFreeTimeSlots = new List<DateTime>();
            var listBusyTimeSlots = new List<DateTime>();

            try
            {
                var calendarService = GetCalendarService(refreshToken, out serviceError);

                if (calendarService != null && string.IsNullOrWhiteSpace(serviceError))
                {
                    var list = calendarService.CalendarList.List().Execute();
                    var calendar = list.Items.SingleOrDefault(c => c.Summary == advisorEmailID);
                    if (calendar != null)
                    {
                        EventsResource.ListRequest request = calendarService.Events.List("primary");
                        request.TimeMin = start;//DateTime.Now;
                        request.ShowDeleted = false;
                        request.SingleEvents = true;
                        request.TimeMax = end;//DateTime.Now.AddDays(1);

                        request.MaxResults = 20;
                        request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;
                        Google.Apis.Calendar.v3.Data.Events events = request.Execute();

                        foreach (var test in events.Items)
                        {
                            var eventStart = Convert.ToDateTime(test.Start.DateTime.ToString());
                            listBusyTimeSlots.Add(eventStart);
                        }

                        var listAllTimeSlots = GetAllTimeSlots();

                        foreach (var timeSlot in listAllTimeSlots)
                        {
                            bool isBusy = false;
                            isBusy = listBusyTimeSlots.Any(x => x == timeSlot);
                            if (!isBusy)
                            {
                                listFreeTimeSlots.Add(timeSlot);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                eventId = string.Empty;
                error = ex.Message;
            }
            return listFreeTimeSlots;
        }
        private List<DateTime> GetAllTimeSlots()
        {
            var listTimeSlots = new List<DateTime>();
            var selectedDate = appointmentDateTime;
            var dateTimeNow = DateTime.Now;
            if (selectedDate.Date == dateTimeNow.Date)
            {
                //Searching for slots on same date
                int remainingMinutes = (selectedDate.Minute >= 30) ? 60 - selectedDate.Minute : 30 - selectedDate.Minute;
                var start = DateTime.Today;
                var clockQuery = from offset in Enumerable.Range(18, 16)
                                 select start.AddMinutes(30 * offset);
                foreach (var time in clockQuery)
                {
                    listTimeSlots.Add(time);
                }
            }
            else if (selectedDate.Date > dateTimeNow.Date)
            {
                //Searching for slots on same date
                int remainingMinutes = (selectedDate.Minute >= 30) ? 60 - selectedDate.Minute : 30 - selectedDate.Minute;
                var start = selectedDate;
                var clockQuery = from offset in Enumerable.Range(18, 16)
                                 select start.AddMinutes(30 * offset);
                foreach (var time in clockQuery)
                {
                    listTimeSlots.Add(time);
                }
            }
            else
            {
            }
            return listTimeSlots;
        }
        public Microsoft.Bot.Connector.Attachment CreateAdaptiveCardwithEntry()
        {
            var listPrograms = GetPrograms();
            var card = new AdaptiveCard()
            {
                Body = new List<CardElement>()
                {
                    new ChoiceSet() {
                        Id = "Major",
                        IsMultiSelect = false,
                        Style = ChoiceInputStyle.Compact,
                        Choices = listPrograms,
                    }
                },
                Actions = new List<ActionBase>()
                {
                    new SubmitAction()
                    {
                        Title = "Search Available Hours",
                        Speak = "<s>Search Available Hours</s>",
                        DataJson = "{ \"Value\": \"SearchTimeSlots\" }"
                    }
                }
            };
            Microsoft.Bot.Connector.Attachment attachment = new Microsoft.Bot.Connector.Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card
            };
            return attachment;
        }
        #endregion


        //    #region Outlook and SMTP approaches


        //    public void SetOutlookAppointment()
        //    {
        //        Microsoft.Office.Interop.Outlook.Application outlookApplication = new Microsoft.Office.Interop.Outlook.Application(); ;
        //        Microsoft.Office.Interop.Outlook.AppointmentItem newAppointment = (Microsoft.Office.Interop.Outlook.AppointmentItem)outlookApplication.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);


        //        newAppointment.Start = DateTime.Now.AddHours(2);
        //        newAppointment.End = DateTime.Now.AddHours(3);
        //        newAppointment.Location = "http://xyz.co/join/7235253940223 ";
        //        newAppointment.Body = "jigi is inviting you to a scheduled XYZ meeting. ";
        //        newAppointment.AllDayEvent = false;
        //        newAppointment.Subject = "XYZ";
        //        newAppointment.Recipients.Add("Roger Harui");
        //        Microsoft.Office.Interop.Outlook.Recipients sentTo = newAppointment.Recipients;
        //        Microsoft.Office.Interop.Outlook.Recipient sentInvite = null;
        //        sentInvite = sentTo.Add("Holly Holt");
        //        sentInvite.Type = (int)Microsoft.Office.Interop.Outlook.OlMeetingRecipientType.olRequired;
        //        sentInvite = sentTo.Add("David Junca ");
        //        sentInvite.Type = (int)Microsoft.Office.Interop.Outlook.OlMeetingRecipientType.olOptional;
        //        sentTo.ResolveAll();
        //        newAppointment.Save();
        //        newAppointment.Display(true);
        //    }
        //    public static string sendOutlookInvitationViaICSFile()
        //    {
        //        try
        //        {
        //            eAppointmentMail objApptEmail = new eAppointmentMail();
        //            objApptEmail.StartDate = appointmentDateTime + DateTime.Now.TimeOfDay;
        //            objApptEmail.Name = "Appointment with advisor";
        //            objApptEmail.Body = "You have an appointment at CECS advising, PEC 2021 on " + appointmentDateTime.ToString() +
        //                " with xyz";
        //            objApptEmail.Subject = "Appointment Confirmation";


        //            Configuration config = System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration(HttpContext.Current.Request.ApplicationPath);
        //            System.Net.Configuration.MailSettingsSectionGroup settings = (System.Net.Configuration.MailSettingsSectionGroup)config.GetSectionGroup("system.net/mailSettings");

        //            SmtpClient sc = new SmtpClient("smtp.gmail.com");

        //            System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();

        //            msg.From = new MailAddress(settings.Smtp.Network.UserName);
        //            //msg.To.Add(new MailAddress(objApptEmail.Email, objApptEmail.Name));
        //            msg.To.Add("hemanth.vaddireddy8@gmail.com,svaddire@umich.edu");
        //            msg.Subject = objApptEmail.Subject;
        //            msg.Body = objApptEmail.Body;

        //            StringBuilder str = new StringBuilder();
        //            str.AppendLine("BEGIN:VCALENDAR");
        //            str.AppendLine("PRODID:-//" + objApptEmail.Email);
        //            str.AppendLine("VERSION:2.0");
        //            str.AppendLine("METHOD:REQUEST");
        //            str.AppendLine("BEGIN:VEVENT");

        //            str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", objApptEmail.StartDate.ToUniversalTime().ToString("yyyyMMdd\\THHmmss\\Z")));
        //            str.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", (objApptEmail.EndDate - objApptEmail.StartDate).Minutes.ToString()));
        //            str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", objApptEmail.EndDate.ToUniversalTime().ToString("yyyyMMdd\\THHmmss\\Z")));

        //            str.AppendLine("LOCATION:" + objApptEmail.Location);
        //            str.AppendLine(string.Format("DESCRIPTION:{0}", objApptEmail.Body));
        //            str.AppendLine(string.Format("X-ALT-DESC;FMTTYPE=text/html:{0}", objApptEmail.Body));
        //            str.AppendLine(string.Format("SUMMARY:{0}", objApptEmail.Subject));
        //            str.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", objApptEmail.Email));

        //            str.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";RSVP=TRUE:mailto:{1}", msg.To[0].DisplayName, msg.To[0].Address));
        //            str.AppendLine("BEGIN:VALARM");
        //            str.AppendLine("TRIGGER:-PT15M");
        //            str.AppendLine("ACTION:DISPLAY");
        //            str.AppendLine("DESCRIPTION:Reminder");
        //            str.AppendLine("END:VALARM");
        //            str.AppendLine("END:VEVENT");
        //            str.AppendLine("END:VCALENDAR");
        //            System.Net.Mime.ContentType ct = new System.Net.Mime.ContentType("text/calendar");
        //            ct.Parameters.Add("method", "REQUEST");
        //            AlternateView avCal = AlternateView.CreateAlternateViewFromString(str.ToString(), ct);
        //            msg.AlternateViews.Add(avCal);
        //            NetworkCredential nc = new NetworkCredential(settings.Smtp.Network.UserName, settings.Smtp.Network.Password);
        //            sc.Port = settings.Smtp.Network.Port;
        //            sc.EnableSsl = true;
        //            sc.Credentials = nc;
        //            try
        //            {
        //                sc.Send(msg);
        //                return "Success";
        //            }
        //            catch
        //            {
        //                return "Fail";
        //            }
        //        }
        //        catch { }
        //        return string.Empty;
        //    }

        //#endregion



        #region Program Details
        [LuisIntent("Programs.BE.Overview")]
        public async Task GetProgramBEOverview(IDialogContext context, LuisResult result)
        {
            string reply = string.Empty;
            var listReplies = new List<string>();
            listReplies = GetProgramOverview("BE").Split('.').ToList();
            foreach (var strReply in listReplies)
            {
                reply = strReply;
                await context.PostAsync(reply);
                Thread.Sleep(3000);
            }
            Thread.Sleep(3000);
            PromptDialog.Text(
                context: context,
                resume: ShareBECurriculumAndSequence,
                prompt: "Do you want us to share the Program curriculum and sample sequence?",
                retry: "Sorry, could you please try again."
                );
        }
        public async Task ShareBECurriculumAndSequence(IDialogContext context, IAwaitable<string> confirmCode)
        {
            var replyMessage = context.MakeMessage();
            //var isRequired = await confirmCode;
            var confirmCodes = new List<string>() { "yes", "yup", "yeah", "ya", "ok", "okay", "k", "sure", "alright", "right", "ofcourse" };
            var isRequired = false;
            isRequired = confirmCodes.Any(confirmCode.ToString().ToLower().Contains);
            if (isRequired)
            {
                replyMessage.Text = "Check the attached BioEngineering course curriculum and Sample sequence";
                List<Microsoft.Bot.Connector.Attachment> attachments = new List<Microsoft.Bot.Connector.Attachment>();
                Microsoft.Bot.Connector.Attachment attachment1 = new Microsoft.Bot.Connector.Attachment
                {
                    ContentType = "application/pdf",
                    ContentUrl = HttpContext.Current.Server.MapPath("~/Courses/BE/Curriculum.pdf"),
                    Name = "Course Curriculum"
                };
                attachments.Add(attachment1);

                Microsoft.Bot.Connector.Attachment attachment2 = new Microsoft.Bot.Connector.Attachment
                {
                    ContentType = "application/pdf",
                    ContentUrl = HttpContext.Current.Server.MapPath("~/Courses/BE/SampleSequence.pdf"),
                    Name = "Course Sample Sequence"
                };
                attachments.Add(attachment2);
                replyMessage.Attachments = attachments;
            }
            else
            {
                replyMessage.Text = "Okay. Can I assist you with anything else?";
            }
            await context.PostAsync(replyMessage);

            context.Wait(this.MessageReceived);
        }
        [LuisIntent("Programs.BE.Office")]
        public async Task BEOfficeDetails(IDialogContext context, LuisResult result)
        {
            var dt = GetProgramOfficeDetails("BE");
            var address = string.Empty;
            var phoneNumber = string.Empty;
            var EmailID = string.Empty;
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dRow in dt.Rows)
                {
                    address = dRow["AddressLine"].ToString();
                    phoneNumber = dRow["PhoneNumber"].ToString();
                    EmailID = dRow["EmailID"].ToString();
                }
            }

            var reply = context.MakeMessage();
            reply.Text = "The advising office of BioEngineering is at " + address;
            await context.PostAsync(reply);

            Thread.Sleep(5000);

            reply.Text = "You can also reach us at " + phoneNumber + " or drop us an email at " + EmailID;
            await context.PostAsync(reply);

            context.Wait(this.MessageReceived);
        }

        [LuisIntent("Programs.DS.Overview")]
        public async Task GetProgramDSOverview(IDialogContext context, LuisResult result)
        {
            string reply = string.Empty;
            var listReplies = new List<string>();
            listReplies = GetProgramOverview("DS").Split('.').ToList();
            foreach (var strReply in listReplies)
            {
                reply = strReply;
                await context.PostAsync(reply);
                Thread.Sleep(3000);
            }
            Thread.Sleep(3000);
            PromptDialog.Text(
                context: context,
                resume: ShareDSCurriculumAndSequence,
                prompt: "Do you want us to share the Program curriculum and sample sequence?",
                retry: "Sorry, could you please try again."
                );
        }
        public async Task ShareDSCurriculumAndSequence(IDialogContext context, IAwaitable<string> confirmCode)
        {
            var replyMessage = context.MakeMessage();
            //var isRequired = await confirmCode;
            var confirmCodes = new List<string>() { "yes", "yup", "yeah", "ya", "ok", "okay", "k", "sure", "alright", "right", "ofcourse" };
            var isRequired = false;
            isRequired = confirmCodes.Any(confirmCode.ToString().ToLower().Contains);
            if (isRequired)
            {
                replyMessage.Text = "Check the attached Data Science course curriculum and Sample sequence";
                List<Microsoft.Bot.Connector.Attachment> attachments = new List<Microsoft.Bot.Connector.Attachment>();
                Microsoft.Bot.Connector.Attachment attachment1 = new Microsoft.Bot.Connector.Attachment
                {
                    ContentType = "application/pdf",
                    ContentUrl = HttpContext.Current.Server.MapPath("~/Courses/DS/Curriculum.pdf"),
                    Name = "Course Curriculum"
                };
                attachments.Add(attachment1);

                Microsoft.Bot.Connector.Attachment attachment2 = new Microsoft.Bot.Connector.Attachment
                {
                    ContentType = "application/pdf",
                    ContentUrl = HttpContext.Current.Server.MapPath("~/Courses/DS/SampleSequence.pdf"),
                    Name = "Course Sample Sequence"
                };
                attachments.Add(attachment2);
                replyMessage.Attachments = attachments;
            }
            else
            {
                replyMessage.Text = "Okay. Can I assist you with anything else?";
            }
            await context.PostAsync(replyMessage);

            context.Wait(this.MessageReceived);
        }
        [LuisIntent("Programs.DS.Office")]
        public async Task DSOfficeDetails(IDialogContext context, LuisResult result)
        {
            var dt = GetProgramOfficeDetails("DS");
            var address = string.Empty;
            var phoneNumber = string.Empty;
            var EmailID = string.Empty;
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dRow in dt.Rows)
                {
                    address = dRow["AddressLine"].ToString();
                    phoneNumber = dRow["PhoneNumber"].ToString();
                    EmailID = dRow["EmailID"].ToString();
                }
            }

            var reply = context.MakeMessage();
            reply.Text = "The advising office of Data Science is at " + address;
            await context.PostAsync(reply);

            Thread.Sleep(5000);

            reply.Text = "You can also reach us at " + phoneNumber + " or drop us an email at " + EmailID;
            await context.PostAsync(reply);

            context.Wait(this.MessageReceived);
        }

        [LuisIntent("Programs.SW.Overview")]
        public async Task GetProgramREOverview(IDialogContext context, LuisResult result)
        {
            string reply = string.Empty;
            var listReplies = new List<string>();
            listReplies = GetProgramOverview("SW").Split('.').ToList();
            foreach (var strReply in listReplies)
            {
                reply = strReply;
                await context.PostAsync(reply);
                Thread.Sleep(3000);
            }
            Thread.Sleep(3000);
            PromptDialog.Text(
                context: context,
                resume: ShareSWCurriculumAndSequence,
                prompt: "Do you want us to share the Program curriculum and sample sequence?",
                retry: "Sorry, could you please try again."
                );
        }
        public async Task ShareSWCurriculumAndSequence(IDialogContext context, IAwaitable<string> confirmCode)
        {
            var replyMessage = context.MakeMessage();
            //var isRequired = await confirmCode;
            var confirmCodes = new List<string>() { "yes", "yup", "yeah", "ya", "ok", "okay", "k", "sure", "alright", "right", "ofcourse" };
            var isRequired = false;
            isRequired = confirmCodes.Any(confirmCode.ToString().ToLower().Contains);
            if (isRequired)
            {
                replyMessage.Text = "Check the attached Software Engineering course curriculum and Sample sequence";
                List<Microsoft.Bot.Connector.Attachment> attachments = new List<Microsoft.Bot.Connector.Attachment>();
                Microsoft.Bot.Connector.Attachment attachment1 = new Microsoft.Bot.Connector.Attachment
                {
                    ContentType = "application/pdf",
                    ContentUrl = HttpContext.Current.Server.MapPath("~/Courses/SW/Curriculum.pdf"),
                    Name = "Course Curriculum"
                };
                attachments.Add(attachment1);

                Microsoft.Bot.Connector.Attachment attachment2 = new Microsoft.Bot.Connector.Attachment
                {
                    ContentType = "application/pdf",
                    ContentUrl = HttpContext.Current.Server.MapPath("~/Courses/SW/SampleSequence.pdf"),
                    Name = "Course Sample Sequence"
                };
                attachments.Add(attachment2);
                replyMessage.Attachments = attachments;
            }
            else
            {
                replyMessage.Text = "Okay. Can I assist you with anything else?";
            }
            await context.PostAsync(replyMessage);

            context.Wait(this.MessageReceived);
        }
        [LuisIntent("Programs.SW.Office")]
        public async Task REOfficeDetails(IDialogContext context, LuisResult result)
        {
            var dt = GetProgramOfficeDetails("SW");
            var address = string.Empty;
            var phoneNumber = string.Empty;
            var EmailID = string.Empty;
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dRow in dt.Rows)
                {
                    address = dRow["AddressLine"].ToString();
                    phoneNumber = dRow["PhoneNumber"].ToString();
                    EmailID = dRow["EmailID"].ToString();
                }
            }

            var reply = context.MakeMessage();
            reply.Text = "The advising office of Software Engineering is at " + address;
            await context.PostAsync(reply);

            Thread.Sleep(5000);

            reply.Text = "You can also reach us at " + phoneNumber + " or drop us an email at " + EmailID;
            await context.PostAsync(reply);

            context.Wait(this.MessageReceived);
        }
        #endregion



        //Database calls
        private string GetWelcomeMessage()
        {
            string reply = string.Empty;

            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select Value from Greeting_Welcome";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    var result = cmd.ExecuteScalar();
                    reply = result.ToString();
                }
                catch (Exception e)
                {
                    if (e.Message != null) { }
                }
                finally
                {
                    sqlConn.Close();
                }
            }
            return reply;
        }
        private string GetAdvisingOfficeAddress()
        {
            string reply = string.Empty;

            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select Value from Advisory_Office_Location";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    var result = cmd.ExecuteScalar();
                    reply = result.ToString();
                }
                catch (Exception e)
                {
                    if (e.Message != null) { }
                }
                finally
                {
                    sqlConn.Close();
                }
            }
            return reply;
        }
        private string GetAdmissionChangeInfo()
        {
            string reply = string.Empty;

            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select Value from AdmissionChange";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    var result = cmd.ExecuteScalar();
                    reply = result.ToString();
                }
                catch (Exception e)
                {
                    if (e.Message != null) { }
                }
                finally
                {
                    sqlConn.Close();
                }
            }
            return reply;
        }
        private string GetCoOpOverview()
        {
            string reply = string.Empty;

            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select Value from CoOp_Overview";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    var result = cmd.ExecuteScalar();
                    reply = result.ToString();
                }
                catch (Exception e)
                {
                    if (e.Message != null) { }
                }
                finally
                {
                    sqlConn.Close();
                }
            }
            return reply;
        }
        private DataTable GetCoOpOffice()
        {
            string reply = string.Empty;
            DataTable dt = new DataTable();

            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select * from CoOp_Office";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    SqlDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                    dt.Load(dr);
                    return dt;
                }
                catch (Exception e)
                {
                    if (e.Message != null) { }
                }
                finally
                {
                    sqlConn.Close();
                }
                return dt;
            }
        }
        public List<string> GetMajors()
        {
            var listPrograms = new List<string>();
            string reply = string.Empty;
            DataTable dt = new DataTable();
            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select * from Departments";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    SqlDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                    dt.Load(dr);

                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow dRow in dt.Rows)
                        {
                            var department = string.Empty;
                            department = dRow["Code"].ToString() + " - " + dRow["DeptName"].ToString();
                            listPrograms.Add(department);
                        }
                    }
                }
                catch (Exception e)
                {
                    if (e.Message != null) { }
                }
                finally
                {
                    sqlConn.Close();
                }
                return listPrograms;
            }
        }
        public List<AdaptiveCards.Choice> GetPrograms()
        {
            var listPrograms = new List<Majors>();
            var listChoices = new List<AdaptiveCards.Choice>();

            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select * from Majors order by MajorName";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    SqlDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                    DataTable dt = new DataTable();
                    dt.Load(dr);

                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            var objMajor = new Majors();
                            objMajor.MajorName = row["MajorName"].ToString();
                            objMajor.MajorID = Convert.ToInt16(row["MajorID"].ToString());

                            listPrograms.Add(objMajor);
                        }
                        listChoices = GetChoices(listPrograms);
                        return listChoices;
                    }
                }
                catch (Exception e)
                {
                    //Do nothing
                }
                finally
                {
                    sqlConn.Close();
                }

                return listChoices;
            }
        }
        private List<AdaptiveCards.Choice> GetChoices(List<Majors> listMajors)
        {
            var listChoices = new List<AdaptiveCards.Choice>();

            if (listMajors.Count > 0)
            {
                foreach (var major in listMajors)
                {
                    var choice = new AdaptiveCards.Choice();
                    choice.Title = major.MajorName;
                    choice.Value = major.MajorID.ToString();

                    listChoices.Add(choice);
                }
            }
            return listChoices;
        }
        private Advisors GetAdvisorDetails()
        {
            var listAdvisors = new List<Advisors>();
            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select * from AdvisorInfo where AdvisorEmail = @AdvisorEmail";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    cmd.Parameters.AddWithValue("@AdvisorEmail", AdvisorEmailID);

                    SqlDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                    DataTable dt = new DataTable();
                    dt.Load(dr);

                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            var objAdvisor = new Advisors();
                            objAdvisor.AdvisorID = Convert.ToInt16(row["AdvisorID"].ToString());
                            objAdvisor.AdvisorName = row["AdvisorName"].ToString();
                            objAdvisor.AdvisorEmail = row["AdvisorEmail"].ToString();
                            objAdvisor.AdvisorImage = row["AdvisorImage"].ToString();

                            listAdvisors.Add(objAdvisor);
                        }

                        return listAdvisors.FirstOrDefault();
                    }
                }
                catch (Exception e)
                {
                    if (e.Message != null) { }
                }
                finally
                {
                    sqlConn.Close();
                }
                return listAdvisors.FirstOrDefault();
            }
        }
        private string GetAdvisorInfo()
        {
            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                var department = Program.Substring(Program.IndexOf('-') + 2);
                var advisorEmail = string.Empty;
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select MajorID from Majors where MajorName = @MajorName";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    cmd.Parameters.AddWithValue("@MajorName", department);
                    var result = cmd.ExecuteScalar();

                    var sqlQuery2 = "select AdvisorEmail from AdvisorInfo where AdvisorID in "
                        + "(select AdvisorID from MajorAdvisorInfo where MajorID = @MajorID)";
                    SqlCommand cmd2 = new SqlCommand(sqlQuery2, sqlConn);
                    cmd2.Parameters.AddWithValue("@MajorID", result);
                    var result2 = cmd2.ExecuteScalar();

                    advisorEmail = result2.ToString();
                    return advisorEmail;
                }
                catch (Exception e)
                {
                    if (e.Message != null) { }
                }
                finally
                {
                    sqlConn.Close();
                }
                return advisorEmail;
            }
        }
        private string GetProgramOverview(string ProgramCode)
        {
            var strProgramOverview = string.Empty;
            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select Overview from Programs where CourseCode = @ProgramCode";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    cmd.Parameters.AddWithValue("@ProgramCode", ProgramCode);
                    var result = cmd.ExecuteScalar();
                    strProgramOverview = result.ToString();
                    return strProgramOverview;
                }
                catch (Exception e)
                {
                    if (e.Message != null) { }
                }
                finally
                {
                    sqlConn.Close();
                }
            }
            return strProgramOverview;
        }
        private DataTable GetProgramOfficeDetails(string ProgramCode)
        {
            string reply = string.Empty;
            DataTable dt = new DataTable();

            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select * from Programs where CourseCode = @ProgramCode";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    cmd.Parameters.AddWithValue("@ProgramCode", ProgramCode);
                    SqlDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                    dt.Load(dr);
                    return dt;
                }
                catch (Exception e)
                {
                    if (e.Message != null) { }
                }
                finally
                {
                    sqlConn.Close();
                }
                return dt;
            }
        }
        private DataTable GetStudentEligibility(string strRegOption)
        {
            DataTable dt = new DataTable();

            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select * from Coop_Student_Eligibility where CoopOption = @CoopOption";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    cmd.Parameters.AddWithValue("@CoopOption", strRegOption);
                    SqlDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                    dt.Load(dr);
                    return dt;
                }
                catch (Exception e)
                {
                    if (e.Message != null) { }
                }
                finally
                {
                    sqlConn.Close();
                }
                return dt;
            }
        }
        private string GetServices()
        {
            string services = string.Empty;
            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select * from CoOp_StudentService where ServiceCode = 'CoOp'";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    services = cmd.ExecuteScalar().ToString();
                    return services;
                }
                catch (Exception e)
                {
                    if (e.Message != null) { }
                }
                finally
                {
                    sqlConn.Close();
                }
                return services;
            }
        }
        private DataTable GetJobLinks(string strStreamName)
        {
            string reply = string.Empty;
            DataTable dt = new DataTable();

            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select * from JobSearchLinks where linkName in (@StreamName, 'LinkedIn', 'NASA Jobs')";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    cmd.Parameters.AddWithValue("@StreamName", strStreamName);
                    SqlDataReader dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                    dt.Load(dr);
                    return dt;
                }
                catch (Exception e)
                {
                    if (e.Message != null) { }
                }
                finally
                {
                    sqlConn.Close();
                }
                return dt;
            }
        }
        private string GetCoOpRequirements(string strCoOpOption)
        {
            var strRequirements = string.Empty;
            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select Requirements from Coop_Student_Eligibility where CoopOption = @CoopOption";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    cmd.Parameters.AddWithValue("@CoopOption", strCoOpOption);
                    strRequirements = cmd.ExecuteScalar().ToString();
                    return strRequirements;
                }
                catch (Exception e)
                {
                    if (e.Message != null) { }
                }
                finally
                {
                    sqlConn.Close();
                }
                return strRequirements;
            }
        }
        private string GetCoOpBenefits(string strCoOpOption)
        {
            var strBenefits = string.Empty;
            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select Benefits from Coop_Student_Eligibility where CoopOption = @CoopOption";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    cmd.Parameters.AddWithValue("@CoopOption", strCoOpOption);
                    strBenefits = cmd.ExecuteScalar().ToString();
                    return strBenefits;
                }
                catch (Exception e)
                {
                    if (e.Message != null) { }
                }
                finally
                {
                    sqlConn.Close();
                }
                return strBenefits;
            }
        }
        private string GetCoOpPreRequisites(string strCoOpOption)
        {
            var strPreRequisites = string.Empty;
            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select PreRequisites from Coop_Student_Eligibility where CoopOption = @CoopOption";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    cmd.Parameters.AddWithValue("@CoopOption", strCoOpOption);
                    strPreRequisites = cmd.ExecuteScalar().ToString();
                    return strPreRequisites;
                }
                catch (Exception e)
                {
                    if (e.Message != null) { }
                }
                finally
                {
                    sqlConn.Close();
                }
                return strPreRequisites;
            }
        }
        private string GetCoOpOverview(string strCoOpOption)
        {
            var strCoopOverview = string.Empty;
            using (SqlConnection sqlConn = new SqlConnection(connString))
            {
                sqlConn.Open();
                try
                {
                    var sqlQuery = "select Overview from Coop_Student_Eligibility where CoopOption = @CoopOption";
                    SqlCommand cmd = new SqlCommand(sqlQuery, sqlConn);
                    cmd.Parameters.AddWithValue("@CoopOption", strCoOpOption);
                    strCoopOverview = cmd.ExecuteScalar().ToString();
                    return strCoopOverview;
                }
                catch (Exception e)
                {
                    if (e.Message != null) { }
                }
                finally
                {
                    sqlConn.Close();
                }
                return strCoopOverview;
            }
        }

        #region AdvisorAuthentication

        public static UserCredential GetGoogleUserCredentialByRefreshToken(string refreshToken, out string error)
        {
            Google.Apis.Auth.OAuth2.Responses.TokenResponse respnseToken = null;
            UserCredential credential = null;
            string flowError;
            error = string.Empty;
            try
            {
                // Get a new IAuthorizationCodeFlow instance
                IAuthorizationCodeFlow flow = GoogleAuthorizationCodeFlow(out flowError);

                respnseToken = new Google.Apis.Auth.OAuth2.Responses.TokenResponse() { RefreshToken = refreshToken };

                // Get a new Credential instance                
                if ((flow != null && string.IsNullOrWhiteSpace(flowError)) && respnseToken != null)
                {
                    credential = new UserCredential(flow, "user", respnseToken);

                }

                // Get a new Token instance
                if (credential != null)
                {
                    bool success = credential.RefreshTokenAsync(CancellationToken.None).Result;
                }

                // Set the new Token instance
                if (credential.Token != null)
                {
                    string newRefreshToken = credential.Token.RefreshToken;
                }
            }
            catch (Exception ex)

            {
                credential = null;
                error = "UserCredential failed: " + ex.ToString();
                //string newRefreshToken = credential.Token.RefreshToken;
            }
            return credential;
        }
        public static IAuthorizationCodeFlow GoogleAuthorizationCodeFlow(out string error)
        {
            IAuthorizationCodeFlow flow = null;
            error = string.Empty;

            try
            {
                flow = new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
                {
                    ClientSecrets = GoogleClientSecrets,
                    Scopes = Scopes
                });
            }
            catch (Exception ex)
            {
                flow = null;
                error = "Failed to AuthorizationCodeFlow Initialization: " + ex.ToString();
            }

            return flow;
        }
        public static CalendarService GetCalendarService(string refreshToken, out string error)
        {
            CalendarService calendarService = null;
            string credentialError;
            error = string.Empty;
            try
            {
                var credential = GetGoogleUserCredentialByRefreshToken(refreshToken, out credentialError);
                if (credential != null && string.IsNullOrWhiteSpace(credentialError))
                {
                    calendarService = new CalendarService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName
                    });
                }
            }
            catch (Exception ex)
            {
                calendarService = null;
                error = "Calendar service failed: " + ex.ToString();
            }
            return calendarService;
        }
        public static string[] Scopes =
        {
                                CalendarService.Scope.Calendar,
                                CalendarService.Scope.CalendarReadonly,
                            };
        public static UserCredential GetUserCredential(string advisorEmailAddress, out string error)
        {
            UserCredential credential = null;
            error = string.Empty;

            try
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                new ClientSecrets
                {
                    ClientId = ClientId,
                    ClientSecret = ClientSecret
                },
                Scopes,
                /*"hemanth260292@gmail.com",*/advisorEmailAddress,
                CancellationToken.None,
                null).Result;
            }
            catch (Exception ex)
            {
                credential = null;
                error = "Failed to UserCredential Initialization: " + ex.ToString();
            }

            return credential;
        }
        #endregion
    }
}