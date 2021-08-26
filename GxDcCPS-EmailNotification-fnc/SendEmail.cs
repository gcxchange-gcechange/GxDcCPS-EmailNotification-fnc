using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Script.Serialization;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Graph;

namespace GxDcCPSEmailNotificationfnc
{
    public static class SendEmail
    {
        [FunctionName("SendEmail")]
        public static void Run([QueueTrigger("email-info", Connection = "")] SiteInfo myQueueItem, TraceWriter log)
        {
            log.Info($"C# Queue trigger function processed: {myQueueItem}");
            string EmailSender = ConfigurationManager.AppSettings["Emailsender"];
            string HD_Email = ConfigurationManager.AppSettings["HD_Email"];
      
            var siteUrl = myQueueItem.siteUrl;
            var displayName = myQueueItem.displayName;
            var emails = myQueueItem.emails;
            var comments = myQueueItem.comments;
            var status = myQueueItem.status;
            var requester = myQueueItem.requesterName;
            var requesterEmail = myQueueItem.requesterEmail;

            var authResult = GetOneAccessToken();
            var graphClient = GetGraphClient(authResult);
            SendEmailToUser(graphClient, log, emails, siteUrl, displayName, status, comments, requester, requesterEmail, EmailSender, HD_Email);

        }
        /// <summary>
        /// Get graph client
        /// </summary>
        /// <param name="authResult"></param>
        /// <returns></returns>
        public static GraphServiceClient GetGraphClient(string authResult)
        {
            GraphServiceClient graphClient = new GraphServiceClient(
                 new DelegateAuthenticationProvider(
            async (requestMessage) =>
            {
                requestMessage.Headers.Authorization =
                    new AuthenticationHeaderValue("bearer",
                    authResult);
            }));
            return graphClient;
        }
        /// <summary>
        /// Get access token from AAD
        /// </summary>
        /// <returns></returns>
        public static string GetOneAccessToken()
        {
            string token = "";
            string CLIENT_ID = ConfigurationManager.AppSettings["CLIENT_ID"];
            string CLIENT_SECERET = ConfigurationManager.AppSettings["CLIENT_SECRET"];
            string TENAT_ID = ConfigurationManager.AppSettings["Tenant_ID"];
            string TOKEN_ENDPOINT = "";
            string MS_GRAPH_SCOPE = "";
            string GRANT_TYPE = "";

            try
            {

                TOKEN_ENDPOINT = "https://login.microsoftonline.com/" + TENAT_ID + "/oauth2/v2.0/token";
                MS_GRAPH_SCOPE = "https://graph.microsoft.com/.default";
                GRANT_TYPE = "client_credentials";

            }
            catch (Exception e)
            {
                Console.WriteLine("A error happened while search config file");
            }
            try
            {
                HttpWebRequest request = WebRequest.Create(TOKEN_ENDPOINT) as HttpWebRequest;
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                StringBuilder data = new StringBuilder();
                data.Append("client_id=" + HttpUtility.UrlEncode(CLIENT_ID));
                data.Append("&scope=" + HttpUtility.UrlEncode(MS_GRAPH_SCOPE));
                data.Append("&client_secret=" + HttpUtility.UrlEncode(CLIENT_SECERET));
                data.Append("&GRANT_TYPE=" + HttpUtility.UrlEncode(GRANT_TYPE));
                byte[] byteData = UTF8Encoding.UTF8.GetBytes(data.ToString());
                request.ContentLength = byteData.Length;
                using (Stream postStream = request.GetRequestStream())
                {
                    postStream.Write(byteData, 0, byteData.Length);
                }

                // Get response

                using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
                {

                    using (var reader = new StreamReader(response.GetResponseStream()))
                    {
                        JavaScriptSerializer js = new JavaScriptSerializer();
                        var objText = reader.ReadToEnd();
                        LgObject myojb = (LgObject)js.Deserialize(objText, typeof(LgObject));
                        token = myojb.access_token;
                    }

                }
                return token;
            }
            catch (Exception e)
            {
                Console.WriteLine("A error happened while connect to server please check config file");
                return "error";
            }
        }
        /// <summary>
        /// Send email to users, when status is submitted, rejected and team created.
        /// </summary>
        /// <param name="graphClient"></param>
        /// <param name="log"></param>
        /// <param name="emails"></param>
        /// <param name="siteUrl"></param>
        /// <param name="displayName"></param>
        /// <param name="status"></param>
        /// <param name="comments"></param>
        /// <param name="requester"></param>
        /// <param name="requesterEmail"></param>
        public static async void SendEmailToUser(GraphServiceClient graphClient, TraceWriter log, string emails, string siteUrl, string displayName, string status, string comments, string requester, string requesterEmail, string EmailSender, string HD_Email)
        {

            switch (status)
            {
                case "Submitted":
                    var submitMsg = new Message
                    {
                        Subject = "We received your request for a gcxchange space / Nous avons reçu votre demande d’espace gcéchange",
                        Body = new ItemBody
                        {
                            ContentType = BodyType.Html,
                            Content = $"<a href=\"#fr\">La version française suit</a><br><br>Hello there from gcxchange! &nbsp;<br><br>Thank you for your request for a team space for <strong>{displayName}</strong>.<br><br><strong>So, what's next?</strong><br><br>We will review your request and if all is in order, it will be created. &nbsp;At that time, you will receive a second e-mail with the details you need to start collaborating.<br><br>This process takes from 24 to 48 business hours from the time you submitted your request.<br><br><strong>In the meantime</strong>,<br><br>Have a look at our <a href=\"https://gcxgce.sharepoint.com/sites/Support\">Support site</a> for: &nbsp;<ul><li>Tips for becoming a collaboration pro</li><li>Best information management practices for gcxchange</li><li>Terms and conditions for gcxchange</li></ul>We at gcxchange are always keeping accessibility in mind, guided by the Accessible Canada Act. <a href=\"https://gcxgce.sharepoint.com/sites/Support\">Visit our support page</a> to learn more. <br><br>Have a great day, <br><br>The gcxchange team&nbsp<br><hr><br><p id=\"fr\">Bonjour de gcéchange! &nbsp;</p><br>Nous vous remercions d’avoir demandé un espace d’équipe pour <strong>{displayName}</strong>.<br><br><strong>Quelle est la suite? </strong><br><br>Nous examinerons votre demande et, si tout est conforme, elle sera créée. À ce moment-là, vous recevrez un deuxième courriel contenant les détails dont vous avez besoin pour commencer à collaborer. <br><br>Ce processus prend de 24 à 48 heures à compter du moment où vous présentez votre demande.<br><br><strong>Entretemps</strong>, <br><br>Jetez un coup d’œil à notre <a href=\"https://gcxgce.sharepoint.com/sites/Support/SitePages/fr/Home.aspx\">site de soutien</a>. Vous y trouverez:<ul><li>Des conseils pour devenir un professionnel de la collaboration </li><li>Les pratiques exemplaires de gestion de l’information pour gcxchange</li><li>Les modalités de gcéchange </li></ul>Chez gcéchange, nous gardons toujours à l’esprit l’accessibilité en fonction de la Loi canadienne sur l’accessibilité. <a href=\"https://gcxgce.sharepoint.com/sites/Support/SitePages/fr/Home.aspx\">Visitez notre page de soutien</a> pour en savoir davantage. <br><br>Bonne journée!<br>L’équipe gcéchange"
                        },
                        ToRecipients = new List<Recipient>()
                    {
                        new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = $"{requesterEmail}"
                            }
                        }
                    },

                    };
                    try
                    {
                     await graphClient.Users[EmailSender]
                        .SendMail(submitMsg)
                        .Request()
                        .PostAsync();

                    log.Info($"Send email to {requesterEmail} successfully.");
                    }
                    catch (ServiceException e)
                    {
                        log.Info($"Error: {e.Message}");
                    }
                
                    break;

                case "Rejected":
                    var rejectMsg = new Message
                    {
                        Subject = "Sorry, your gcxchange team space was not created /  Nous avons le regret de vous aviser que votre espace d’équipe gcéchange n’a pas été créé",
                        Body = new ItemBody
                        {
                            ContentType = BodyType.Html,
                            Content = $"<a href=\"#fr\">La version française suit</a><br><br>We have looked at your request for a gcxchange team space <strong>{displayName}</strong>.<br><br>Your requested team space has not been approved at this time. &nbsp;We were not able to create it for the following reason(s):<br><br>{comments}<br><br>We are here to help! &nbsp;If you are still interested in obtaining a team space or you think our decision has been made in error, please contact us via our <a href=\"https://gcxgce.sharepoint.com/sites/Support\">Support site</a>.<br><br>Come back soon to gcxchange to stay current, connect, and collaborate.<br><br>Have a great day,<br><br>The gcxchange team<br><br>We at gcxchange are always keeping accessibility in mind, guided by the Accessible Canada Act. <a href=\"https://gcxgce.sharepoint.com/sites/Support\">Visit our support site</a> to learn more.&nbsp;<br><br><hr><br><br><p id=\"fr\">Nous avons examiné votre demande d’espace d’équipe gcéchange <strong>{displayName}</strong>.</p><br>L’espace d’équipe demandé n’a pas été approuvé pour l’instant. Nous n’avons pas pu le créer pour les raisons suivantes:<br><br> {comments} <br><br>Nous sommes là pour vous aider! Si vous souhaitez toujours obtenir un espace d’équipe ou si vous pensez que notre décision est erronée, veuillez communiquer avec nous par l’intermédiaire de notre <a href=\"https://gcxgce.sharepoint.com/sites/Support/SitePages/fr/Home.aspx\">site de soutien.</a><br><br>Revenez bientôt à gcéchange pour rester à jour, vous connecter et collaborer.<br><br>Bonne journée!<br><br>L’équipe gcéchange<br><br>Chez gcéchange, nous gardons toujours à l’esprit l’accessibilité en fonction de la Loi canadienne sur l’accessibilité. <a href=\"https://gcxgce.sharepoint.com/sites/Support/SitePages/fr/Home.aspx\">Visitez notre site de soutien</a> pour en savoir davantage.&nbsp;"
                        },
                        ToRecipients = new List<Recipient>()
                    {
                        new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = $"{requesterEmail}"
                            }
                        }
                    },

                    };

                    //     var saveToSentItems = false;

                    await graphClient.Users[EmailSender]
                        .SendMail(rejectMsg)
                        .Request()
                        .PostAsync();

                    log.Info($"Send email to {requesterEmail} successfully.");
                    break;

                case "Team Created":
                    Regex emailRegex = new Regex(@"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*",
RegexOptions.IgnoreCase);
                    MatchCollection ownersEmail = emailRegex.Matches(emails);
                    foreach (Match i in ownersEmail)
                    {
                        var message = new Message
                        {
                            Subject = "Your gcxchange team space is ready / Votre espace d’équipe gcéchange est prêt",
                            Body = new ItemBody
                            {
                                ContentType = BodyType.Html,
                                Content = $"<a href=\"#fr\">La version française suit</a><p><a id=\"en\"></a><strong>Welcome to collaboration with gcxchange!</strong><br><br>{requester},your community,{displayName} is now ready for take-off!<br><br><strong>How do you find and access your community on gcxchange?</strong><br><br>Easy! Follow this <a href='{siteUrl}'>link.</a><br><br><strong>How do you find and access your community on Microsoft (MS) Teams?</strong></p><p>To access your Teams channel, select the button on the left-hand side of your community page called “Conversations”.<br>Pro tip: open Teams using the web app the first time you access your Teams channel for a smoother transition.<br>From now on, when you open MS Teams you can switch between your new gcxchange Teams account and your department&#39;s Teams account.  To switch between gcxchange and your department:</p><ol><li>Open MS Teams</li><li>Select your Avatar on the top right of the page</li><li>Select either gcxchange or your department, and Teams will switch over</li></ol><p>If you do not see the drop down, try to restart your device and launch MS Teams again.  If the drop down is still not visible, please submit a service request to gcxchange via our <a href='https://gcxgce.sharepoint.com/sites/Support/SitePages/Submit-a-ticket.aspx'>Support</a> page.</p><strong>So, what's next?</strong><ol><li>Start adding content to your community page on gcxchange.</li><li>Start conversations and upload files to your Teams channel.</li></ol><p>Visit our gcxchange <a href='https://gcxgce.sharepoint.com/sites/Support/'>Support</a> page for more information about communities.</p><strong>Questions? Need additional support?</strong><p>You can contact the support team at <a href='mailto:support-soutien@gcx-gce.gc.ca'>support-soutien@gcx-gce.gc.ca</a>.<br>We at gcxchange are always keeping accessibility in mind, guided by the Accessible Canada Act. Visit our <a href='https://gcxgce.sharepoint.com/sites/Support/'>Support</a> page to learn more.</p><p>___________________________________________________________________________________________________________________</p><a href=\"#en\">The English version precedes</a><br><br><a id=\"fr\"><strong>Bienvenue à la collaboration avec gcéchange!</strong><br><br><p>{requester}, votre communauté, {displayName}, est maintenant prête pour le décollage!</p><strong>Comment trouver votre communauté et y accéder sur gcéchange?</strong><p>C’est facile! Suivez ce <a href='{siteUrl}'>lien</a>.</p><strong>Comment trouver votre communauté et y accéder dans Microsoft (MS) Teams?</strong><p>Pour accéder à votre canal dans Teams, sélectionnez le bouton à gauche de la page de votre communauté appelé « Conversations ».<br>Conseil d’expert : ouvrez Teams à partir de l’application Web la première fois que vous accédez à votre canal dans Teams pour faciliter la transition.<br>À partir de maintenant, lorsque vous ouvrez MS Teams, vous pouvez passer de votre nouveau compte Teams gcéchange au compte Teams de votre ministère. Pour passer de votre compte gcéchange à celui de votre ministère :</p><ol><li>Ouvrez MS Teams.</li><li>Sélectionnez votre Avatar dans le coin supérieur droit de la page.</li><li>Sélectionnez gcéchange ou votre ministère, et Teams passera à l’autre.</li></ol><p>Si vous ne voyez pas le menu déroulant, essayez de redémarrer votre appareil et de relancer MS Teams. Si le menu déroulant n’est toujours pas visible, veuillez soumettre une demande de service à gcéchange en envoyant un commentaire sur notre page de <a href='https://gcxgce.sharepoint.com/sites/Support/SitePages/Submit-a-ticket.aspx'>soutien</a>.</p><strong>Et ensuite?</strong><ol><li>Commencez à ajouter du contenu à votre page communautaire sur gcéchange.</li><li>Commencez des conversations et téléchargez des fichiers dans votre canal dans Teams.</li></ol><p>Visitez notre page de <a href='https://gcxgce.sharepoint.com/sites/Support/'>soutien</a> gcéchange pour obtenir de plus amples renseignements sur les communautés.</p><strong>Avez-vous des questions? Avez-vous besoin de soutien supplémentaire?</strong><p>Vous pouvez communiquer avec l’équipe de soutien à l’adresse <a href='mailto:support-soutien@gcx-gce.gc.ca'>support-soutien@gcx-gce.gc.ca</a>.<br>Chez gcéchange, nous gardons toujours à l’esprit l’accessibilité, guidée par la Loi canadienne sur l’accessibilité. Visitez notre page de <a href='https://gcxgce.sharepoint.com/sites/Support/'>soutien</a> pour en savoir plus.</p>"
                            },
                            ToRecipients = new List<Recipient>()
                    {
                        new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = $"{i}"
                            }
                        }
                    },

                        };

                        var saveToSentItems = false;

                        await graphClient.Users[EmailSender]
                            .SendMail(message, saveToSentItems)
                            .Request()
                            .PostAsync();

                        log.Info($"Send email to {i} successfully.");
                    }
                    break;

                    case "Notif_HD":
                    var HD_Msg = new Message
                    {
                        Subject = $"New pending request! {displayName}",
                        Body = new ItemBody
                        {
                            ContentType = BodyType.Html,
                            Content = $"<a href=\"https://gcxgce.sharepoint.com/teams/scw\">Click here</a> to review the request."
                        },
                        ToRecipients = new List<Recipient>()
                    {
                        new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = $"{HD_Email}"
                            }
                        }
                    },

                    };
                    try
                    {
                     await graphClient.Users[EmailSender]
                        .SendMail(HD_Msg)
                        .Request()
                        .PostAsync();

                    log.Info($"Send email to {HD_Email} successfully.");
                    }
                    catch (ServiceException e)
                    {
                        log.Info($"Error: {e.Message}");
                    }
                
                    break;

                default:
                    log.Info($"The status was {status}. This status is not part of the switch statement.");
                    break;
            };

        }
    }
}
