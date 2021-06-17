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
                        Subject = "We received your request for a gcxchange space / Nous avons re�u votre demande d�espace gcxchange",
                        Body = new ItemBody
                        {
                            ContentType = BodyType.Html,
                            Content = $"<a href=\"#fr\">La version fran�aise suit</a><br><br>Hello there from gcxchange! &nbsp;<br><br>Thank you for your request for a team space for <strong>{displayName}</strong>.<br><br><strong>So, what�s next?</strong><br><br>We will review your request and if all is in order, it will be created. &nbsp;At that time, you will receive a second e-mail with the details you need to start collaborating.<br><br>This process takes from 24 to 48 hours from the time you submitted your request.<br><br><strong>In the meantime</strong>,<br><br>Have a look at our &lt;&lt;Support site&gt;&gt; for: &nbsp;<ul><li>Tips for becoming a collaboration pro</li><li>Best information management practices for gcxchange</li><li>Terms and conditions for gcxchange</li></ul>We at gcxchange are always keeping accessibility in mind, guided by the Accessible Canada Act. &lt;&lt;Visit our support page&gt;&gt; to learn more. <br><br>Have a great day, <br><br>The gcxchange team&nbsp<br><hr><br><a id=\"fr\"><p>Bonjour de gcxchange! �&nbsp;<br><br>Nous vous remercions d�avoir demand� un espace d��quipe pour <strong>{displayName}</strong>.<br><br><strong>Quelle est la suite? </strong><br><br>Nous examinerons votre demande et, si tout est conforme, elle sera cr��e. � ce moment-l�, vous recevrez un deuxi�me courriel contenant les d�tails dont vous avez besoin pour commencer � collaborer. <br><br>Ce processus prend de 24 � 48?heures � compter du moment o� vous pr�sentez votre demande.<br><br><strong>Entretemps</strong>, <br><br>jetez un coup d��il � notre site de soutien. Vous y trouverez?:<ul><li>Des conseils pour devenir un professionnel de la collaboration </li><li>Les pratiques exemplaires de gestion de l�information pour gcxchange</li><li>Les modalit�s de gcxchange </li></ul>Chez gcxchange, nous gardons toujours � l�esprit l�accessibilit� en fonction de la Loi canadienne sur l�accessibilit�. Visitez notre page de soutien pour en savoir davantage. <br><br>Bonne journ�e!<br>L��quipe gcxchange�"
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
                        Subject = "Sorry, your gcxchange team space was not created /  Nous avons le regret de vous aviser que votre espace d��quipe gcxchange n�a pas �t� cr��",
                        Body = new ItemBody
                        {
                            ContentType = BodyType.Html,
                            Content = $"<a href=\"#fr\"><p>La version fran�aise suit</a><br>We have looked at your request for a gcxchange team space {displayName}.<br><br>Your requested team space has not been approved at this time. &nbsp;We were not able to create it for the following reason(s):<br><br>{comments}<br><br>We are here to help! &nbsp;If you are still interested in obtaining a team space or you think our decision has been made in error, please contact us via our &lt;&lt;Support site&gt;&gt;.<br><br>Come back soon to gcxchange to stay current, connect, and collaborate.<br><br>Have a great day,<br><br>The gcxchange team<br><br>We at gcxchange are always keeping accessibility in mind, guided by the Accessible Canada Act. &lt;&lt;Visit our support site&gt;&gt; to learn more.&nbsp;<br><br><hr><br><br>Nous avons examin� votre demande d�espace d��quipe gcxchange {displayName}.<br><br>L�espace d��quipe demand� n�a pas �t� approuv� pour l�instant. Nous n�avons pas pu le cr�er pour les raisons suivantes :<br><br> {comments} <br><br>Nous sommes l� pour vous aider! Si vous souhaitez toujours obtenir un espace d��quipe ou si vous pensez que notre d�cision est erron�e, veuillez communiquer avec nous par l�interm�diaire de notre site de soutien.<br><br>Revenez bient�t � gcxchange pour rester � jour, vous connecter et collaborer.<br><br>Bonne journ�e!<br><br>L��quipe gcxchange<br><br>Chez gcxchange, nous gardons toujours � l�esprit l�accessibilit� en fonction de la Loi canadienne sur l�accessibilit�. Visitez notre site de soutien pour en savoir davantage.&nbsp;</p>"
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
                            Subject = "Your gcxchange team space is ready / Votre espace d��quipe gcxchange est pr�t",
                            Body = new ItemBody
                            {
                                ContentType = BodyType.Html,
                                Content = $"<img src=\"https://tbssctdev.sharepoint.com/:i:/r/teams/scw/SiteAssets/Picture1.png?csf=1&web=1&e=dcaaBd\" alt=\"gcxchange\"><br><a href=\"#fr\">La version fran�aise suit</a><p><strong> Welcome to gcxchange collaboration! <br></strong><br>{requester} recently requested a team space &lt;&lt;{displayName}&gt;&gt; and it is now ready for take-off!<br><br><strong> How do I find and access my team space?<br></strong><br>The link to your SharePoint storefront is:<br><a href=\"{siteUrl}\">{siteUrl}</a><br>The left-hand menu has a link called <b>Open MSTeams</b> that will take you to your new Team space.<br><strong>If you have Microsoft Teams installed on your device:<br></strong><br>When you open Microsoft Teams you now have the ability to switch between your new gcxchange Teams space and your Department�s Teams space. &nbsp;To switch between gcxchange and your Department:<br><ol><li>Launch Microsoft Teams</li><li> Select the drop down in the Teams title bar(located immediately to the left of your Avatar)</li><li> Select gcxchange or your Department and Teams will switch over </li></ol>If you do not see the drop down, try to restart your device and launch MS Teams again. &nbsp;If the drop down is still not visible, please submit a service request to gcxchange via our support site.<br><br><strong>So, what's next?</strong><br><ol><li>Add members to your team. &nbsp;As an Owner, you can manage your team members using the Group and User Management program (located in the left menu of your SharePoint storefront).</li><li>Start adding great content to your storefront in SharePoint</li><li>Create channels and upload files to Microsoft Teams</li></ol>Visit our gcxchange support page &lt;&lt;hyperlink to Support main page&gt;&gt; for information on how to complete your next steps.<br><br><strong>Questions? Need additional support?<br></strong><br>Contact us on gcxchange support portal &lt;&lt;hyperlink to Support main page&gt;&gt;<br><br>We at gcxchange are always keeping accessibility in mind, guided by the Accessible Canada Act. Visit our support page (hyperlink) to learn more.<br><br><br><hr><strong><a id=\"fr\"></a>Bienvenue � la collaboration de gcxchange!<br></strong><br>{requester} a r�cemment demand� un espace {displayName}d��quipe et celui ci est maintenant pr�t pour l�aventure!<br><br><strong>Comment puis je trouver mon �quipe et avoir acc�s � celle ci?<br></strong><br>Voici le lien vers votre vitrine SharePoint :<br><a href=\"{siteUrl}\">{siteUrl}</a><br>Le menu de gauche contient un lien intitul� Open MSTeams qui vous am�nera � votre nouvel espace Teams.<br><strong>Quelle est la suite?</strong><br><ol><li>Ajoutez des membres � votre �quipe. En tant que propri�taire, vous pouvez g�rer les membres de votre �quipe au moyen du programme de gestion des groupes et des utilisateurs (qui se trouve dans le menu de gauche de votre vitrine SharePoint) :</li><li>Commencez � ajouter du contenu de qualit� dans SharePoint.</li><li>Cr�ez des canaux et t�l�chargez des fichiers dans Microsoft Teams</li></ol><strong>Si Microsoft Teams est install� sur votre appareil :<br></strong><br>Vous avez la possibilit� de passer de l�espace de vos nouvelles �quipes gcxchange � l�espace des �quipes de votre minist�re. Pour passer de gcxchange � votre minist�re :<br><ol><li>Lancez Microsoft Teams</li><li> S�lectionnez le menu d�roulant dans la barre de titre Teams (situ� imm�diatement � gauche de votre Avatar, pr�s du coins sup�rieur droit de la fen�tre Teams).</li><li>S�lectionnez gcxchange sinon votre minist�re et vos �quipes basculeront.</li></ol>Si vous ne voyez pas le menu d�roulant, essayez de red�marrer votre appareil et de relancer MS Teams. Si le menu d�roulant n�est toujours pas visible, veuillez soumettre une demande d�intervention � gcxchange par l�interm�diaire de notre site d�assistance.<br><br>Visitez notre page de soutien gcxchange pour savoir comment effectuer les prochaines �tapes.<br><br><strong>Des questions? Besoin de soutien suppl�mentaire?</strong><br><br>Communiquez avec nous sur le portail de soutien gcxchange<br><br>Chez gcxchange, nous gardons toujours � l�esprit l�accessibilit� en fonction de la Loi canadienne sur l�accessibilit�. Visitez notre page de soutien pour en savoir davantage.<br></p> "
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
