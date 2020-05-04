using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using Rock.Model;
using Rock.Web.Cache;
using RestSharp;
using RestSharp.Authenticators;
using SendGrid.Helpers.Mail;
using System.Text.RegularExpressions;
using Rock.Data;
using System.IO;
using Rock.Transactions;
using SendGrid;
using Rock.Attribute;

namespace Rock.Communication.Transport
{
    [TextField( "API Key", "The API Key provided by SendGrid.", true, "", "", 3, "APIKey" )]
    public class SendGridSmtp : TransportComponent
    {
        public override bool Send( RockMessage rockMessage, int mediumEntityTypeId, Dictionary<string, string> mediumAttributes, out List<string> errorMessages )
        {
            errorMessages = new List<string>();

            var emailMessage = rockMessage as RockEmailMessage;
            if ( emailMessage == null )
            {
                return false;
            }

            var mergeFields = GetAllMergeFields( rockMessage );

            var globalAttributes = GlobalAttributesCache.Get();

            var fromAddress = GetFromAddress( emailMessage, mergeFields, globalAttributes );
            var fromName = GetFromName( emailMessage, mergeFields, globalAttributes );

            if ( fromAddress.IsNullOrWhiteSpace() )
            {
                errorMessages.Add( "A From address was not provided." );
                return false;
            }

            var fromMailAddress = new MailAddress( fromAddress, fromName );
            var emailRecipients = rockMessage.GetRecipients();
            var organizationEmail = globalAttributes.GetValue( "OrganizationEmail" );
            var client = new SendGridClient( GetAttributeValue( "APIKey" ) );

            foreach ( var rockMessageRecipient in rockMessage.GetRecipients() )
            {
                try
                {
                    foreach ( var mergeField in mergeFields )
                    {
                        rockMessageRecipient.MergeFields.AddOrIgnore( mergeField.Key, mergeField.Value );
                    }

                    var sendGridMessage = new SendGridMessage();

                    // To
                    var toEmailAddress = new EmailAddress
                    {
                        Email = rockMessageRecipient
                            .To
                            .ResolveMergeFields( rockMessageRecipient.MergeFields, emailMessage.CurrentPerson, emailMessage.EnabledLavaCommands ),
                        Name = rockMessageRecipient
                            .Name
                            .ResolveMergeFields( rockMessageRecipient.MergeFields, emailMessage.CurrentPerson, emailMessage.EnabledLavaCommands )
                    };

                    sendGridMessage.AddTo( toEmailAddress );

                    var checkResult = MailTransportHelper.CheckSafeSender( new List<string> { toEmailAddress.Email }, fromMailAddress, organizationEmail );

                    // Reply To
                    if ( checkResult.IsUnsafeDomain )
                    {
                        sendGridMessage.ReplyTo = CreateEmailAddressFromMailAddress( checkResult.SafeFromAddress );
                    }
                    else if ( emailMessage.ReplyToEmail.IsNotNullOrWhiteSpace() )
                    {
                        sendGridMessage.ReplyTo = new EmailAddress
                        {
                            Email = emailMessage.ReplyToEmail.ResolveMergeFields( mergeFields, emailMessage.CurrentPerson, emailMessage.EnabledLavaCommands )
                        };
                    }

                    // From
                    if ( checkResult.IsUnsafeDomain )
                    {
                        sendGridMessage.From = CreateEmailAddressFromMailAddress( checkResult.SafeFromAddress );
                    }
                    else
                    {
                        sendGridMessage.From = CreateEmailAddressFromMailAddress( fromMailAddress );
                    }

                    // CC
                    var ccEmailAddresses = emailMessage
                                            .CCEmails
                                            .Where( e => e != string.Empty )
                                            .Select( cc => new EmailAddress { Email = cc } )
                                            .ToList();
                    sendGridMessage.AddCcs( ccEmailAddresses );

                    // BCC
                    var bccEmailAddresses = emailMessage
                        .BCCEmails
                        .Where( e => e != string.Empty )
                        .Select( cc => new EmailAddress { Email = cc } )
                        .ToList();
                    sendGridMessage.AddBccs( bccEmailAddresses );

                    // Subject
                    string subject = ResolveText( emailMessage.Subject, emailMessage.CurrentPerson, emailMessage.EnabledLavaCommands, rockMessageRecipient.MergeFields, emailMessage.AppRoot, emailMessage.ThemeRoot ).Left( 998 );
                    sendGridMessage.Subject = subject;

                    // Body (html)
                    string body = ResolveText( emailMessage.Message, emailMessage.CurrentPerson, emailMessage.EnabledLavaCommands, rockMessageRecipient.MergeFields, emailMessage.AppRoot, emailMessage.ThemeRoot );
                    body = Regex.Replace( body, @"\[\[\s*UnsubscribeOption\s*\]\]", string.Empty );
                    sendGridMessage.HtmlContent = body;

                    // Communication record for tracking opens & clicks
                    var metaData = new Dictionary<string, string>( emailMessage.MessageMetaData );
                    Guid? recipientGuid = null;

                    if ( emailMessage.CreateCommunicationRecord )
                    {
                        recipientGuid = Guid.NewGuid();
                        metaData.Add( "communication_recipient_guid", recipientGuid.Value.ToString() );
                    }

                    sendGridMessage.CustomArgs = metaData;

                    sendGridMessage.TrackingSettings.OpenTracking.Enable = CanTrackOpens;
                    sendGridMessage.TrackingSettings.ClickTracking.Enable = CanTrackOpens;

                    // Attachments
                    if ( emailMessage.Attachments.Any() )
                    {
                        using ( var rockContext = new RockContext() )
                        {
                            var binaryFileService = new BinaryFileService( rockContext );
                            foreach ( var binaryFileId in emailMessage.Attachments.Where( a => a != null ).Select( a => a.Id ) )
                            {
                                var attachment = binaryFileService.Get( binaryFileId );
                                if ( attachment != null )
                                {
                                    MemoryStream ms = new MemoryStream();
                                    attachment.ContentStream.CopyTo( ms );
                                    sendGridMessage.AddAttachment( attachment.FileName, Convert.ToBase64String( ms.ToArray() ) );
                                }
                            }
                        }
                    }

                    // Send it
                    var response = client.SendEmailAsync( sendGridMessage );

                    // Create the communication record
                    if ( emailMessage.CreateCommunicationRecord )
                    {
                        var transaction = new SaveCommunicationTransaction( rockMessageRecipient, emailMessage.FromName, emailMessage.FromEmail, subject, body );
                        transaction.RecipientGuid = recipientGuid;
                        RockQueue.TransactionQueue.Enqueue( transaction );
                    }
                }
                catch ( Exception ex )
                {
                    errorMessages.Add( ex.Message );
                    ExceptionLogService.LogException( ex );
                }
            }

            return !errorMessages.Any();
        }

        /// <summary>
        /// Sends the specified communication.
        /// </summary>
        /// <param name="communication">The communication.</param>
        /// <param name="mediumEntityTypeId">The medium entity type identifier.</param>
        /// <param name="mediumAttributes">The medium attributes.</param>
        public override void Send( Model.Communication communication, int mediumEntityTypeId, Dictionary<string, string> mediumAttributes )
        {
            using ( var communicationRockContext = new RockContext() )
            {
                // Requery the Communication
                communication = new CommunicationService( communicationRockContext )
                    .Queryable().Include( a => a.CreatedByPersonAlias.Person ).Include( a => a.CommunicationTemplate )
                    .FirstOrDefault( c => c.Id == communication.Id );

                // If there are no pending recipients than just exit the method
                if ( communication != null &&
                    communication.Status == Model.CommunicationStatus.Approved &&
                    ( !communication.FutureSendDateTime.HasValue || communication.FutureSendDateTime.Value.CompareTo( RockDateTime.Now ) <= 0 ) )
                {
                    var qryRecipients = new CommunicationRecipientService( communicationRockContext ).Queryable();
                    if ( !qryRecipients
                        .Where( r =>
                            r.CommunicationId == communication.Id &&
                            r.Status == Model.CommunicationRecipientStatus.Pending &&
                            r.MediumEntityTypeId.HasValue &&
                            r.MediumEntityTypeId.Value == mediumEntityTypeId )
                        .Any() )
                    {
                        return;
                    }
                }

                var currentPerson = communication.CreatedByPersonAlias?.Person;
                var globalAttributes = GlobalAttributesCache.Get();
                string publicAppRoot = globalAttributes.GetValue( "PublicApplicationRoot" ).EnsureTrailingForwardslash();
                var mergeFields = Lava.LavaHelper.GetCommonMergeFields( null, currentPerson );
                var cssInliningEnabled = communication.CommunicationTemplate?.CssInliningEnabled ?? false;

                string fromAddress = string.IsNullOrWhiteSpace( communication.FromEmail ) ? globalAttributes.GetValue( "OrganizationEmail" ) : communication.FromEmail;
                string fromName = string.IsNullOrWhiteSpace( communication.FromName ) ? globalAttributes.GetValue( "OrganizationName" ) : communication.FromName;

                // Resolve any possible merge fields in the from address
                fromAddress = fromAddress.ResolveMergeFields( mergeFields, currentPerson, communication.EnabledLavaCommands );
                fromName = fromName.ResolveMergeFields( mergeFields, currentPerson, communication.EnabledLavaCommands );
                Parameter replyTo = new Parameter();

                // Reply To
                if ( communication.ReplyToEmail.IsNotNullOrWhiteSpace() )
                {
                    // Resolve any possible merge fields in the replyTo address
                    replyTo.Name = "h:Reply-To";
                    replyTo.Type = ParameterType.GetOrPost;
                    replyTo.Value = communication.ReplyToEmail.ResolveMergeFields( mergeFields, currentPerson );
                }

                var personEntityTypeId = EntityTypeCache.Get( "Rock.Model.Person" ).Id;
                var communicationEntityTypeId = EntityTypeCache.Get( "Rock.Model.Communication" ).Id;
                var communicationCategoryGuid = Rock.SystemGuid.Category.HISTORY_PERSON_COMMUNICATIONS.AsGuid();

                RestRequest restRequest = null;

                // Loop through recipients and send the email
                bool recipientFound = true;
                while ( recipientFound )
                {
                    var recipientRockContext = new RockContext();
                    var recipient = Model.Communication.GetNextPending( communication.Id, mediumEntityTypeId, recipientRockContext );

                    // This means we are done, break the loop
                    if ( recipient == null )
                    {
                        recipientFound = false;
                        break;
                    }

                    // Not valid save the obj with the status messages then go to the next one
                    if ( !ValidRecipient( recipient, communication.IsBulkCommunication ) )
                    {
                        recipientRockContext.SaveChanges();
                        continue;
                    }

                    try
                    {
                        // Create merge field dictionary
                        var mergeObjects = recipient.CommunicationMergeValues( mergeFields );

                        // Create the request obj
                        restRequest = new RestRequest( GetAttributeValue( "Resource" ), Method.POST );
                        restRequest.AddParameter( "domian", GetAttributeValue( "Domain" ), ParameterType.UrlSegment );

                        // ReplyTo
                        if ( communication.ReplyToEmail.IsNotNullOrWhiteSpace() )
                        {
                            restRequest.AddParameter( replyTo );
                        }

                        // From
                        var fromEmailAddress = new MailAddress( fromAddress, fromName );
                        restRequest.AddParameter( "from", fromEmailAddress.ToString() );

                        // To
                        restRequest.AddParameter( "to", new MailAddress( recipient.PersonAlias.Person.Email, recipient.PersonAlias.Person.FullName ).ToString() );

                        // Safe sender checks
                        HandleUnsafeSender( restRequest, fromEmailAddress, globalAttributes.GetValue( "OrganizationEmail" ) );

                        // CC
                        if ( communication.CCEmails.IsNotNullOrWhiteSpace() )
                        {
                            string[] ccRecipients = communication.CCEmails.ResolveMergeFields( mergeObjects, currentPerson ).Replace( ";", "," ).Split( ',' );
                            foreach ( var ccRecipient in ccRecipients )
                            {
                                restRequest.AddParameter( "cc", ccRecipient );
                            }
                        }

                        // BCC
                        if ( communication.BCCEmails.IsNotNullOrWhiteSpace() )
                        {
                            string[] bccRecipients = communication.BCCEmails.ResolveMergeFields( mergeObjects, currentPerson ).Replace( ";", "," ).Split( ',' );
                            foreach ( var bccRecipient in bccRecipients )
                            {
                                restRequest.AddParameter( "bcc", bccRecipient );
                            }
                        }

                        // Subject
                        string subject = ResolveText( communication.Subject, currentPerson, communication.EnabledLavaCommands, mergeObjects, publicAppRoot );
                        restRequest.AddParameter( "subject", subject );

                        // Body Plain Text
                        if ( mediumAttributes.ContainsKey( "DefaultPlainText" ) )
                        {
                            string plainText = ResolveText( mediumAttributes["DefaultPlainText"], currentPerson, communication.EnabledLavaCommands, mergeObjects, publicAppRoot );
                            if ( !string.IsNullOrWhiteSpace( plainText ) )
                            {
                                AlternateView plainTextView = AlternateView.CreateAlternateViewFromString( plainText, new ContentType( MediaTypeNames.Text.Plain ) );
                                restRequest.AddParameter( "text", plainTextView );
                            }
                        }

                        // Body HTML
                        string htmlBody = communication.Message;

                        // Get the unsubscribe content and add a merge field for it
                        if ( communication.IsBulkCommunication && mediumAttributes.ContainsKey( "UnsubscribeHTML" ) )
                        {
                            string unsubscribeHtml = ResolveText( mediumAttributes["UnsubscribeHTML"], currentPerson, communication.EnabledLavaCommands, mergeObjects, publicAppRoot );
                            mergeObjects.AddOrReplace( "UnsubscribeOption", unsubscribeHtml );
                            htmlBody = ResolveText( htmlBody, currentPerson, communication.EnabledLavaCommands, mergeObjects, publicAppRoot );

                            // Resolve special syntax needed if option was included in global attribute
                            if ( Regex.IsMatch( htmlBody, @"\[\[\s*UnsubscribeOption\s*\]\]" ) )
                            {
                                htmlBody = Regex.Replace( htmlBody, @"\[\[\s*UnsubscribeOption\s*\]\]", unsubscribeHtml );
                            }

                            // Add the unsubscribe option at end if it wasn't included in content
                            if ( !htmlBody.Contains( unsubscribeHtml ) )
                            {
                                htmlBody += unsubscribeHtml;
                            }
                        }
                        else
                        {
                            htmlBody = ResolveText( htmlBody, currentPerson, communication.EnabledLavaCommands, mergeObjects, publicAppRoot );
                            htmlBody = Regex.Replace( htmlBody, @"\[\[\s*UnsubscribeOption\s*\]\]", string.Empty );
                        }

                        if ( !string.IsNullOrWhiteSpace( htmlBody ) )
                        {
                            if ( cssInliningEnabled )
                            {
                                // move styles inline to help it be compatible with more email clients
                                htmlBody = htmlBody.ConvertHtmlStylesToInlineAttributes();
                            }

                            // add the main Html content to the email
                            restRequest.AddParameter( "html", htmlBody );
                        }

                        // Headers
                        AddAdditionalHeaders( restRequest, new Dictionary<string, string>() { { "communication_recipient_guid", recipient.Guid.ToString() } } );

                        // Attachments
                        foreach ( var attachment in communication.GetAttachments( CommunicationType.Email ).Select( a => a.BinaryFile ) )
                        {
                            MemoryStream ms = new MemoryStream();
                            attachment.ContentStream.CopyTo( ms );
                            restRequest.AddFile( "attachment", ms.ToArray(), attachment.FileName );
                        }

                        // Send the email
                        // Send it
                        RestClient restClient = new RestClient
                        {
                            BaseUrl = new Uri( GetAttributeValue( "BaseURL" ) ),
                            Authenticator = new HttpBasicAuthenticator( "api", GetAttributeValue( "APIKey" ) )
                        };

                        // Call the API and get the response
                        Response = restClient.Execute( restRequest );

                        // Update recipient status and status note
                        recipient.Status = Response.StatusCode == HttpStatusCode.OK ? CommunicationRecipientStatus.Delivered : CommunicationRecipientStatus.Failed;
                        recipient.StatusNote = Response.StatusDescription;
                        recipient.TransportEntityTypeName = this.GetType().FullName;

                        // Log it
                        try
                        {
                            var historyChangeList = new History.HistoryChangeList();
                            historyChangeList.AddChange(
                                History.HistoryVerb.Sent,
                                History.HistoryChangeType.Record,
                                $"Communication" )
                                .SetRelatedData( fromName, communicationEntityTypeId, communication.Id )
                                .SetCaption( subject );

                            HistoryService.SaveChanges( recipientRockContext, typeof( Rock.Model.Person ), communicationCategoryGuid, recipient.PersonAlias.PersonId, historyChangeList, false, communication.SenderPersonAliasId );
                        }
                        catch ( Exception ex )
                        {
                            ExceptionLogService.LogException( ex, null );
                        }
                    }
                    catch ( Exception ex )
                    {
                        ExceptionLogService.LogException( ex );
                        recipient.Status = CommunicationRecipientStatus.Failed;
                        recipient.StatusNote = "Exception: " + ex.Messages().AsDelimited( " => " );
                    }

                    recipientRockContext.SaveChanges();
                }
            }
        }

        private string GetFromName( RockEmailMessage emailMessage, Dictionary<string, object> mergeFields, GlobalAttributesCache globalAttributes )
        {
            string fromName = emailMessage.FromName.ResolveMergeFields( mergeFields, emailMessage.CurrentPerson, emailMessage.EnabledLavaCommands );
            fromName = fromName.IsNullOrWhiteSpace() ? globalAttributes.GetValue( "OrganizationName" ) : fromName;
            return fromName;
        }

        private string GetFromAddress( RockEmailMessage emailMessage, Dictionary<string, object> mergeFields, GlobalAttributesCache globalAttributes )
        {

            // Resolve any possible merge fields in the from address
            string fromAddress = emailMessage.FromEmail.ResolveMergeFields( mergeFields, emailMessage.CurrentPerson, emailMessage.EnabledLavaCommands );
            fromAddress = fromAddress.IsNullOrWhiteSpace() ? globalAttributes.GetValue( "OrganizationEmail" ) : fromAddress;
            return fromAddress;
        }

        private Dictionary<string, object> GetAllMergeFields( RockMessage rockMessage )
        {

            // Common Merge Field
            var mergeFields = Lava.LavaHelper.GetCommonMergeFields( null, rockMessage.CurrentPerson );
            foreach ( var mergeField in rockMessage.AdditionalMergeFields )
            {
                mergeFields.AddOrReplace( mergeField.Key, mergeField.Value );
            }

            return mergeFields;
        }

        private EmailAddress CreateEmailAddressFromMailAddress( MailAddress address )
        {
            return new EmailAddress
            {
                Email = address.Address,
                Name = address.DisplayName
            };
        }

        public override void Send( Model.Communication communication )
        {
            throw new NotImplementedException();
        }

        public override void Send( SystemEmail template, List<RecipientData> recipients, string appRoot, string themeRoot )
        {
            throw new NotImplementedException();
        }

        public override void Send( Dictionary<string, string> mediumData, List<string> recipients, string appRoot, string themeRoot )
        {
            throw new NotImplementedException();
        }

        public override void Send( List<string> recipients, string from, string subject, string body, string appRoot = null, string themeRoot = null )
        {
            throw new NotImplementedException();
        }

        public override void Send( List<string> recipients, string from, string subject, string body, string appRoot = null, string themeRoot = null, List<System.Net.Mail.Attachment> attachments = null )
        {
            throw new NotImplementedException();
        }

        public override void Send( List<string> recipients, string from, string fromName, string subject, string body, string appRoot = null, string themeRoot = null, List<System.Net.Mail.Attachment> attachments = null )
        {
            throw new NotImplementedException();
        }
    }

    internal class SendGridPersonalizations
    {
        public Email
    }
}
