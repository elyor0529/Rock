using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text.RegularExpressions;
using Rock.Attribute;
using Rock.Data;
using Rock.Model;
using Rock.Transactions;
using Rock.Web.Cache;
using SendGrid;
using SendGrid.Helpers.Mail;

namespace Rock.Communication.Transport
{
    public class EmailSendResponse
    {
        public CommunicationRecipientStatus Status { get; set; }
        public string StatusNote { get; set; }
    }

    [TextField( "API Key", "The API Key provided by SendGrid.", true, "", "", 3, "APIKey" )]
    public class SendGridHttp : TransportComponent
    {
        public override bool Send( RockMessage rockMessage, int mediumEntityTypeId, Dictionary<string, string> mediumAttributes, out List<string> errorMessages )
        {
            errorMessages = new List<string>();

            var emailMessage = rockMessage as RockEmailMessage;
            if ( emailMessage == null )
            {
                return false;
            }

            var mergeFields = GetAllMergeFields( rockMessage.CurrentPerson, rockMessage.AdditionalMergeFields );
            var globalAttributes = GlobalAttributesCache.Get();
            var fromAddress = GetFromAddress( emailMessage, mergeFields, globalAttributes );

            if ( fromAddress.IsNullOrWhiteSpace() )
            {
                errorMessages.Add( "A From address was not provided." );
                return false;
            }

            var templateMailMessage = GetMailMessage( emailMessage, mergeFields, globalAttributes );
            var organizationEmail = globalAttributes.GetValue( "OrganizationEmail" );

            foreach ( var rockMessageRecipient in rockMessage.GetRecipients() )
            {
                try
                {
                    var recipientEmailMessage = GetRecipientRockEmailMessage( templateMailMessage, rockMessageRecipient, mergeFields, organizationEmail );

                    var result = SendEmail( recipientEmailMessage );
                    
                    // Create the communication record
                    if ( recipientEmailMessage.CreateCommunicationRecord )
                    {
                        var transaction = new SaveCommunicationTransaction( rockMessageRecipient, recipientEmailMessage.FromName, recipientEmailMessage.FromEmail, recipientEmailMessage.Subject, recipientEmailMessage.Message );
                        transaction.RecipientGuid = recipientEmailMessage.MessageMetaData["communication_recipient_guid"].AsGuidOrNull();
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

        public override void Send( Model.Communication communication, int mediumEntityTypeId, Dictionary<string, string> mediumAttributes )
        {
            using ( var communicationRockContext = new RockContext() )
            {
                // Requery the Communication
                communication = new CommunicationService( communicationRockContext )
                    .Queryable()
                    .Include( a => a.CreatedByPersonAlias.Person )
                    .Include( a => a.CommunicationTemplate )
                    .FirstOrDefault( c => c.Id == communication.Id );

                var isApprovedCommunication = communication != null && communication.Status == Model.CommunicationStatus.Approved;
                var isReadyToSend = communication != null &&
                                        ( !communication.FutureSendDateTime.HasValue
                                        || communication.FutureSendDateTime.Value.CompareTo( RockDateTime.Now ) <= 0 );

                if ( !isApprovedCommunication || !isReadyToSend )
                {
                    return;
                }

                // If there are no pending recipients than just exit the method
                var communicationRecipientService = new CommunicationRecipientService( communicationRockContext );

                var hasUnprocessedRecipients = communicationRecipientService
                    .Queryable()
                    .ByCommunicationId( communication.Id )
                    .ByStatus( CommunicationRecipientStatus.Pending )
                    .ByMediumEntityTypeId( mediumEntityTypeId )
                    .Any();

                if ( !hasUnprocessedRecipients )
                {
                    return;
                }

                var currentPerson = communication.CreatedByPersonAlias?.Person;
                var mergeFields = GetAllMergeFields(currentPerson, communication.AdditionalLavaFields);
                
                var globalAttributes = GlobalAttributesCache.Get();

                var templateEmailMessage = GetMailMessage( communication, mergeFields, globalAttributes );
                var organizationEmail = globalAttributes.GetValue( "OrganizationEmail" );

                var publicAppRoot = globalAttributes.GetValue( "PublicApplicationRoot" ).EnsureTrailingForwardslash();

                var cssInliningEnabled = communication.CommunicationTemplate?.CssInliningEnabled ?? false;

                var personEntityTypeId = EntityTypeCache.Get( "Rock.Model.Person" ).Id;
                var communicationEntityTypeId = EntityTypeCache.Get( "Rock.Model.Communication" ).Id;
                var communicationCategoryGuid = Rock.SystemGuid.Category.HISTORY_PERSON_COMMUNICATIONS.AsGuid();
                
                // Loop through recipients and send the email
                var recipientFound = true;
                while ( recipientFound )
                {
                    using ( var recipientRockContext = new RockContext() )
                    {

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
                            var recipientEmailMessage = GetRecipientRockEmailMessage( templateEmailMessage, communication, recipient, mergeObjects, organizationEmail, mediumAttributes );

                            var result = SendEmail( recipientEmailMessage );
                            
                            // Update recipient status and status note
                            recipient.Status = result.Status;
                            recipient.TransportEntityTypeName = this.GetType().FullName;

                            // Log it
                            try
                            {
                                var historyChangeList = new History.HistoryChangeList();
                                historyChangeList.AddChange(
                                    History.HistoryVerb.Sent,
                                    History.HistoryChangeType.Record,
                                    $"Communication" )
                                    .SetRelatedData( recipientEmailMessage.FromName, communicationEntityTypeId, communication.Id )
                                    .SetCaption( recipientEmailMessage.Subject );

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
        }

        private RockEmailMessage GetMailMessage( RockMessage rockMessage, Dictionary<string, object> mergeFields, GlobalAttributesCache globalAttributes )
        {
            var resultEmailMessage = new RockEmailMessage();

            var emailMessage = rockMessage as RockEmailMessage;
            if ( emailMessage == null )
            {
                return null;
            }

            resultEmailMessage.CurrentPerson = emailMessage.CurrentPerson;
            resultEmailMessage.EnabledLavaCommands = emailMessage.EnabledLavaCommands;

            var fromAddress = GetFromAddress( emailMessage, mergeFields, globalAttributes );
            var fromName = GetFromName( emailMessage, mergeFields, globalAttributes );

            if ( fromAddress.IsNullOrWhiteSpace() )
            {
                return null;
            }

            var fromMailAddress = new MailAddress( fromAddress, fromName );
            var organizationEmail = globalAttributes.GetValue( "OrganizationEmail" );

            // CC
            resultEmailMessage.CCEmails = emailMessage.CCEmails;

            // BCC
            resultEmailMessage.BCCEmails = emailMessage.BCCEmails;


            // Attachments
            resultEmailMessage.Attachments = emailMessage.Attachments;

            // Communication record for tracking opens & clicks
            resultEmailMessage.MessageMetaData = emailMessage.MessageMetaData;

            return resultEmailMessage;
        }

        private RockEmailMessage GetMailMessage( Model.Communication communication, Dictionary<string, object> mergeFields, GlobalAttributesCache globalAttributes )
        {
            var resultEmailMessage = new RockEmailMessage();

            var publicAppRoot = globalAttributes.GetValue( "PublicApplicationRoot" ).EnsureTrailingForwardslash();
            var cssInliningEnabled = communication.CommunicationTemplate?.CssInliningEnabled ?? false;

            resultEmailMessage.AppRoot = publicAppRoot;
            resultEmailMessage.CssInliningEnabled = cssInliningEnabled;
            resultEmailMessage.CurrentPerson = communication.CreatedByPersonAlias?.Person;
            resultEmailMessage.EnabledLavaCommands = communication.EnabledLavaCommands;
            resultEmailMessage.FromEmail = communication.FromEmail;
            resultEmailMessage.FromName = communication.FromName;

            var fromAddress = GetFromAddress( resultEmailMessage, mergeFields, globalAttributes );
            var fromName = GetFromName( resultEmailMessage, mergeFields, globalAttributes );

            if ( fromAddress.IsNullOrWhiteSpace() )
            {
                return null;
            }

            resultEmailMessage.FromEmail = fromAddress;
            resultEmailMessage.FromName = fromName;

            // Reply To
            var replyToEmail = "";
            if ( communication.ReplyToEmail.IsNotNullOrWhiteSpace() )
            {
                // Resolve any possible merge fields in the replyTo address
                replyToEmail = communication.ReplyToEmail.ResolveMergeFields( mergeFields, resultEmailMessage.CurrentPerson );
            }
            resultEmailMessage.ReplyToEmail = replyToEmail;

            // Attachments
            resultEmailMessage.Attachments = communication.GetAttachments( CommunicationType.Email ).Select( a => a.BinaryFile ).ToList();

            return resultEmailMessage;
        }

        private RockEmailMessage GetRecipientRockEmailMessage( RockEmailMessage emailMessage, RockMessageRecipient rockMessageRecipient, Dictionary<string, object> mergeFields, string organizationEmail )
        {
            var recipientEmail = new RockEmailMessage();
            recipientEmail.CurrentPerson = emailMessage.CurrentPerson;
            recipientEmail.EnabledLavaCommands = emailMessage.EnabledLavaCommands;
            // CC
            recipientEmail.CCEmails = emailMessage.CCEmails;

            // BCC
            recipientEmail.BCCEmails = emailMessage.BCCEmails;


            // Attachments
            recipientEmail.Attachments = emailMessage.Attachments;

            // Communication record for tracking opens & clicks
            recipientEmail.MessageMetaData = new Dictionary<string, string>( emailMessage.MessageMetaData );

            foreach ( var mergeField in mergeFields )
            {
                rockMessageRecipient.MergeFields.AddOrIgnore( mergeField.Key, mergeField.Value );
            }

            // To
            var toEmailAddress = new RockEmailMessageRecipient( null, null )
            {
                To = rockMessageRecipient
                    .To
                    .ResolveMergeFields( rockMessageRecipient.MergeFields, emailMessage.CurrentPerson, emailMessage.EnabledLavaCommands ),
                Name = rockMessageRecipient
                    .Name
                    .ResolveMergeFields( rockMessageRecipient.MergeFields, emailMessage.CurrentPerson, emailMessage.EnabledLavaCommands )
            };

            recipientEmail.SetRecipients( new List<RockEmailMessageRecipient> { toEmailAddress } );

            var fromMailAddress = new MailAddress( emailMessage.FromEmail, emailMessage.FromName );
            var checkResult = MailTransportHelper.CheckSafeSender( new List<string> { toEmailAddress.EmailAddress }, fromMailAddress, organizationEmail );

            // Reply To
            if ( checkResult.IsUnsafeDomain )
            {
                recipientEmail.ReplyToEmail = checkResult.SafeFromAddress.Address;
            }
            else if ( emailMessage.ReplyToEmail.IsNotNullOrWhiteSpace() )
            {
                recipientEmail.ReplyToEmail = emailMessage.ReplyToEmail.ResolveMergeFields( mergeFields, emailMessage.CurrentPerson, emailMessage.EnabledLavaCommands );
            }

            // From
            if ( checkResult.IsUnsafeDomain )
            {
                recipientEmail.FromName = checkResult.SafeFromAddress.DisplayName;
                recipientEmail.FromEmail = checkResult.SafeFromAddress.Address;
            }
            else
            {
                recipientEmail.FromName = fromMailAddress.DisplayName;
                recipientEmail.FromEmail = fromMailAddress.Address;
            }

            // Subject
            string subject = ResolveText( emailMessage.Subject, emailMessage.CurrentPerson, emailMessage.EnabledLavaCommands, rockMessageRecipient.MergeFields, emailMessage.AppRoot, emailMessage.ThemeRoot ).Left( 998 );
            recipientEmail.Subject = subject;

            // Body (HTML)
            string body = ResolveText( emailMessage.Message, emailMessage.CurrentPerson, emailMessage.EnabledLavaCommands, rockMessageRecipient.MergeFields, emailMessage.AppRoot, emailMessage.ThemeRoot );
            body = Regex.Replace( body, @"\[\[\s*UnsubscribeOption\s*\]\]", string.Empty );
            recipientEmail.Message = body;


            Guid? recipientGuid = null;
            recipientEmail.CreateCommunicationRecord = emailMessage.CreateCommunicationRecord;
            if ( emailMessage.CreateCommunicationRecord )
            {
                recipientGuid = Guid.NewGuid();
                recipientEmail.MessageMetaData["communication_recipient_guid"] = recipientGuid.ToString();
            }

            return recipientEmail;
        }

        public RockEmailMessage GetRecipientRockEmailMessage( RockEmailMessage emailMessage, Model.Communication communication, CommunicationRecipient communicationRecipient, Dictionary<string, object> mergeFields, string organizationEmail, Dictionary<string, string> mediumAttributes )
        {
            var recipientEmail = new RockEmailMessage();
            recipientEmail.CurrentPerson = emailMessage.CurrentPerson;
            recipientEmail.EnabledLavaCommands = emailMessage.EnabledLavaCommands;

            // CC
            if ( communication.CCEmails.IsNotNullOrWhiteSpace() )
            {
                string[] ccRecipients = communication
                    .CCEmails
                    .ResolveMergeFields( mergeFields, emailMessage.CurrentPerson )
                    .Replace( ";", "," )
                    .Split( ',' );

                foreach ( var ccRecipient in ccRecipients )
                {
                    recipientEmail.CCEmails.Add( ccRecipient );
                }
            }

            // BCC
            if ( communication.BCCEmails.IsNotNullOrWhiteSpace() )
            {
                string[] bccRecipients = communication
                    .BCCEmails
                    .ResolveMergeFields( mergeFields, emailMessage.CurrentPerson )
                    .Replace( ";", "," )
                    .Split( ',' );

                foreach ( var bccRecipient in bccRecipients )
                {
                    recipientEmail.BCCEmails.Add( bccRecipient );
                }
            }

            // Attachments
            recipientEmail.Attachments = emailMessage.Attachments;

            // Communication record for tracking opens & clicks
            recipientEmail.MessageMetaData = new Dictionary<string, string>( emailMessage.MessageMetaData );

            // To
            var toEmailAddress = new RockEmailMessageRecipient( null, null )
            {
                To = communicationRecipient.PersonAlias.Person.Email,
                Name = communicationRecipient.PersonAlias.Person.FullName
            };

            recipientEmail.SetRecipients( new List<RockEmailMessageRecipient> { toEmailAddress } );

            var fromMailAddress = new MailAddress( emailMessage.FromEmail, emailMessage.FromName );
            var checkResult = MailTransportHelper.CheckSafeSender( new List<string> { toEmailAddress.EmailAddress }, fromMailAddress, organizationEmail );

            // Reply To
            if ( checkResult.IsUnsafeDomain )
            {
                recipientEmail.ReplyToEmail = checkResult.SafeFromAddress.Address;
            }
            else if ( emailMessage.ReplyToEmail.IsNotNullOrWhiteSpace() )
            {
                recipientEmail.ReplyToEmail = emailMessage.ReplyToEmail.ResolveMergeFields( mergeFields, emailMessage.CurrentPerson, emailMessage.EnabledLavaCommands );
            }

            // From
            if ( checkResult.IsUnsafeDomain )
            {
                recipientEmail.FromName = checkResult.SafeFromAddress.DisplayName;
                recipientEmail.FromEmail = checkResult.SafeFromAddress.Address;
            }
            else
            {
                recipientEmail.FromName = fromMailAddress.DisplayName;
                recipientEmail.FromEmail = fromMailAddress.Address;
            }

            // Subject
            var subject = ResolveText( communication.Subject, emailMessage.CurrentPerson, communication.EnabledLavaCommands, mergeFields, emailMessage.AppRoot );
            recipientEmail.Subject = subject;

            // Body Plain Text
            if ( mediumAttributes.ContainsKey( "DefaultPlainText" ) )
            {
                var plainText = ResolveText( mediumAttributes["DefaultPlainText"],
                    emailMessage.CurrentPerson,
                    communication.EnabledLavaCommands,
                    mergeFields,
                    emailMessage.AppRoot );

                if ( !string.IsNullOrWhiteSpace( plainText ) )
                {
                    recipientEmail.PlainTextMessage = plainText;
                }
            }

            // Body (HTML)
            string htmlBody = communication.Message;

            // Get the unsubscribe content and add a merge field for it
            if ( communication.IsBulkCommunication && mediumAttributes.ContainsKey( "UnsubscribeHTML" ) )
            {
                string unsubscribeHtml = ResolveText( mediumAttributes["UnsubscribeHTML"],
                    emailMessage.CurrentPerson,
                    communication.EnabledLavaCommands,
                    mergeFields,
                    emailMessage.AppRoot );

                mergeFields.AddOrReplace( "UnsubscribeOption", unsubscribeHtml );

                htmlBody = ResolveText( htmlBody, emailMessage.CurrentPerson, communication.EnabledLavaCommands, mergeFields, emailMessage.AppRoot );

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
                htmlBody = ResolveText( htmlBody, emailMessage.CurrentPerson, communication.EnabledLavaCommands, mergeFields, emailMessage.AppRoot );
                htmlBody = Regex.Replace( htmlBody, @"\[\[\s*UnsubscribeOption\s*\]\]", string.Empty );
            }

            if ( !string.IsNullOrWhiteSpace( htmlBody ) )
            {
                if ( emailMessage.CssInliningEnabled )
                {
                    // move styles inline to help it be compatible with more email clients
                    htmlBody = htmlBody.ConvertHtmlStylesToInlineAttributes();
                }

                // add the main Html content to the email
                recipientEmail.Message = htmlBody;
            }

            recipientEmail.MessageMetaData["communication_recipient_guid"] = communicationRecipient.Guid.ToString();

            return recipientEmail;
        }

        private SendGridMessage GetSendGridMessageFromRockEmailMessage( RockEmailMessage rockEmailMessage )
        {
            var sendGridMessage = new SendGridMessage();

            // To
            rockEmailMessage.GetRecipients().ForEach( r => sendGridMessage.AddTo( r.To, r.Name ) );
            sendGridMessage.ReplyTo = new EmailAddress( rockEmailMessage.ReplyToEmail );
            sendGridMessage.From = new EmailAddress( rockEmailMessage.FromEmail, rockEmailMessage.FromName );

            // CC
            var ccEmailAddresses = rockEmailMessage
                                    .CCEmails
                                    .Where( e => e != string.Empty )
                                    .Select( cc => new EmailAddress { Email = cc } )
                                    .ToList();
            sendGridMessage.AddCcs( ccEmailAddresses );

            // BCC
            var bccEmailAddresses = rockEmailMessage
                .BCCEmails
                .Where( e => e != string.Empty )
                .Select( cc => new EmailAddress { Email = cc } )
                .ToList();
            sendGridMessage.AddBccs( bccEmailAddresses );

            // Subject
            sendGridMessage.Subject = rockEmailMessage.Subject;

            // Body (plain text)
            sendGridMessage.PlainTextContent = rockEmailMessage.PlainTextMessage;

            // Body (html)
            sendGridMessage.HtmlContent = rockEmailMessage.Message;

            // Communication record for tracking opens & clicks
            sendGridMessage.CustomArgs = rockEmailMessage.MessageMetaData;

            sendGridMessage.TrackingSettings.OpenTracking.Enable = CanTrackOpens;
            sendGridMessage.TrackingSettings.ClickTracking.Enable = CanTrackOpens;

            // Attachments
            if ( rockEmailMessage.Attachments.Any() )
            {
                using ( var rockContext = new RockContext() )
                {
                    var binaryFileService = new BinaryFileService( rockContext );
                    foreach ( var binaryFileId in rockEmailMessage.Attachments.Where( a => a != null ).Select( a => a.Id ) )
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

            return sendGridMessage;
        }

        private EmailSendResponse SendEmail(RockEmailMessage rockEmailMessage )
        {
            var client = new SendGridClient( GetAttributeValue( "APIKey" ) );
            var sendGridMessage = GetSendGridMessageFromRockEmailMessage( rockEmailMessage );

            // Send it
            var response = client.SendEmailAsync( sendGridMessage ).GetAwaiter().GetResult();
            return new EmailSendResponse
            {
                Status = response.StatusCode == HttpStatusCode.OK ? CommunicationRecipientStatus.Delivered : CommunicationRecipientStatus.Failed
            };
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

        private Dictionary<string, object> GetAllMergeFields( Person currentPerson, Dictionary<string, object> additionalMergeFields )
        {
            // Common Merge Field
            var mergeFields = Lava.LavaHelper.GetCommonMergeFields( null, currentPerson );
            foreach ( var mergeField in additionalMergeFields )
            {
                mergeFields.AddOrReplace( mergeField.Key, mergeField.Value );
            }

            return mergeFields;
        }
    }
}
