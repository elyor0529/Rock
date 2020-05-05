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
    [TextField( "API Key", "The API Key provided by SendGrid.", true, "", "", 3, "APIKey" )]
    public class SendGridHttp : EmailTransportComponent
    {
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
                foreach ( var attachment in rockEmailMessage.Attachments )
                {
                    if ( attachment != null )
                    {
                        MemoryStream ms = new MemoryStream();
                        attachment.ContentStream.CopyTo( ms );
                        sendGridMessage.AddAttachment( attachment.FileName, Convert.ToBase64String( ms.ToArray() ) );
                    }
                }
            }

            return sendGridMessage;
        }

        protected override EmailSendResponse SendEmail(RockEmailMessage rockEmailMessage )
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
    }
}
