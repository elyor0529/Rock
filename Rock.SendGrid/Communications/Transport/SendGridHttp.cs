using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Composition;
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
    [Description( "Sends a communication through SendGrid's HTTP API" )]
    [Export( typeof( TransportComponent ) )]
    [ExportMetadata( "ComponentName", "SendGrid HTTP" )]

    [TextField( "Base URL",
        Description = "The API URL provided by SendGrid, keep the default in most cases.",
        IsRequired = true,
        DefaultValue = @"https://api.sendgrid.com",
        Order = 0,
        Key = AttributeKey.BaseUrl )]
    [TextField( "API Key",
        Description = "The API Key provided by SendGrid.",
        IsRequired = true,
        Order = 3,
        Key = AttributeKey.ApiKey )]
    [BooleanField( "Track Opens",
        Description = "Allow SendGrid to track opens, clicks, and unsubscribes.",
        DefaultValue = "true",
        Order = 4,
        Key = AttributeKey.TrackOpens )]
    public class SendGridHttp : EmailTransportComponent
    {
        public class AttributeKey
        {
            public const string TrackOpens = "TrackOpens";
            public const string ApiKey = "APIKey";
            public const string BaseUrl = "BaseURL";
        }

        /// <summary>
        /// Gets a value indicating whether transport has ability to track recipients opening the communication.
        /// Mailgun automatically trackes opens, clicks, and unsubscribes. Use this to override domain setting.
        /// </summary>
        /// <value>
        /// <c>true</c> if transport can track opens; otherwise, <c>false</c>.
        /// </value>
        public override bool CanTrackOpens
        {
            get { return GetAttributeValue( AttributeKey.TrackOpens ).AsBoolean( true ); }
        }

        private SendGridMessage GetSendGridMessageFromRockEmailMessage( RockEmailMessage rockEmailMessage )
        {
            var sendGridMessage = new SendGridMessage();

            // To
            rockEmailMessage.GetRecipients().ForEach( r => sendGridMessage.AddTo( r.To, r.Name ) );

            if ( rockEmailMessage.ReplyToEmail.IsNotNullOrWhiteSpace() )
            {
                sendGridMessage.ReplyTo = new EmailAddress( rockEmailMessage.ReplyToEmail );
            }

            sendGridMessage.From = new EmailAddress( rockEmailMessage.FromEmail, rockEmailMessage.FromName );

            // CC
            var ccEmailAddresses = rockEmailMessage
                                    .CCEmails
                                    .Where( e => e != string.Empty )
                                    .Select( cc => new EmailAddress { Email = cc } )
                                    .ToList();
            if ( ccEmailAddresses.Count > 0 )
            {
                sendGridMessage.AddCcs( ccEmailAddresses );
            }

            // BCC
            var bccEmailAddresses = rockEmailMessage
                .BCCEmails
                .Where( e => e != string.Empty )
                .Select( cc => new EmailAddress { Email = cc } )
                .ToList();
            if ( bccEmailAddresses.Count > 0 )
            {
                sendGridMessage.AddBccs( bccEmailAddresses );
            }

            // Subject
            sendGridMessage.Subject = rockEmailMessage.Subject;

            // Body (plain text)
            sendGridMessage.PlainTextContent = rockEmailMessage.PlainTextMessage;

            // Body (html)
            sendGridMessage.HtmlContent = rockEmailMessage.Message;

            // Communication record for tracking opens & clicks
            sendGridMessage.CustomArgs = rockEmailMessage.MessageMetaData;

            if ( CanTrackOpens )
            {
                sendGridMessage.TrackingSettings = new TrackingSettings
                {
                    ClickTracking = new ClickTracking { Enable = true },
                    OpenTracking = new OpenTracking { Enable = true }
                };
            }

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

        protected override EmailSendResponse SendEmail( RockEmailMessage rockEmailMessage )
        {
            var client = new SendGridClient( GetAttributeValue( AttributeKey.ApiKey ), host: GetAttributeValue( AttributeKey.BaseUrl ) );
            var sendGridMessage = GetSendGridMessageFromRockEmailMessage( rockEmailMessage );

            // Send it
            var response = client.SendEmailAsync( sendGridMessage ).GetAwaiter().GetResult();
            return new EmailSendResponse
            {
                Status = response.StatusCode == HttpStatusCode.Accepted ? CommunicationRecipientStatus.Delivered : CommunicationRecipientStatus.Failed
            };
        }
    }
}
