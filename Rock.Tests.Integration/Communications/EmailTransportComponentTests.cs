using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Moq.Protected;
using Rock.Communication;
using Rock.Communication.Transport;
using Rock.Data;
using Rock.Model;
using Rock.Tests.Shared;
using Rock.Web.Cache;

namespace Rock.Tests.Integration.Communications
{
    [TestClass]
    public class EmailTransportComponentTests
    {
        #region Rock Email Message Test
        [TestMethod]
        public void SendRockMessageShouldReplaceUnsafeFromWithOrganizationEmail()
        {
            var expectedFromEmail = "test@org.com";
            var expectedFromName = "Test Name";

            var globalAttributes = GlobalAttributesCache.Get();
            globalAttributes.SetValue( "OrganizationEmail", expectedFromEmail, false, null );

            var actualEmail = new RockEmailMessage()
            {
                FromEmail = "info@test.com",
                FromName = expectedFromName
            };

            actualEmail.AddRecipient( new RockEmailMessageRecipient( new Person
            {
                Email = "test@test.com",
                FirstName = "Test",
                LastName = "User"
            }, new Dictionary<string, object>() ) );

            var expectedEmail = new RockEmailMessage()
            {
                FromName = expectedFromName,
                FromEmail = expectedFromEmail
            };

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse )
                .Verifiable();

            emailTransport
                .Object
                .Send( actualEmail, 0, new Dictionary<string, string>(), out var errorMessages );

            Assert.That.IsEmpty( errorMessages );

            emailTransport
                .Protected()
                .Verify( "SendEmail",
                    Times.Once(),
                    ItExpr.Is<RockEmailMessage>( rem =>
                        rem.FromEmail == expectedEmail.FromEmail &&
                        rem.FromName == expectedEmail.FromName
                    )
                );
        }

        [TestMethod]
        public void SendRockMessageShouldNotReplaceSafeFromEmail()
        {
            AddSafeDomains();

            var expectedFromEmail = "test@org.com";
            var expectedFromName = "Test Name";

            var globalAttributes = GlobalAttributesCache.Get();
            globalAttributes.SetValue( "OrganizationEmail", "test@organization.com", false, null );

            var actualEmail = new RockEmailMessage()
            {
                FromEmail = expectedFromEmail,
                FromName = expectedFromName
            };

            actualEmail.AddRecipient( new RockEmailMessageRecipient( new Person
            {
                Email = "test@test.com",
                FirstName = "Test",
                LastName = "User"
            }, new Dictionary<string, object>() ) );

            var expectedEmail = new RockEmailMessage()
            {
                FromName = expectedFromName,
                FromEmail = expectedFromEmail
            };

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse )
                .Verifiable();

            emailTransport
                .Object
                .Send( actualEmail, 0, new Dictionary<string, string>(), out var errorMessages );

            Assert.That.IsEmpty( errorMessages );

            emailTransport
                .Protected()
                .Verify( "SendEmail",
                    Times.Once(),
                    ItExpr.Is<RockEmailMessage>( rem =>
                        rem.FromEmail == expectedEmail.FromEmail &&
                        rem.FromName == expectedEmail.FromName
                    )
                );
        }

        [TestMethod]
        public void SendRockMessageShouldNotReplaceUnsafeFromButSafeToEmail()
        {
            AddSafeDomains();

            var expectedFromEmail = "from@organization.com";
            var expectedFromName = "Test Name";

            var globalAttributes = GlobalAttributesCache.Get();
            globalAttributes.SetValue( "OrganizationEmail", "org@organization.com", false, null );

            var actualEmail = new RockEmailMessage()
            {
                FromEmail = expectedFromEmail,
                FromName = expectedFromName
            };

            actualEmail.AddRecipient( new RockEmailMessageRecipient( new Person
            {
                Email = "test@org.com",
                FirstName = "Test",
                LastName = "User"
            }, new Dictionary<string, object>() ) );

            var expectedEmail = new RockEmailMessage()
            {
                FromName = expectedFromName,
                FromEmail = expectedFromEmail
            };

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse )
                .Verifiable();

            emailTransport
                .Object
                .Send( actualEmail, 0, new Dictionary<string, string>(), out var errorMessages );

            Assert.That.IsEmpty( errorMessages );

            emailTransport
                .Protected()
                .Verify( "SendEmail",
                    Times.Once(),
                    ItExpr.Is<RockEmailMessage>( rem =>
                        rem.FromEmail == expectedEmail.FromEmail &&
                        rem.FromName == expectedEmail.FromName
                    )
                );
        }

        [TestMethod]
        public void SendRockMessageShouldReplaceFromWithSafeFromEmailWhenNoSafeToEmailFound()
        {
            AddSafeDomains();

            var expectedFromEmail = "org@organization.com";
            var expectedFromName = "Test Name";

            var globalAttributes = GlobalAttributesCache.Get();
            globalAttributes.SetValue( "OrganizationEmail", expectedFromEmail, false, null );

            var actualEmail = new RockEmailMessage()
            {
                FromEmail = "from@organization.com",
                FromName = expectedFromName
            };

            actualEmail.AddRecipient( new RockEmailMessageRecipient( new Person
            {
                Email = "test@test.com",
                FirstName = "Test",
                LastName = "User"
            }, new Dictionary<string, object>() ) );

            var expectedEmail = new RockEmailMessage()
            {
                FromName = expectedFromName,
                FromEmail = expectedFromEmail
            };

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse )
                .Verifiable();

            emailTransport
                .Object
                .Send( actualEmail, 0, new Dictionary<string, string>(), out var errorMessages );

            Assert.That.IsEmpty( errorMessages );

            emailTransport
                .Protected()
                .Verify( "SendEmail",
                    Times.Once(),
                    ItExpr.Is<RockEmailMessage>( rem =>
                        rem.FromEmail == expectedEmail.FromEmail &&
                        rem.FromName == expectedEmail.FromName
                    )
                );
        }

        [TestMethod]
        public void SendRockMessageWithNoFromEmailShouldGetOrgEmail()
        {
            var expectedFromEmail = "org@organization.com";
            var expectedFromName = "Test Name";

            var globalAttributes = GlobalAttributesCache.Get();
            globalAttributes.SetValue( "OrganizationEmail", expectedFromEmail, false, null );

            var actualEmail = new RockEmailMessage()
            {
                FromEmail = "",
                FromName = expectedFromName
            };

            actualEmail.AddRecipient( new RockEmailMessageRecipient( new Person
            {
                Email = "test@test.com",
                FirstName = "Test",
                LastName = "User"
            }, new Dictionary<string, object>() ) );

            var expectedEmail = new RockEmailMessage()
            {
                FromName = expectedFromName,
                FromEmail = expectedFromEmail
            };

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse )
                .Verifiable();

            emailTransport
                .Object
                .Send( actualEmail, 0, new Dictionary<string, string>(), out var errorMessages );

            Assert.That.IsEmpty( errorMessages );

            emailTransport
                .Protected()
                .Verify( "SendEmail",
                    Times.Once(),
                    ItExpr.Is<RockEmailMessage>( rem =>
                        rem.FromEmail == expectedEmail.FromEmail &&
                        rem.FromName == expectedEmail.FromName
                    )
                );
        }

        [TestMethod]
        public void SendRockMessageWithNoFromEmailAfterLavaEvaluationShouldGetOrgEmail()
        {
            var expectedFromEmail = "org@organization.com";
            var expectedFromName = "Test Name";

            var globalAttributes = GlobalAttributesCache.Get();
            globalAttributes.SetValue( "OrganizationEmail", expectedFromEmail, false, null );

            var actualEmail = new RockEmailMessage()
            {
                FromEmail = "{{fromEmail}}",
                FromName = expectedFromName
            };

            actualEmail.AddRecipient( new RockEmailMessageRecipient( new Person
            {
                Email = "test@test.com",
                FirstName = "Test",
                LastName = "User"
            }, new Dictionary<string, object>() ) );

            var expectedEmail = new RockEmailMessage()
            {
                FromName = expectedFromName,
                FromEmail = expectedFromEmail
            };

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse )
                .Verifiable();

            emailTransport
                .Object
                .Send( actualEmail, 0, new Dictionary<string, string> { { "fromEmail", "" } }, out var errorMessages );

            Assert.That.IsEmpty( errorMessages );

            emailTransport
                .Protected()
                .Verify( "SendEmail",
                    Times.Once(),
                    ItExpr.Is<RockEmailMessage>( rem =>
                        rem.FromEmail == expectedEmail.FromEmail &&
                        rem.FromName == expectedEmail.FromName
                    )
                );
        }

        [TestMethod]
        public void SendRockMessageWithAnInvalidFromEmailShouldCauseError()
        {
            var actualEmail = new RockEmailMessage()
            {
                FromEmail = "invalidEmailAddress",
                FromName = "Test Name"
            };

            actualEmail.AddRecipient( new RockEmailMessageRecipient( new Person
            {
                Email = "test@test.com",
                FirstName = "Test",
                LastName = "User"
            }, new Dictionary<string, object>() ) );

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse );

            emailTransport
                .Object
                .Send( actualEmail, 0, new Dictionary<string, string>(), out var errorMessages );

            Assert.That.AreEqual( 1, errorMessages.Count );
            Assert.That.AreEqual( "The specified string is not in the form required for an e-mail address.", errorMessages[0] );
        }

        [TestMethod]
        public void SendRockMessageWithAnNoFromEmailAndNoOrgEmailShouldCauseError()
        {
            var globalAttributes = GlobalAttributesCache.Get();
            globalAttributes.SetValue( "OrganizationEmail", "", false, null );

            var actualEmail = new RockEmailMessage()
            {
                FromEmail = "",
                FromName = "Test Name"
            };

            actualEmail.AddRecipient( new RockEmailMessageRecipient( new Person
            {
                Email = "test@test.com",
                FirstName = "Test",
                LastName = "User"
            }, new Dictionary<string, object>() ) );

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse );

            emailTransport
                .Object
                .Send( actualEmail, 0, new Dictionary<string, string>(), out var errorMessages );

            Assert.That.AreEqual( 1, errorMessages.Count );
            Assert.That.AreEqual( "A From address was not provided.", errorMessages[0] );
        }

        [TestMethod]
        public void SendRockMessageShouldPopulatePropertiesCorrectly()
        {
            var expectedEmail = "test@test.com";
            var globalAttributes = GlobalAttributesCache.Get();
            globalAttributes.SetValue( "OrganizationEmail", expectedEmail, false, null );

            var actualPerson = new Person
            {
                Email = "test@test.com",
                FirstName = "Test",
                LastName = "User"
            };

            var actualEmailMessage = new RockEmailMessage()
            {
                FromEmail = expectedEmail,
                FromName = "Test Name",
                AppRoot = "test/approot",
                BCCEmails = new List<string> { "bcc1@test.com", "bcc2@test.com" },
                CCEmails = new List<string> { "cc1@test.com", "cc2@test.com" },
                CssInliningEnabled = true,
                CurrentPerson = actualPerson,
                EnabledLavaCommands = "RockEntity",
                Message = "HTML Message",
                MessageMetaData = new Dictionary<string, string> { { "test", "test1" } },
                PlainTextMessage = "Text Message",
                ReplyToEmail = "replyto@email.com",
                Subject = "Test Subject",
            };
            actualEmailMessage.AddRecipient( new RockEmailMessageRecipient( actualPerson, new Dictionary<string, object>() ) );

            var expectedEmailMessage = new RockEmailMessage()
            {
                FromEmail = expectedEmail,
                FromName = "Test Name",
                AppRoot = "test/approot",
                Attachments = new List<BinaryFile> { new BinaryFile { FileName = "test.txt" } },
                BCCEmails = new List<string> { "bcc1@test.com", "bcc2@test.com" },
                CCEmails = new List<string> { "cc1@test.com", "cc2@test.com" },
                CssInliningEnabled = true,
                CurrentPerson = actualPerson,
                EnabledLavaCommands = "RockEntity",
                Message = "HTML Message",
                MessageMetaData = new Dictionary<string, string> { { "test", "test1" } },
                PlainTextMessage = "Text Message",
                ReplyToEmail = "replyto@email.com",
                Subject = "Test Subject",
            };
            expectedEmailMessage.AddRecipient( new RockEmailMessageRecipient( actualPerson, new Dictionary<string, object>() ) );

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse )
                .Verifiable();

            emailTransport
                .Object
                .Send( expectedEmailMessage, 0, new Dictionary<string, string>(), out var errorMessages );

            Assert.That.IsEmpty( errorMessages );

            emailTransport
                .Protected()
                .Verify( "SendEmail",
                    Times.Once(),
                    ItExpr.Is<RockEmailMessage>( rem =>
                        rem.FromEmail == expectedEmailMessage.FromEmail &&
                        rem.FromName == expectedEmailMessage.FromName &&
                        rem.AppRoot == expectedEmailMessage.AppRoot &&
                        rem.CssInliningEnabled == expectedEmailMessage.CssInliningEnabled &&
                        rem.EnabledLavaCommands == expectedEmailMessage.EnabledLavaCommands &&
                        rem.Message == expectedEmailMessage.Message &&
                        rem.PlainTextMessage == expectedEmailMessage.PlainTextMessage &&
                        rem.ReplyToEmail.Contains(expectedEmailMessage.ReplyToEmail ) &&
                        rem.Subject == expectedEmailMessage.Subject &&
                        AreEquivelent(rem.BCCEmails, expectedEmailMessage.BCCEmails ) &&
                        AreEquivelent( rem.CCEmails, expectedEmailMessage.CCEmails )
                    )
                );
        }
        #endregion

        #region Rock Communications Test
        [TestMethod]
        public void SendCommunicationShouldReplaceUnsafeFromWithOrganizationEmail()
        {
            var expectedFromEmail = "test@org.com";
            var expectedFromName = "Test Name";

            var globalAttributes = GlobalAttributesCache.Get();
            globalAttributes.SetValue( "OrganizationEmail", expectedFromEmail, false, null );

            var actualPerson = new Person
            {
                Email = "test@test.com",
                FirstName = "Test",
                LastName = "User"
            };

            var actualCommunication = new Rock.Model.Communication()
            {
                FromEmail = "info@test.com",
                FromName = expectedFromName
            };


            actualCommunication.Recipients.Add( new CommunicationRecipient
            {
                PersonAlias = new PersonAlias { Person = actualPerson }
            } );

            var expectedCommunication = new RockEmailMessage()
            {
                FromName = expectedFromName,
                FromEmail = expectedFromEmail
            };

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse )
                .Verifiable();

            emailTransport
                .Object
                .Send( actualCommunication, 0, new Dictionary<string, string>());

            emailTransport
                .Protected()
                .Verify( "SendEmail",
                    Times.Once(),
                    ItExpr.Is<RockEmailMessage>( rem =>
                        rem.FromEmail == expectedCommunication.FromEmail &&
                        rem.FromName == expectedCommunication.FromName
                    )
                );
        }

        [TestMethod]
        public void SendCommunicationShouldNotReplaceSafeFromEmail()
        {
            AddSafeDomains();

            var expectedFromEmail = "test@org.com";
            var expectedFromName = "Test Name";

            var globalAttributes = GlobalAttributesCache.Get();
            globalAttributes.SetValue( "OrganizationEmail", "test@organization.com", false, null );

            var actualEmail = new RockEmailMessage()
            {
                FromEmail = expectedFromEmail,
                FromName = expectedFromName
            };

            actualEmail.AddRecipient( new RockEmailMessageRecipient( new Person
            {
                Email = "test@test.com",
                FirstName = "Test",
                LastName = "User"
            }, new Dictionary<string, object>() ) );

            var expectedEmail = new RockEmailMessage()
            {
                FromName = expectedFromName,
                FromEmail = expectedFromEmail
            };

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse )
                .Verifiable();

            emailTransport
                .Object
                .Send( actualEmail, 0, new Dictionary<string, string>(), out var errorMessages );

            Assert.That.IsEmpty( errorMessages );

            emailTransport
                .Protected()
                .Verify( "SendEmail",
                    Times.Once(),
                    ItExpr.Is<RockEmailMessage>( rem =>
                        rem.FromEmail == expectedEmail.FromEmail &&
                        rem.FromName == expectedEmail.FromName
                    )
                );
        }

        [TestMethod]
        public void SendCommunicationShouldNotReplaceUnsafeFromButSafeToEmail()
        {
            AddSafeDomains();

            var expectedFromEmail = "from@organization.com";
            var expectedFromName = "Test Name";

            var globalAttributes = GlobalAttributesCache.Get();
            globalAttributes.SetValue( "OrganizationEmail", "org@organization.com", false, null );

            var actualEmail = new RockEmailMessage()
            {
                FromEmail = expectedFromEmail,
                FromName = expectedFromName
            };

            actualEmail.AddRecipient( new RockEmailMessageRecipient( new Person
            {
                Email = "test@org.com",
                FirstName = "Test",
                LastName = "User"
            }, new Dictionary<string, object>() ) );

            var expectedEmail = new RockEmailMessage()
            {
                FromName = expectedFromName,
                FromEmail = expectedFromEmail
            };

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse )
                .Verifiable();

            emailTransport
                .Object
                .Send( actualEmail, 0, new Dictionary<string, string>(), out var errorMessages );

            Assert.That.IsEmpty( errorMessages );

            emailTransport
                .Protected()
                .Verify( "SendEmail",
                    Times.Once(),
                    ItExpr.Is<RockEmailMessage>( rem =>
                        rem.FromEmail == expectedEmail.FromEmail &&
                        rem.FromName == expectedEmail.FromName
                    )
                );
        }

        [TestMethod]
        public void SendCommunicationShouldReplaceFromWithSafeFromEmailWhenNoSafeToEmailFound()
        {
            AddSafeDomains();

            var expectedFromEmail = "org@organization.com";
            var expectedFromName = "Test Name";

            var globalAttributes = GlobalAttributesCache.Get();
            globalAttributes.SetValue( "OrganizationEmail", expectedFromEmail, false, null );

            var actualEmail = new RockEmailMessage()
            {
                FromEmail = "from@organization.com",
                FromName = expectedFromName
            };

            actualEmail.AddRecipient( new RockEmailMessageRecipient( new Person
            {
                Email = "test@test.com",
                FirstName = "Test",
                LastName = "User"
            }, new Dictionary<string, object>() ) );

            var expectedEmail = new RockEmailMessage()
            {
                FromName = expectedFromName,
                FromEmail = expectedFromEmail
            };

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse )
                .Verifiable();

            emailTransport
                .Object
                .Send( actualEmail, 0, new Dictionary<string, string>(), out var errorMessages );

            Assert.That.IsEmpty( errorMessages );

            emailTransport
                .Protected()
                .Verify( "SendEmail",
                    Times.Once(),
                    ItExpr.Is<RockEmailMessage>( rem =>
                        rem.FromEmail == expectedEmail.FromEmail &&
                        rem.FromName == expectedEmail.FromName
                    )
                );
        }

        [TestMethod]
        public void SendCommunicationWithNoFromEmailShouldGetOrgEmail()
        {
            var expectedFromEmail = "org@organization.com";
            var expectedFromName = "Test Name";

            var globalAttributes = GlobalAttributesCache.Get();
            globalAttributes.SetValue( "OrganizationEmail", expectedFromEmail, false, null );

            var actualEmail = new RockEmailMessage()
            {
                FromEmail = "",
                FromName = expectedFromName
            };

            actualEmail.AddRecipient( new RockEmailMessageRecipient( new Person
            {
                Email = "test@test.com",
                FirstName = "Test",
                LastName = "User"
            }, new Dictionary<string, object>() ) );

            var expectedEmail = new RockEmailMessage()
            {
                FromName = expectedFromName,
                FromEmail = expectedFromEmail
            };

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse )
                .Verifiable();

            emailTransport
                .Object
                .Send( actualEmail, 0, new Dictionary<string, string>(), out var errorMessages );

            Assert.That.IsEmpty( errorMessages );

            emailTransport
                .Protected()
                .Verify( "SendEmail",
                    Times.Once(),
                    ItExpr.Is<RockEmailMessage>( rem =>
                        rem.FromEmail == expectedEmail.FromEmail &&
                        rem.FromName == expectedEmail.FromName
                    )
                );
        }

        [TestMethod]
        public void SendCommunicationWithNoFromEmailAfterLavaEvaluationShouldGetOrgEmail()
        {
            var expectedFromEmail = "org@organization.com";
            var expectedFromName = "Test Name";

            var globalAttributes = GlobalAttributesCache.Get();
            globalAttributes.SetValue( "OrganizationEmail", expectedFromEmail, false, null );

            var actualEmail = new RockEmailMessage()
            {
                FromEmail = "{{fromEmail}}",
                FromName = expectedFromName
            };

            actualEmail.AddRecipient( new RockEmailMessageRecipient( new Person
            {
                Email = "test@test.com",
                FirstName = "Test",
                LastName = "User"
            }, new Dictionary<string, object>() ) );

            var expectedEmail = new RockEmailMessage()
            {
                FromName = expectedFromName,
                FromEmail = expectedFromEmail
            };

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse )
                .Verifiable();

            emailTransport
                .Object
                .Send( actualEmail, 0, new Dictionary<string, string> { { "fromEmail", "" } }, out var errorMessages );

            Assert.That.IsEmpty( errorMessages );

            emailTransport
                .Protected()
                .Verify( "SendEmail",
                    Times.Once(),
                    ItExpr.Is<RockEmailMessage>( rem =>
                        rem.FromEmail == expectedEmail.FromEmail &&
                        rem.FromName == expectedEmail.FromName
                    )
                );
        }

        [TestMethod]
        public void SendCommunicationWithAnInvalidFromEmailShouldCauseError()
        {
            var actualEmail = new RockEmailMessage()
            {
                FromEmail = "invalidEmailAddress",
                FromName = "Test Name"
            };

            actualEmail.AddRecipient( new RockEmailMessageRecipient( new Person
            {
                Email = "test@test.com",
                FirstName = "Test",
                LastName = "User"
            }, new Dictionary<string, object>() ) );

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse );

            emailTransport
                .Object
                .Send( actualEmail, 0, new Dictionary<string, string>(), out var errorMessages );

            Assert.That.AreEqual( 1, errorMessages.Count );
            Assert.That.AreEqual( "The specified string is not in the form required for an e-mail address.", errorMessages[0] );
        }

        [TestMethod]
        public void SendCommunicationWithAnNoFromEmailAndNoOrgEmailShouldCauseError()
        {
            var globalAttributes = GlobalAttributesCache.Get();
            globalAttributes.SetValue( "OrganizationEmail", "", false, null );

            var actualEmail = new RockEmailMessage()
            {
                FromEmail = "",
                FromName = "Test Name"
            };

            actualEmail.AddRecipient( new RockEmailMessageRecipient( new Person
            {
                Email = "test@test.com",
                FirstName = "Test",
                LastName = "User"
            }, new Dictionary<string, object>() ) );

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse );

            emailTransport
                .Object
                .Send( actualEmail, 0, new Dictionary<string, string>(), out var errorMessages );

            Assert.That.AreEqual( 1, errorMessages.Count );
            Assert.That.AreEqual( "A From address was not provided.", errorMessages[0] );
        }

        [TestMethod]
        public void SendCommunicationShouldPopulatePropertiesCorrectly()
        {
            var expectedEmail = "test@test.com";
            var globalAttributes = GlobalAttributesCache.Get();
            globalAttributes.SetValue( "OrganizationEmail", expectedEmail, false, null );

            var actualPerson = new Person
            {
                Email = "test@test.com",
                FirstName = "Test",
                LastName = "User"
            };

            var actualEmailMessage = new RockEmailMessage()
            {
                FromEmail = expectedEmail,
                FromName = "Test Name",
                AppRoot = "test/approot",
                BCCEmails = new List<string> { "bcc1@test.com", "bcc2@test.com" },
                CCEmails = new List<string> { "cc1@test.com", "cc2@test.com" },
                CssInliningEnabled = true,
                CurrentPerson = actualPerson,
                EnabledLavaCommands = "RockEntity",
                Message = "HTML Message",
                MessageMetaData = new Dictionary<string, string> { { "test", "test1" } },
                PlainTextMessage = "Text Message",
                ReplyToEmail = "replyto@email.com",
                Subject = "Test Subject",
            };
            actualEmailMessage.AddRecipient( new RockEmailMessageRecipient( actualPerson, new Dictionary<string, object>() ) );

            var expectedEmailMessage = new RockEmailMessage()
            {
                FromEmail = expectedEmail,
                FromName = "Test Name",
                AppRoot = "test/approot",
                Attachments = new List<BinaryFile> { new BinaryFile { FileName = "test.txt" } },
                BCCEmails = new List<string> { "bcc1@test.com", "bcc2@test.com" },
                CCEmails = new List<string> { "cc1@test.com", "cc2@test.com" },
                CssInliningEnabled = true,
                CurrentPerson = actualPerson,
                EnabledLavaCommands = "RockEntity",
                Message = "HTML Message",
                MessageMetaData = new Dictionary<string, string> { { "test", "test1" } },
                PlainTextMessage = "Text Message",
                ReplyToEmail = "replyto@email.com",
                Subject = "Test Subject",
            };
            expectedEmailMessage.AddRecipient( new RockEmailMessageRecipient( actualPerson, new Dictionary<string, object>() ) );

            var emailSendResponse = new EmailSendResponse
            {
                Status = Rock.Model.CommunicationRecipientStatus.Delivered,
                StatusNote = "Email Sent."
            };

            var emailTransport = new Mock<EmailTransportComponent>()
            {
                CallBase = true
            };

            emailTransport
                .Protected()
                .Setup<EmailSendResponse>( "SendEmail", ItExpr.IsAny<RockEmailMessage>() )
                .Returns( emailSendResponse )
                .Verifiable();

            emailTransport
                .Object
                .Send( expectedEmailMessage, 0, new Dictionary<string, string>(), out var errorMessages );

            Assert.That.IsEmpty( errorMessages );

            emailTransport
                .Protected()
                .Verify( "SendEmail",
                    Times.Once(),
                    ItExpr.Is<RockEmailMessage>( rem =>
                        rem.FromEmail == expectedEmailMessage.FromEmail &&
                        rem.FromName == expectedEmailMessage.FromName &&
                        rem.AppRoot == expectedEmailMessage.AppRoot &&
                        rem.CssInliningEnabled == expectedEmailMessage.CssInliningEnabled &&
                        rem.EnabledLavaCommands == expectedEmailMessage.EnabledLavaCommands &&
                        rem.Message == expectedEmailMessage.Message &&
                        rem.PlainTextMessage == expectedEmailMessage.PlainTextMessage &&
                        rem.ReplyToEmail.Contains( expectedEmailMessage.ReplyToEmail ) &&
                        rem.Subject == expectedEmailMessage.Subject &&
                        AreEquivelent( rem.BCCEmails, expectedEmailMessage.BCCEmails ) &&
                        AreEquivelent( rem.CCEmails, expectedEmailMessage.CCEmails )
                    )
                );
        }
        #endregion
        private void AddSafeDomains()
        {
            using ( var rockContext = new RockContext() )
            {
                var definedTypeService = new DefinedValueService( rockContext );
                var definedType = new DefinedTypeService( rockContext ).Get( SystemGuid.DefinedType.COMMUNICATION_SAFE_SENDER_DOMAINS.AsGuid() );

                rockContext.Database.ExecuteSqlCommand( $"DELETE DefinedValue WHERE DefinedTypeId = {definedType.Id} AND Value = 'org.com'" );

                var definedValue = new DefinedValue { Id = 0 };
                definedValue.DefinedTypeId = definedType.Id;
                definedValue.IsSystem = false;
                definedValue.Value = "org.com";
                definedValue.Description = "This is a test safe domain.";
                definedValue.IsActive = true;
                definedTypeService.Add( definedValue );
                rockContext.SaveChanges();

                definedValue.LoadAttributes();
                definedValue.SetAttributeValue( "SafeToSendTo", "true" );
                definedValue.SaveAttributeValues( rockContext );
            }
        }

        private bool AreEquivelent<T>(List<T> list1, List<T> list2 )
        {
            Assert.That.AreEqual( list1, list2 );
            return true;
        }
    }
}
