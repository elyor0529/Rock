using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using Rock.Web.Cache;

namespace Rock.Communication.Transport
{
    public class SafeSenderResult
    {
        public bool IsUnsafeDomain {get; set;}
        public MailAddress SafeFromAddress { get; set; }
    }

    public static class MailTransportHelper
    {
        public static SafeSenderResult CheckSafeSender( List<string> toEmailAddresses, MailAddress fromEmail, string organizationEmail )
        {
            var result = new SafeSenderResult();

            // Get the safe sender domains
            var safeDomainValues = DefinedTypeCache.Get( SystemGuid.DefinedType.COMMUNICATION_SAFE_SENDER_DOMAINS.AsGuid() ).DefinedValues;
            var safeDomains = safeDomainValues.Select( v => v.Value ).ToList();

            // Check to make sure the From email domain is a safe sender, if so then we are done.
            var fromParts = fromEmail.Address.Split( new char[] { '@' }, StringSplitOptions.RemoveEmptyEntries );
            if ( fromParts.Length == 2 && safeDomains.Contains( fromParts[1], StringComparer.OrdinalIgnoreCase ) )
            {
                return result;
            }

            // The sender domain is not considered safe so check all the recipients to see if they have a domain that does not requrie a safe sender
            foreach ( var toEmailAddress in toEmailAddresses )
            {
                bool safe = false;
                var toParts = toEmailAddress.Split( new char[] { '@' }, StringSplitOptions.RemoveEmptyEntries );
                if ( toParts.Length == 2 && safeDomains.Contains( toParts[1], StringComparer.OrdinalIgnoreCase ) )
                {
                    var domain = safeDomainValues.FirstOrDefault( dv => dv.Value.Equals( toParts[1], StringComparison.OrdinalIgnoreCase ) );
                    safe = domain != null && domain.GetAttributeValue( "SafeToSendTo" ).AsBoolean();
                }

                if ( !safe )
                {
                    result.IsUnsafeDomain = true;
                    break;
                }
            }

            if ( result.IsUnsafeDomain )
            {
                if ( !string.IsNullOrWhiteSpace( organizationEmail ) && !organizationEmail.Equals( fromEmail.Address, StringComparison.OrdinalIgnoreCase ) )
                {
                    result.SafeFromAddress = new MailAddress(organizationEmail, fromEmail.DisplayName);
                }
            }

            return result;
        }
    }
}
