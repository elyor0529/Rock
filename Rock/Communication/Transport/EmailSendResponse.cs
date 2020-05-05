using Rock.Model;

namespace Rock.Communication.Transport
{
    public class EmailSendResponse
    {
        public CommunicationRecipientStatus Status { get; set; }
        public string StatusNote { get; set; }
    }
}
