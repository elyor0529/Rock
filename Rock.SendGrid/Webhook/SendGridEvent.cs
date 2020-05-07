// <copyright>
// Copyright by the Spark Development Network
//
// Licensed under the Rock Community License (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
// http://www.rockrms.com/license
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
// </copyright>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using UAParser;

namespace Rock.SendGrid.Webhook
{
    [JsonObject( IsReference = false )]
    public class SendGridEvent
    {
        [JsonProperty( PropertyName = "email" )]
        public string Email { get; set; }

        [JsonProperty( PropertyName = "timestamp" )]
        public int Timestamp { get; set; }

        [JsonProperty( PropertyName = "event" )]
        public string EventType { get; set; }

        [JsonProperty( PropertyName = "smtp-id" )]
        public string SmtpId { get; set; }

        [JsonProperty( PropertyName = "useragent" )]
        public string UserAgent { get; set; }

        [JsonProperty( PropertyName = "ip" )]
        public string IpAddress { get; set; }

        [JsonProperty( PropertyName = "sg_event_id" )]
        public string SendGridEventId { get; set; }

        [JsonProperty( PropertyName = "sg_message_id" )]
        public string SendGridMessageId { get; set; }

        [JsonProperty( PropertyName = "reason" )]
        public string EventTypeReason { get; set; }

        [JsonProperty( PropertyName = "status" )]
        public string Status { get; set; }

        [JsonProperty( PropertyName = "response" )]
        public string ServerResponse { get; set; }

        [JsonProperty( PropertyName = "tls" )]
        public string Tls { get; set; }

        [JsonProperty( PropertyName = "url" )]
        public string Url { get; set; }

        [JsonProperty( PropertyName = "urloffset" )]
        public int UrlOffset { get; set; }

        [JsonProperty( PropertyName = "attempt" )]
        public string DeliveryAttemptCount { get; set; }

        [JsonProperty( PropertyName = "category" )]
        public string Category { get; set; }

        [JsonProperty( PropertyName = "type" )]
        public string BounceType { get; set; }

        /// <summary>
        /// Gets or sets the workflow action unique identifier.
        /// </summary>
        /// <value>
        /// The workflow action unique identifier.
        /// </value>
        [JsonProperty( PropertyName = "workflow_action_guid" )]
        public string WorkflowActionGuid { get; set; }

        /// <summary>
        /// Gets or sets the communication recipient unique identifier.
        /// </summary>
        /// <value>
        /// The communication recipient unique identifier.
        /// </value>
        [JsonProperty( PropertyName = "communication_recipient_guid" )]
        public string CommunicationRecipientGuid { get; set; }

        public string ClientOs
        {
            get
            {
                var clientInfo = GetClientInfo();
                return clientInfo.OS.Family;
            }
        }

        public string ClientBrowser
        {
            get
            {
                var clientInfo = GetClientInfo();
                return clientInfo.UA.Family;
            }
        }

        public string ClientDeviceType
        {
            get
            {
                var clientInfo = GetClientInfo();
                return clientInfo.Device.Family;
            }
        }

        public string ClientDeviceBrand
        {
            get
            {
                var clientInfo = GetClientInfo();
                return clientInfo.Device.Brand;
            }
        }

        private ClientInfo _clientInfo = null;
        private ClientInfo GetClientInfo()
        {
            if( _clientInfo == null && UserAgent.IsNotNullOrWhiteSpace() )
            {
                var parser = Parser.GetDefault();
                _clientInfo = parser.Parse( UserAgent );
            }
            return _clientInfo;
        }
    }
}
