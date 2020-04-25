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
//
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.ModelConfiguration;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.Serialization;
using EntityFramework.Utilities;
using Rock.Data;

namespace Rock.Model
{
    /// <summary>
    /// A table of all available response codes for <see cref="CommunicationRecipient"/>
    /// This will have a fix number of records from 100-99999 (excluding 666 and 911)
    /// a prefix of '@'. For example, '@126345'
    /// </summary>
    [RockDomain( "Communication" )]
    [Table( "CommunicationRecipientResponseCode" )]
    [DataContract]
    public class CommunicationRecipientResponseCode
    {
        /// <summary>
        /// Gets or sets the identifier.
        /// </summary>
        /// <value>
        /// The identifier.
        /// </value>
        [DataMember]
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// The response code from 100-99999 (excluding 666 and 911)
        /// a prefix of '@'. For example, '@126345'
        /// Note: this numeric portion must be between 3 and 5 digits
        /// </summary>
        /// <value>
        /// The response code.
        /// </value>
        [DataMember]
        [MaxLength( 6 )]
        [Index( "IX_ResponseCode", IsUnique = true )]
        public string ResponseCode { get; set; }

        /// <summary>
        /// The last date time that this <see cref="ResponseCode"/> was used for a <see cref="CommunicationRecipient" />
        /// When <see cref="CommunicationRecipient"/> needs a response code, it'll pick a random one that hasn't been used in the last 30 days
        /// </summary>
        /// <value>
        /// The last used date t IME.
        /// </value>
        [Index( "IX_LastUsedDateTime" )]
        public DateTime? LastUsedDateTime { get; set; }

        /// <summary>
        /// The response code maximum
        /// </summary>
        private const int RESPONSE_CODE_MAX = 99999;

        /// <summary>
        /// number of days between token reuse
        /// </summary>
        private const int TOKEN_REUSE_DURATION_DAYS = 30; // 

        /// <summary>
        /// The response code blacklist
        /// </summary>
        private static readonly int[] RESPONSE_CODE_BLACKLIST = new int[] { 666, 911 };

        /// <summary>
        /// 
        /// </summary>
        public class NewResponseCode
        {
            /// <summary>
            /// The response code identifier to use for CommunicationRecipient.CommunicationRecipientResponseCodeId
            /// </summary>
            public int ResponseCodeId { get; set; }

            /// <summary>
            /// Gets or sets the response code, for example '@123456'
            /// </summary>
            /// <value>
            /// The response code.
            /// </value>
            public string ResponseCode { get; private set; }
        }

        /// <summary>
        /// Gets the response code identifier from response code.
        /// </summary>
        /// <param name="responseCode">The response code.</param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        [RockObsolete( "1.11" )]
        [Obsolete( "Use GetNewResponseCode instead" )]
        public static int? GetResponseCodeIdFromResponseCode( string responseCode )
        {
            // TODO
            using ( var rockContext = new RockContext() )
            {
                var responseCodeRecord = rockContext.CommunicationRecipientResponseCodes.Where( a => a.ResponseCode == responseCode ).FirstOrDefault();
                if ( responseCodeRecord != null )
                {
                    responseCodeRecord.LastUsedDateTime = RockDateTime.Now;
                    rockContext.SaveChanges( disablePrePostProcessing: true );
                    return responseCodeRecord.Id;
                }
            }

            return null;
        }

        /// <summary>
        /// Gets the new response code identifier.
        /// </summary>
        /// <returns></returns>
        /// <exception cref="Exception">Could not find an available response code.</exception>
        public static NewResponseCode GetNewResponseCode()
        {
            NewResponseCode newResponseCode;

            using ( var rockContext = new RockContext() )
            {
                // select a 1000 Unused codes, randomize them, then return the first one in the list
                // this keeps it random without needing to randomize all 99999 rows (which slows it down by 30x)
                // getting a random unused one this way only takes about 2 ms (vs 30ms if we randomized the first 99999 unused ones)
                int sampleGroupSize = 1000;

                // Use raw SQL to get an unused code that immediately get marked as used before return the result.
                // This along with the UPDLOCK will make it impossible to have multiple different threads get the same ResponseCodeId.
                // Also, using sampleGroupSize and TOKEN_REUSE_DURATION_DAYS performs better when they are *not* set using SqlParameeter
                newResponseCode = rockContext.Database.SqlQuery<NewResponseCode>( $@"
UPDATE [CommunicationRecipientResponseCode]
SET [LastUsedDateTime] = @requestDateTime
OUTPUT inserted.Id [ResponseCodeId], inserted.ResponseCode
WHERE [Id] = (
		SELECT TOP 1 [Id]
		FROM (
			SELECT TOP ({sampleGroupSize}) [Id]
			FROM [CommunicationRecipientResponseCode] with (UPDLOCK)
			WHERE [LastUsedDateTime] IS NULL OR [LastUsedDateTime] < DATEADD(DAY, -{TOKEN_REUSE_DURATION_DAYS}, @requestDateTime)
			) x
		ORDER BY NEWID()
		)"
                , new SqlParameter( "@requestDateTime", RockDateTime.Now ) ).FirstOrDefault();
            }

            if ( newResponseCode != null )
            {
                return newResponseCode;
            }

            throw new Exception( "Could not find an available response code." );
        }

        /// <summary>
        /// Ensures the CommunicationRecipientResponseCode is populated with the predefined list of ResponseCodes
        /// </summary>
        public static void EnsurePopulated()
        {
            List<string> responseCodeList = new List<string>();
            for ( int responseCode = 100; responseCode <= RESPONSE_CODE_MAX; responseCode++ )
            {
                if ( !RESPONSE_CODE_BLACKLIST.Contains( responseCode ) )
                {
                    responseCodeList.Add( $"@{responseCode}" );
                }
            }

            using ( var rockContext = new RockContext() )
            {
                var existingCommunicationRecipientResponseCodes = new HashSet<string>( rockContext.CommunicationRecipientResponseCodes.Select( a => a.ResponseCode ).ToArray() );
                var missingResponseCodes = responseCodeList
                    .Where( a => !existingCommunicationRecipientResponseCodes.Contains( a ) )
                    .ToList();

                var communicationRecipientResponseCodesToInsert = new List<CommunicationRecipientResponseCode>();

                foreach ( var missingResponseCode in missingResponseCodes )
                {
                    var communicationRecipientResponseCode = new CommunicationRecipientResponseCode
                    {
                        ResponseCode = missingResponseCode,
                        LastUsedDateTime = null
                    };

                    communicationRecipientResponseCodesToInsert.Add( communicationRecipientResponseCode );
                }

                // NOTE: We can't use rockContext.BulkInsert because that enforces that the <T> is Rock.Data.IEntity, so we'll just use EFBatchOperation directly
                EFBatchOperation.For( rockContext, rockContext.CommunicationRecipientResponseCodes ).InsertAll( communicationRecipientResponseCodesToInsert );
            }
        }
    }

    #region Entity Configuration

    /// <summary>
    /// 
    /// </summary>
    public partial class CommunicationRecipientResponseCodeConfiguration : EntityTypeConfiguration<CommunicationRecipientResponseCode>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AnalyticsSourcePersonHistoricalConfiguration"/> class.
        /// </summary>
        public CommunicationRecipientResponseCodeConfiguration()
        {
        }
    }

    #endregion Entity Configuration
}
