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
namespace Rock.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    /// <summary>
    ///
    /// </summary>
    public partial class CommunicationRecipientResponseCode : Rock.Migrations.RockMigration
    {
        /// <summary>
        /// Operations to be performed during the upgrade process.
        /// </summary>
        public override void Up()
        {
            CreateTable(
                "dbo.CommunicationRecipientResponseCode",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        ResponseCode = c.String(maxLength: 6),
                        LastUsedDateTime = c.DateTime(),
                    })
                .PrimaryKey(t => t.Id)
                .Index(t => t.ResponseCode, unique: true)
                .Index(t => t.LastUsedDateTime);
            
            AddColumn("dbo.CommunicationRecipient", "CommunicationRecipientResponseCodeId", c => c.Int());
            CreateIndex("dbo.CommunicationRecipient", "CommunicationRecipientResponseCodeId");

            AddForeignKey( "dbo.CommunicationRecipient", "CommunicationRecipientResponseCodeId", "dbo.CommunicationRecipientResponseCode", "Id");

            // Populate Response Codes
            Sql( @"
DECLARE @responseCodeNumber INT = 100;
-- populate a temp table so that we can do a batch insert
DECLARE @responseCodeIds TABLE (Id INT)

WHILE (@responseCodeNumber <= 99999)
BEGIN
	IF (@responseCodeNumber != 666 AND @responseCodeNumber != 911)
	BEGIN
		INSERT INTO @responseCodeIds (id)
		VALUES (@responseCodeNumber)
	END

	SET @responseCodeNumber = @responseCodeNumber + 1;
END

INSERT INTO CommunicationRecipientResponseCode (ResponseCode)
SELECT CONCAT ('@', Id)
FROM @responseCodeIds
" );

            // Update CommunicationRecipient to set CommunicationRecipientResponseCodeId based on ResponseCode
            Sql( @"
UPDATE CR
SET cr.CommunicationRecipientResponseCodeId = rc.Id
FROM CommunicationRecipient cr
JOIN CommunicationRecipientResponseCode rc ON cr.ResponseCode = rc.ResponseCode
WHERE cr.ResponseCode IS NOT NULL
and cr.CommunicationRecipientResponseCodeId != rc.Id");

            // Set the LastUsedDateTime on CommunicationRecipientResponseCode to the most recent usage in CommunicationRecipient
            Sql( @"
UPDATE rc
SET rc.LastUsedDateTime = CASE 
		WHEN rc.LastUsedDateTime < cr.CreatedDateTime
			THEN cr.CreatedDateTime
		ELSE rc.LastUsedDateTime
		END
FROM CommunicationRecipient cr
JOIN CommunicationRecipientResponseCode rc ON cr.ResponseCode = rc.ResponseCode
WHERE cr.ResponseCode IS NOT NULL
AND  rc.LastUsedDateTime != CASE 
		WHEN rc.LastUsedDateTime < cr.CreatedDateTime
			THEN cr.CreatedDateTime
		ELSE rc.LastUsedDateTime
		END
" );
        }
        
        /// <summary>
        /// Operations to be performed during the downgrade process.
        /// </summary>
        public override void Down()
        {
            DropForeignKey("dbo.CommunicationRecipient", "CommunicationRecipientResponseCodeId", "dbo.CommunicationRecipientResponseCode");
            DropIndex("dbo.CommunicationRecipientResponseCode", new[] { "LastUsedDateTime" });
            DropIndex("dbo.CommunicationRecipientResponseCode", new[] { "ResponseCode" });
            DropIndex("dbo.CommunicationRecipient", new[] { "CommunicationRecipientResponseCodeId" });
            DropColumn("dbo.CommunicationRecipient", "CommunicationRecipientResponseCodeId");
            DropTable("dbo.CommunicationRecipientResponseCode");
        }
    }
}
