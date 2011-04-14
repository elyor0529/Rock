//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by the T4\Model.tt template.
//
//     Changes to this file will be lost when the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------
//
// THIS WORK IS LICENSED UNDER A CREATIVE COMMONS ATTRIBUTION-NONCOMMERCIAL-
// SHAREALIKE 3.0 UNPORTED LICENSE:
// http://creativecommons.org/licenses/by-nc-sa/3.0/
//
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data.Entity.ModelConfiguration;
using System.Linq;
using System.Text;

using Rock.Models;

namespace Rock.Models.Core
{
    [Table( "coreEntityChange" )]
    public partial class EntityChange : ModelWithAttributes
    {
		public Guid Guid { get; set; }
		
		public Guid ChangeSet { get; set; }
		
		[MaxLength( 10 )]
		public string ChangeType { get; set; }
		
		[MaxLength( 100 )]
		public string EntityType { get; set; }
		
		public int EntityId { get; set; }
		
		[MaxLength( 100 )]
		public string Property { get; set; }
		
		public string OriginalValue { get; set; }
		
		public string CurrentValue { get; set; }
		
		[NotMapped]
		public override string AuthEntity { get { return "Core.EntityChange"; } }
    }

    public partial class EntityChangeConfiguration : EntityTypeConfiguration<EntityChange>
    {
        public EntityChangeConfiguration()
        {
		}
    }
}
