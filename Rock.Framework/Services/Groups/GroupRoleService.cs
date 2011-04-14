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
using System.Linq;
using System.Text;

using Rock.Models.Groups;
using Rock.Repository.Groups;

namespace Rock.Services.Groups
{
    public partial class GroupRoleService : Rock.Services.Service
    {
        private IGroupRoleRepository _repository;

        public GroupRoleService()
			: this( new EntityGroupRoleRepository() )
        { }

        public GroupRoleService( IGroupRoleRepository GroupRoleRepository )
        {
            _repository = GroupRoleRepository;
        }

        public IQueryable<Rock.Models.Groups.GroupRole> Queryable()
        {
            return _repository.AsQueryable();
        }

        public Rock.Models.Groups.GroupRole GetGroupRole( int id )
        {
            return _repository.FirstOrDefault( t => t.Id == id );
        }
		
        public IEnumerable<Rock.Models.Groups.GroupRole> GetGroupRolesByGuid( Guid guid )
        {
            return _repository.Find( t => t.Guid == guid );
        }
		
        public IEnumerable<Rock.Models.Groups.GroupRole> GetGroupRolesByOrder( int? order )
        {
            return _repository.Find( t => t.Order == order );
        }
		
        public void AddGroupRole( Rock.Models.Groups.GroupRole GroupRole )
        {
            if ( GroupRole.Guid == Guid.Empty )
                GroupRole.Guid = Guid.NewGuid();

            _repository.Add( GroupRole );
        }

        public void DeleteGroupRole( Rock.Models.Groups.GroupRole GroupRole )
        {
            _repository.Delete( GroupRole );
        }

        public void Save( Rock.Models.Groups.GroupRole GroupRole, int? personId )
        {
            List<Rock.Models.Core.EntityChange> entityChanges = _repository.Save( GroupRole, personId );

			if ( entityChanges != null )
            {
                Rock.Services.Core.EntityChangeService entityChangeService = new Rock.Services.Core.EntityChangeService();

                foreach ( Rock.Models.Core.EntityChange entityChange in entityChanges )
                {
                    entityChange.EntityId = GroupRole.Id;
                    entityChangeService.AddEntityChange ( entityChange );
                    entityChangeService.Save( entityChange, personId );
                }
            }
        }
    }
}
