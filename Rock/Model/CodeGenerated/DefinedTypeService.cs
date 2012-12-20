//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by the Rock.CodeGeneration project
//     Changes to this file will be lost when the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------
//
// THIS WORK IS LICENSED UNDER A CREATIVE COMMONS ATTRIBUTION-NONCOMMERCIAL-
// SHAREALIKE 3.0 UNPORTED LICENSE:
// http://creativecommons.org/licenses/by-nc-sa/3.0/
//

using System;
using System.Linq;

using Rock.Data;

namespace Rock.Model
{
    /// <summary>
    /// DefinedType Service class
    /// </summary>
    public partial class DefinedTypeService : Service<DefinedType>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DefinedTypeService"/> class
        /// </summary>
        public DefinedTypeService()
            : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DefinedTypeService"/> class
        /// </summary>
        public DefinedTypeService(IRepository<DefinedType> repository) : base(repository)
        {
        }

        /// <summary>
        /// Determines whether this instance can delete the specified item.
        /// </summary>
        /// <param name="item">The item.</param>
        /// <param name="errorMessage">The error message.</param>
        /// <returns>
        ///   <c>true</c> if this instance can delete the specified item; otherwise, <c>false</c>.
        /// </returns>
        public bool CanDelete( DefinedType item, out string errorMessage )
        {
            errorMessage = string.Empty;
 
            if ( new Service<DefinedValue>().Queryable().Any( a => a.DefinedTypeId == item.Id ) )
            {
                errorMessage = string.Format( "This {0} is assigned to a {1}.", DefinedType.FriendlyTypeName, DefinedValue.FriendlyTypeName );
                return false;
            }  
            return true;
        }
    }
}
