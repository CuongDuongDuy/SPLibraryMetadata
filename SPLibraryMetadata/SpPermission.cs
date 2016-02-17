using System.Collections.Generic;
using Microsoft.SharePoint.Client;

namespace SPLibraryMetadata
{
    public class SpPermission
    {
        public RoleType RoleType { get; set; }
        public string[] UsersOrGroups { get; set; }

        public SpPermission(string[] usersOrGroups, RoleType roleType)
        {
            UsersOrGroups = usersOrGroups;
            RoleType = roleType;
        }
    }
}
