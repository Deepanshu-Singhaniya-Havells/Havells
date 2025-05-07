using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.DataMigration
{
    public class MigrateUserSetup:HelperClass
    {
        private static EntityReference systemAdminRef = null;
        static MigrateUserSetup()
        {
            systemAdminRef = new EntityReference("systemuser", new Guid(SystemAdmin));
        }
        public static void migrateUserSetup(IOrganizationService _servicePrd, IOrganizationService _serviceDev)
        {
            MigrateBulkRecords(_servicePrd, _serviceDev, "position", systemAdminRef,false);
        }
    }
}
