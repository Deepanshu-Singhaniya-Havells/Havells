using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.DataMigration
{
    public class MigrateMaterialGroups : HelperClass
    {
        private static EntityReference systemAdminRef = null;
        static MigrateMaterialGroups()
        {
            systemAdminRef = new EntityReference("systemuser", new Guid(SystemAdmin));
        }
        public static void migrateMaterialGropus(IOrganizationService _servicePrd, IOrganizationService _serviceDev)
        {
            MigrateBulkRecords(_servicePrd, _serviceDev, "hil_materialgroup", systemAdminRef, true);
            MigrateBulkRecords(_servicePrd, _serviceDev, "hil_materialgroup2", systemAdminRef, true);
            MigrateBulkRecords(_servicePrd, _serviceDev, "hil_materialgroup3", systemAdminRef, true);
            MigrateBulkRecords(_servicePrd, _serviceDev, "hil_materialgroup4", systemAdminRef, true);
            MigrateBulkRecords(_servicePrd, _serviceDev, "hil_materialgroup5", systemAdminRef, true);
        }
    }
}
