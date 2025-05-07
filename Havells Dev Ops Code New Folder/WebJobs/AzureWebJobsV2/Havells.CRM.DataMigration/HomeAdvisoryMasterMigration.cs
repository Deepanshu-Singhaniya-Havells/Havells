using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.DataMigration
{
    public class HomeAdvisoryMasterMigration : HelperClass
    {
        private static EntityReference systemAdminRef = null;
        static HomeAdvisoryMasterMigration()
        {
            systemAdminRef = new EntityReference("systemuser", new Guid(SystemAdmin));
        }
        public static void homeAdvisoryMasterMigration(IOrganizationService _servicePrd, IOrganizationService _serviceDev)
        {
            //   MigrateBulkRecords(_servicePrd, _serviceDev, "hil_enquirytype", systemAdminRef,true);
            //   MigrateBulkRecords(_servicePrd, _serviceDev, "hil_typeofproduct", systemAdminRef,true);
            MigrateBulkRecords(_servicePrd, _serviceDev, "hil_advisormaster", systemAdminRef, true);
        }
    }
}
