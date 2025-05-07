using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Description;

namespace Havells.CRM.DataMigration
{
    internal class ServiceMasterMigration : HelperClass
    {
        private static EntityReference systemAdminRef = null;
        static ServiceMasterMigration()
        {
            systemAdminRef = new EntityReference("systemuser", new Guid(SystemAdmin));
        }
        public static void serviceMasterMigration(IOrganizationService _servicePrd, IOrganizationService _serviceDev)
        {
            string[] entityList = {
                "hil_calltype",
                "hil_callsubtype",
                "hil_claimtype",
                "hil_claimcategory",
                "hil_claimperiod",
                "hil_claimpostingsetup",
                "hil_serviceactionwork",
                "hil_sawcategoryapprovals",
                "hil_serviceactionworksetup",
                "hil_consumercategory",
                "hil_consumertype"
            };
            Parallel.ForEach(entityList, entityName =>
            {
                MigrateBulkRecords(_servicePrd, _serviceDev, entityName, systemAdminRef, true);
            });
        }
    }
}
