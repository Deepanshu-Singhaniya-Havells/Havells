using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.DataMigration
{
    public class MigratePriceList : HelperClass
    {
        private static EntityReference systemAdminRef = null;
        static MigratePriceList()
        {
            systemAdminRef = new EntityReference("systemuser", new Guid(SystemAdmin));
        }
        public static void migratePriceList(IOrganizationService _servicePrd, IOrganizationService _serviceDev)
        {
            MigrateBulkRecords(_servicePrd, _serviceDev, "pricelevel", systemAdminRef,false);
        }
    }
}
