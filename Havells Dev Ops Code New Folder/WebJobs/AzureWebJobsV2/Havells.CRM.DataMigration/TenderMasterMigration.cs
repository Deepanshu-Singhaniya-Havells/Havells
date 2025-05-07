using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.DataMigration
{
    public class TenderMasterMigration : HelperClass
    {
        private static EntityReference systemAdminRef = null;
        static TenderMasterMigration()
        {
            systemAdminRef = new EntityReference("systemuser", new Guid(SystemAdmin));
        }
        public static void tenderMasterMigration(IOrganizationService _servicePrd, IOrganizationService _serviceDev)
        {
            //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_enquirydepartment", systemAdminRef,true);
            //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_attachmentdocumenttype", systemAdminRef,true);
            // MigrateBulkRecords(_servicePrd, _serviceDev, "hil_userbranchmapping", systemAdminRef,true);
            //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_designteambranchmapping", systemAdminRef,true);
            //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_bdteam", systemAdminRef, true);
            //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_bdteammember", systemAdminRef, true);
            //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_approvalmatrix", systemAdminRef, true);
            //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_approvalmatrixline", systemAdminRef, true);
            //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_constructiontype", systemAdminRef, true);
            //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_effeciency", systemAdminRef, true);

            //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_enclosuretype", systemAdminRef, true);
            //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_enquerysegment", systemAdminRef, true);
            //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_entityfieldsmetadata", systemAdminRef, true);
            //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_frequency", systemAdminRef, true);


            //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_hsncode", systemAdminRef, true);
            //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_industrytype", systemAdminRef, true);
            //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_inspectiontype", systemAdminRef, true);
            //MigrateBulkRecords(_servicePrd, _serviceDev, "hil_industrysubtype", systemAdminRef, true);
        }
    }
}
