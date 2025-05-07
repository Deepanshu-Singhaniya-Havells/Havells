using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin
{
    public class HelperPO
    {
        public static Guid CreatePO(IOrganizationService service,Guid fsOwnerId, Guid fsJobId,Guid fsPOHeader, int iQuantity, Guid fsAccountId, Guid fsPartId, int PoType,OptionSetValue opCategory,String sCustomerSAPCOde,String sDistributionChannel,OptionSetValue opWarrantyType)
        {
            try
            {
                hil_productrequest PO = new hil_productrequest();
                PO.hil_SuperFranchiseeDSEName = new EntityReference(Account.EntityLogicalName, fsAccountId);
                PO.hil_PRType = new OptionSetValue(PoType);
                PO.hil_Category = opCategory;
                PO.hil_Quantity = iQuantity;
                PO.hil_PRDate = DateTime.UtcNow;
                PO.hil_CustomerSAPCode = sCustomerSAPCOde;
                PO.hil_DistributionChannel = sDistributionChannel;
                PO.hil_WarrantyStatus = opWarrantyType;
                if (fsOwnerId != Guid.Empty)
                {
                    PO.OwnerId = new EntityReference(SystemUser.EntityLogicalName,fsOwnerId);
                }
                PO.hil_PartCode = new EntityReference(Product.EntityLogicalName, fsPartId);
                if (fsJobId != Guid.Empty)
                {
                    PO.hil_Job = new EntityReference(msdyn_workorder.EntityLogicalName, fsJobId);
                }
                if (fsPOHeader != Guid.Empty)
                {
                    PO.hil_PRHeader = new EntityReference(hil_productrequestheader.EntityLogicalName, fsPOHeader);
                }
                Guid fsPOId= service.Create(PO);
                return fsPOId;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.HelpserClasses.HelperPO.CreatePO" + ex.Message);
            }
        }
        public static Guid CreatePOHeader(IOrganizationService service, EntityReference JobRef, EntityReference RefOwner, EntityReference AccountIdRef, EntityReference DivisionRef, String sDivisionSAPCOde,  int PoType, String sCustomerSAPCOde, OptionSetValue opWarrantyType)
        {
            try
            {
                hil_productrequestheader PO = new hil_productrequestheader();
                if (AccountIdRef != null)
                    PO.hil_Account = AccountIdRef;
                PO.hil_PRType = new OptionSetValue(PoType);
                PO.hil_CustomerSAPCode = sCustomerSAPCOde;
                PO.hil_WarrantyStatus = opWarrantyType;
                PO.hil_DivisionCode = sDivisionSAPCOde;
                if (DivisionRef!= null)
                {
                    PO.hil_Division = DivisionRef;
                }
                if (JobRef!= null)
                {
                    PO.hil_Job = JobRef;
                }
                if (RefOwner != null)
                {
                    PO.OwnerId = new EntityReference(SystemUser.EntityLogicalName, RefOwner.Id);
                }
                Guid fsPOId = service.Create(PO);
                return fsPOId;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.HelpserClasses.HelperPO.CreatePOHeader" + ex.Message);
            }
        }
        public static String getDistributionChannel(OptionSetValue opWarrantyStatus, OptionSetValue opAccountType, IOrganizationService service)
        {
            String result = String.Empty;
            try
            {
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    var obj = from _config in orgContext.CreateQuery<hil_integrationconfiguration>()
                              where _config.new_WarrantyStatus == opWarrantyStatus
                              && _config.new_CustomerType == opAccountType
                              select new
                              {
                                  _config.new_DistributionChannel
                              };
                    foreach (var iobj in obj)
                    {
                        if (iobj.new_DistributionChannel != null)
                        {
                            result = iobj.new_DistributionChannel;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.HelpserClasses.HelperPO.getDistributionChannel" + ex.Message);
            }
            return result;
        }
    }
}
