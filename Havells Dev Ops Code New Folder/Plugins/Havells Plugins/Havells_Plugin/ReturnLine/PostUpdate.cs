using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;

namespace Havells_Plugin.ReturnLine
{
 public   class PostUpdate : IPlugin
    {
        static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
             tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                tracingService.Trace("1");
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == hil_ReturnLine.EntityLogicalName.ToLower()
                    && context.MessageName.ToUpper() == "UPDATE")
                {

                    tracingService.Trace("2");
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("2.1"+entity.LogicalName+" "+entity.Id);
                    //When inspection person click inspection done
                    onInspectionPersonApproval(entity, service, tracingService);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ReturnLine.PostUpdate.Execute" + ex.Message);
            }
        }

        public static void onInspectionPersonApproval(Entity entity, IOrganizationService service, ITracingService tracing)
        {
            try
            {
                // tracingService.Trace("3");
                //If approval status is done
                //if (entity.Contains("hil_approverstatus") )
                    if(entity.GetAttributeValue<OptionSetValue>("hil_approverstatus")!=null)
                {
                  //  tracingService.Trace("4");
                    if (entity.GetAttributeValue<OptionSetValue>("hil_approverstatus").Value == 1)//Done
                    {
                    //    tracingService.Trace("5");
                        hil_ReturnLine ReturnLine = (hil_ReturnLine)service.Retrieve(hil_ReturnLine.EntityLogicalName, entity.Id, new ColumnSet(true));
                        using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                        {
                            if (ReturnLine.hil_ApprovedQuantity != null)//Approved qty
                            {
                                Guid fsParentAccountId = Guid.Empty; //get parent account id only if it is super franchisee
                                Guid fsParentAccountOwnerId = Guid.Empty;//get parent account owner id only if it is super franchisee
                                int iQuantity = ReturnLine.hil_ApprovedQuantity.Value;
                                //int iDifferenceQty = ReturnLine.hil_Quantity.Value - ReturnLine.hil_ApprovedQuantity.Value;
                                Guid fsPartId = Guid.Empty;
                                Guid fsAccountId = Guid.Empty;
                                Guid fsOwnerId = Guid.Empty;
                                OptionSetValue opInventoryType = null;
                                OptionSetValue opReturnType = null;
                                Guid fsReturnLineId = ReturnLine.Id;
                                #region GetParentFranchiseeAccountDetails
                                var obj = from _Account in orgContext.CreateQuery<Account>()
                                          join _ParentAccount in orgContext.CreateQuery<Account>()
                                          on _Account.ParentAccountId.Id equals _ParentAccount.AccountId.Value
                                          where _Account.AccountId == ReturnLine.hil_Account.Id
                                          // && _ParentAccount.CustomerTypeCode.Equals(Account_CustomerTypeCode.SuperFranchisee)
                                          select new
                                          {
                                              _ParentAccount.CustomerTypeCode
                                              ,
                                              _ParentAccount.OwnerId
                                              ,
                                              _ParentAccount.AccountId
                                          };
                                foreach (var iobj in obj)
                                {
                                    if (iobj.CustomerTypeCode != null)
                                    {
                                        if (iobj.CustomerTypeCode.Value == ((int)Account_CustomerTypeCode.SuperFranchisee))
                                        {
                                            fsParentAccountId = iobj.AccountId.Value;
                                            fsParentAccountOwnerId = iobj.OwnerId.Id;
                                        }
                                    }
                                }
                                #endregion

                                if (ReturnLine.hil_ProductCode != null) fsPartId = ReturnLine.hil_ProductCode.Id;
                                if (ReturnLine.hil_Account != null) fsAccountId = ReturnLine.hil_Account.Id;
                                if (ReturnLine.OwnerId != null) fsOwnerId = ReturnLine.OwnerId.Id;
                                if (ReturnLine.hil_WarrantyStatus != null) opInventoryType = ReturnLine.hil_WarrantyStatus;
                                if (ReturnLine.hil_ReturnType != null) opReturnType = ReturnLine.hil_ReturnType;

                                if (opReturnType.Value == 3)//Good
                                {
                                    //Reduce Good inventory of this account
                                    HelperInvJournal.CreateInvJournal(fsPartId, fsAccountId, fsOwnerId, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, fsReturnLineId, service, tracing, -iQuantity);
                                    if (fsParentAccountId != Guid.Empty)
                                    {
                                        HelperInvJournal.CreateInvJournal(fsPartId, fsParentAccountId, fsParentAccountOwnerId, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, fsReturnLineId, service, tracing, iQuantity);
                                    }
                                }
                                else if (opReturnType.Value == 2)//Def
                                {
                                    //Reduce Def inventory of this account
                                    HelperInvJournal.CreateInvJournalDefective(fsPartId, fsAccountId, fsOwnerId, opInventoryType, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, fsReturnLineId, service, tracing, - iQuantity);
                                    if (fsParentAccountId != Guid.Empty)
                                    {
                                        HelperInvJournal.CreateInvJournalDefective(fsPartId, fsParentAccountId, fsParentAccountOwnerId, opInventoryType, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, fsReturnLineId, service, tracing, iQuantity);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ReturnLine.PostUpdate.onInspectionPersonApproval" + ex.Message);
            }
        }

    }
}
