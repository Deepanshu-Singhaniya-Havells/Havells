using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System.Linq;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;

namespace Havells_Plugin.ProductRequest
{
   public class PostCreate : IPlugin
    {
        static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == hil_productrequest.EntityLogicalName.ToLower()
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    AssignPONextLevel(entity, service);
                    OnPOLineCreate(entity, service);
                    //Change 27 Aug 2018
                    TagPOHeader(entity, service);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ProductRequest.PreCreate.Execute" + ex.Message);
            }
        }
        public static void OnPOLineCreate(Entity entity, IOrganizationService service)
        {
            hil_productrequest iPOReq = entity.ToEntity<hil_productrequest>();
            hil_productrequest iPOUpdate = new hil_productrequest();
            string iChannel = string.Empty;
            Account iDSEFranch = new Account();
            if(iPOReq.hil_SuperFranchiseeDSEName != null)
            {
                iDSEFranch = (Account)service.Retrieve(Account.EntityLogicalName, iPOReq.hil_SuperFranchiseeDSEName.Id, new ColumnSet("customertypecode"));
                if (iPOReq.hil_PRType != null && iPOReq.hil_PRType.Value == 910590001)
                {
                    iChannel = HelperPO.getDistributionChannel(iPOReq.hil_WarrantyStatus, iDSEFranch.CustomerTypeCode, service);
                    iPOUpdate.Id = iPOReq.Id;
                    iPOUpdate.hil_DistributionChannel = iChannel;
                    iPOUpdate.hil_Category = iDSEFranch.CustomerTypeCode;
                    service.Update(iPOUpdate);
                }
            }
        }
        public static void TagPOHeader(Entity entity,IOrganizationService service)
        {
            try
            {
                hil_productrequest poReq = entity.ToEntity<hil_productrequest>();
                if (poReq.hil_PRType.Value == ((int)hil_PRType.Emergency))
                {
                    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                    {
                        Guid fsPRHeader = Guid.Empty;

                        #region GetExistingAsset
                        var obj = from _POHeader in orgContext.CreateQuery<hil_productrequestheader>()
                                  where _POHeader.hil_Job.Id == poReq.hil_Job.Id
                                  && _POHeader.hil_WarrantyStatus.Value==poReq.hil_WarrantyStatus.Value
                                  select new
                                  {
                                      _POHeader.hil_productrequestheaderId
                                  };
                        foreach (var iobj in obj)
                        {
                            if (iobj.hil_productrequestheaderId != null)
                            {
                                fsPRHeader = iobj.hil_productrequestheaderId.Value;
                            }
                        }
                        #endregion
                        if (fsPRHeader == Guid.Empty)
                        {
                          //  tracingService.Trace("1 "+poReq.OwnerId.Name + poReq.OwnerId.Id);
                          //  tracingService.Trace("2 " +poReq.hil_SuperFranchiseeDSEName.Name + poReq.hil_SuperFranchiseeDSEName.Id);
                            fsPRHeader =HelperPO.CreatePOHeader(service, poReq.hil_Job, poReq.OwnerId, poReq.hil_SuperFranchiseeDSEName,poReq.hil_Division, poReq.hil_DivisionSapCode, poReq.hil_PRType.Value, poReq.hil_CustomerSAPCode, poReq.hil_WarrantyStatus);
                        }
                        #region UpdatePOHeaderonpoReq
                        hil_productrequest UppoReq = new hil_productrequest();
                        UppoReq.hil_productrequestId = poReq.hil_productrequestId;
                        UppoReq.hil_PRHeader = new EntityReference(hil_productrequestheader.EntityLogicalName, fsPRHeader);
                        service.Update(UppoReq);
                        #endregion
                    }

                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ProductRequest.PreCreate.TagPOHeader" + ex.Message);
            }
        }
        public static void AssignPONextLevel(Entity entity, IOrganizationService service)
        {
            try
            {
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    hil_productrequest PO = entity.ToEntity<hil_productrequest>();
                    if (PO.hil_PRType.Value==((int)hil_PRType.Emergency)) //
                    {
                        PO.hil_SyncStatus = new OptionSetValue(1);//pending for sync
                        service.Update(PO);
                    }
                    else if (PO.hil_PRType.Value==((int)hil_PRType.MSL))
                    {
                       //if (PO.hil_Category != null && (PO.hil_Category.Equals(hil_Category.DirectEngineer)))
                        {
                            Guid fsUserId =HelperAccount.GetBSHOwnerIdofAccount(PO.hil_SuperFranchiseeDSEName.Id,service);
                            if (fsUserId != Guid.Empty)
                            {
                                Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, fsUserId, PO.Id, service);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ProductRequest.PostCreate.AssignPONextLevel" + ex.Message);
            }
        }
        public static Guid GetNextLevelUser(Guid fsPO, IOrganizationService service)
        {
            Guid result = Guid.Empty;
            try
            {
                
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    var obj = from _PO in orgContext.CreateQuery<hil_productrequest>()
                              join _Account in orgContext.CreateQuery<Account>()
                              on _PO.hil_SuperFranchiseeDSEName.Id equals _Account.Id
                              join _ParentAccount in orgContext.CreateQuery<Account>()
                              on _Account.ParentAccountId.Id equals _ParentAccount.AccountId
                              where _PO.hil_productrequestId.Value == fsPO
                              && _ParentAccount.CustomerTypeCode.Equals(Account_CustomerTypeCode.SuperFranchisee)
                              select new
                              {
                                  _ParentAccount.OwnerId
                              };
                    foreach (var iobj in obj)
                    {
                        result = iobj.OwnerId.Id;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ProductRequest.PostCreate.GetNextLevelUser" + ex.Message);
            }
            return result;

        }
        public static Guid GetBSH(Guid fsPO, IOrganizationService service)
        {
            Guid result = Guid.Empty;
            try
            {
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    var obj = from _PO in orgContext.CreateQuery<hil_productrequest>()
                              join _Account in orgContext.CreateQuery<Account>()
                              on _PO.hil_SuperFranchiseeDSEName.Id equals _Account.Id
                              join _ParentAccount in orgContext.CreateQuery<Account>()
                              on _Account.ParentAccountId.Id equals _ParentAccount.AccountId
                              where _PO.Id == fsPO
                              && _ParentAccount.CustomerTypeCode.Equals(Account_CustomerTypeCode.Branch)
                              select new
                              {
                                  _ParentAccount.OwnerId
                              };
                    foreach (var iobj in obj)
                    {
                        result = iobj.OwnerId.Id;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ProductRequest.PostCreate.GetNextLevelUser" + ex.Message);
            }
            return result;
        }
    }
}
