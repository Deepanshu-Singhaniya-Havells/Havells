using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Linq;

// Microsoft Dynamics CRM namespace(s)

namespace Havells_Plugin.ReturnHeader
{
    public  class PostUpdate : IPlugin
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
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == hil_returnheader.EntityLogicalName.ToLower()
                    && context.MessageName.ToUpper() == "UPDATE")
                {
                   Entity entity=(Entity)context.InputParameters["Target"];
                    onSubmitofReturnHeader(entity, service, tracingService);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ReturnHeader.PostUpdate.Execute" + ex.Message);
            }
        }

        public static void onSubmitofReturnHeader(Entity entity, IOrganizationService service, ITracingService iTrace)
        {
            try
            {
                //tracingService.Trace("1");
                hil_returnheader ReturnHeader = entity.ToEntity<hil_returnheader>();

                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    //tracingService.Trace("2");
                    if (ReturnHeader.statuscode != null && ReturnHeader.statuscode.Value == 910590000)//Submitted
                    {
                        
                        #region GetCustomerType
                        OptionSetValue opThisAccountType = null;
                        var obj = from _ReturnHeader in orgContext.CreateQuery<hil_returnheader>()
                                  join _Account in orgContext.CreateQuery<Account>()
                                  on _ReturnHeader.hil_Account.Id equals _Account.AccountId.Value
                                  where _ReturnHeader.hil_returnheaderId == ReturnHeader.hil_returnheaderId
                                  select new
                                  {
                                      _Account.CustomerTypeCode
                                  };
                        foreach (var iobj in obj)
                        {
                            if (iobj.CustomerTypeCode != null)
                                opThisAccountType = iobj.CustomerTypeCode;
                        }
                        #endregion

                        //tracingService.Trace("4");

                        if (opThisAccountType != null)
                        {
                            if (opThisAccountType.Value == ((int)Account_CustomerTypeCode.DirectEngineer))//DE
                            {
                                //tracingService.Trace("5");
                                reduceInventory(entity.Id, service, iTrace);
                            }
                            else if (opThisAccountType.Value == ((int)Account_CustomerTypeCode.Franchisee) || opThisAccountType.Value == ((int)Account_CustomerTypeCode.SuperFranchisee))//not DE then inspection
                            {
                                ShareReturnLinetoInspectionPerson(entity.Id, service);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ReturnHeader.PostUpdate.onSubmitofReturnLine" + ex.Message);
            }
        }

        public static void onSubmitofReturnHeader1(Entity entity, IOrganizationService service, ITracingService iTrace)
        {
            try
            {
                tracingService.Trace("1");
                hil_returnheader ReturnHeader = entity.ToEntity<hil_returnheader>();

                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    tracingService.Trace("2");
                    if (ReturnHeader.statuscode != null && ReturnHeader.statuscode.Value == 910590000)//Submitted
                    {
                        ReturnHeader = (hil_returnheader)service.Retrieve(ReturnHeader.LogicalName, ReturnHeader.Id, new ColumnSet(true));
                        tracingService.Trace("3");

                        int opThisAccountTypeValue = -1;
                      //  opThisAccountTypeValue = 9;
                        //#region GetCustomerType

                        tracingService.Trace("3.0");
                        Entity enAccount = service.Retrieve(Account.EntityLogicalName, ReturnHeader.hil_Account.Id, new ColumnSet("customertypecode"));
                        if (enAccount.Contains("customertypecode"))
                        {
                            opThisAccountTypeValue = enAccount.GetAttributeValue<OptionSetValue>("customertypecode").Value;
                        }

                        //#endregion
                        tracingService.Trace("4");

                        ////If it's  direct engineer then reduce inventory on submit
                        if (opThisAccountTypeValue == ((int)Account_CustomerTypeCode.DirectEngineer))//DE
                        {
                            tracingService.Trace("5");
                            reduceInventory(entity.Id, service, iTrace);
                        }
                        else if (opThisAccountTypeValue != -1)//not DE then inspection
                        {
                            ShareReturnLinetoInspectionPerson(entity.Id, service);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ReturnHeader.PostUpdate.onSubmitofReturnLine" + ex.Message);
            }
        }

        public static void ShareReturnLinetoInspectionPerson(Guid fsReturnHeaderId, IOrganizationService service)
        {
            try
            {
                tracingService.Trace("s1");
                hil_returnheader ReturnHeader = (hil_returnheader)service.Retrieve(hil_returnheader.EntityLogicalName, fsReturnHeaderId, new ColumnSet(true));

                tracingService.Trace("s2");
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {

                    tracingService.Trace("s3");
                    Guid fsBSH = HelperAccount.GetBSHOwnerIdofAccount(ReturnHeader.hil_Account.Id, service);
                    Guid fsInspectionPerson = Guid.Empty;
                    tracingService.Trace("s4 BSH " + fsBSH);
                    #region  GetInspectionPerson
                    var obj1 = from _User in orgContext.CreateQuery<SystemUser>()
                               join _Position in orgContext.CreateQuery<Position>() on _User.PositionId.Id equals _Position.Id
                               where _User.ParentSystemUserId.Id == fsBSH
                               && _Position.Name == "Inspection Person"
                               select new
                               {
                                   _User.SystemUserId
                               };
                    foreach (var iobj1 in obj1)
                    {
                        fsInspectionPerson = iobj1.SystemUserId.Value;
                    }
                    #endregion

                    tracingService.Trace("s5 Inspection Guid= " + fsInspectionPerson);
                    var obj = from _ReturnLine in orgContext.CreateQuery<hil_ReturnLine>()
                              where _ReturnLine.hil_ReturnHeader.Id == fsReturnHeaderId
                              select new
                              {
                                  _ReturnLine.hil_Quantity,
                                  _ReturnLine.hil_ProductCode,
                                  _ReturnLine.hil_Account,
                                  _ReturnLine.OwnerId,
                                  _ReturnLine.hil_ReturnType,
                                  _ReturnLine.hil_WarrantyStatus,
                                  _ReturnLine.Id
                              };
                    foreach (var iobj in obj)
                    {
                        if (iobj.hil_ReturnType.Value == 2 || iobj.hil_ReturnType.Value == 3)//2 3 def or fresh
                        {


                            //add record to access team with this inspection person
                            if (fsInspectionPerson != Guid.Empty)
                            {
                                EntityReference enRef = new EntityReference(hil_ReturnLine.EntityLogicalName, iobj.Id);
                                Helper.SharetoAccessTeam(enRef, fsInspectionPerson, "Return Line Access Team Template", service);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ReturnHeader.PostUpdate.ShareReturnLinetoInspectionPerson" + ex.Message);
            }

        }
        public static void reduceInventory(Guid fsReturnHeaderId, IOrganizationService service, ITracingService iTrace)
        {
            try
            {
                tracingService.Trace("6");
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    tracingService.Trace("7");
                    var obj = from _ReturnLine in orgContext.CreateQuery<hil_ReturnLine>()
                              where _ReturnLine.hil_ReturnHeader.Id == fsReturnHeaderId
                              select new
                              {
                                  _ReturnLine.hil_Quantity
                                  ,_ReturnLine.hil_ProductCode
                                  ,_ReturnLine.hil_Account
                                  ,_ReturnLine.OwnerId
                                  ,_ReturnLine.hil_ReturnType
                                  ,_ReturnLine.hil_WarrantyStatus
                                  ,_ReturnLine.Id
                              };
                    foreach (var iobj in obj)
                    {
                        tracingService.Trace("8");
                        if (iobj.hil_ReturnType.Value == 2 || iobj.hil_ReturnType.Value == 3)//2 3 def or fresh
                        {
                            Guid fsPart = Guid.Empty;
                            Guid fsAccount = Guid.Empty;
                            Guid fsOwner = Guid.Empty;
                            OptionSetValue opInventoryType = null;
                            Int32 iQuantity = 0;
                            Guid fsReturnLineId = Guid.Empty;

                            tracingService.Trace("9");

                            if (iobj.hil_ProductCode != null) fsPart = iobj.hil_ProductCode.Id;
                            if (iobj.hil_Account != null) fsAccount = iobj.hil_Account.Id;
                            if (iobj.OwnerId != null) fsOwner= iobj.OwnerId.Id;
                            if (iobj.hil_WarrantyStatus != null) opInventoryType = iobj.hil_WarrantyStatus;
                            if (iobj.hil_Quantity!= null) iQuantity = iobj.hil_Quantity.Value;

                            tracingService.Trace("10");
                            fsReturnLineId = iobj.Id;
                            //Create inv journal for grn line's account
                            if (iobj.hil_ReturnType.Value == 2)//Defecative Part Return
                            {
                                HelperInvJournal.CreateInvJournalDefective(fsPart, fsAccount, fsOwner, opInventoryType, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, fsReturnLineId, service, iTrace , - iQuantity);
                            }
                            if (iobj.hil_ReturnType.Value == 3)//Fresh part return
                            {

                                tracingService.Trace("11");
                                HelperInvJournal.CreateInvJournal(fsPart, fsAccount, fsOwner, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, fsReturnLineId, service, iTrace ,- iQuantity);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ReturnHeader.PostUpdate.reduceInventory" + ex.Message);
            }
            
            }
    }
}
