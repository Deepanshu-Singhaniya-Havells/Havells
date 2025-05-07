using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Havells_Plugin;
using System.Linq;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin.WorkOrderIncident
{
    public class PostUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == msdyn_workorderincident.EntityLogicalName
                    && context.MessageName.ToUpper() == "UPDATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("1");
                    //Set warranty Status
                    HelperWOProduct.SetWarrantyStatus(entity, service);
                    tracingService.Trace("2");
                    //Job class half day-full day:- Vishal
                    //Common.getandSetJobClass(entity, context, service);
                    //Set model lookup and model name:- Vishal
                    Guid Model = Common.SetModeldetail(entity, service);
                    tracingService.Trace("3");
                    if (entity.Attributes.Contains("hil_productreplacement"))
                    {
                        msdyn_workorderincident _entWrkInc = entity.ToEntity<msdyn_workorderincident>();
                        if (_entWrkInc.hil_productreplacement.Value == 1)
                        {
                            ProductReplacementProcessInitiation(service, entity.Id, tracingService);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Havells_Plugin.WorkOrderIncident.PostUpdate" + ex.Message);
            }
        }
        #region Delete Existing Products
        public static void DeleteExisting(IOrganizationService service, Guid IncId)
        {
            QueryByAttribute Query = new QueryByAttribute(msdyn_workorderproduct.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.AddAttributeValue("msdyn_workorderincident", IncId);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count >= 1)
            {
                foreach (msdyn_workorderproduct Pdt in Found.Entities)
                {
                    service.Delete(msdyn_workorderproduct.EntityLogicalName, Pdt.Id);
                }
            }
        }
        #endregion
        #region Product Replacement Process
        //public static void InitiateProductReplacement(IOrganizationService service, Guid _entWrkInc)
        //{
        //    msdyn_workorderincident Inc = (msdyn_workorderincident)service.Retrieve(msdyn_workorderincident.EntityLogicalName, _entWrkInc, new ColumnSet(true));
        //    if(Inc.msdyn_CustomerAsset != null && Inc.msdyn_WorkOrder != null)
        //    {
        //        msdyn_customerasset Asst = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, Inc.msdyn_CustomerAsset.Id, new ColumnSet(true));
        //        msdyn_workorder Wo = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, Inc.msdyn_WorkOrder.Id, new ColumnSet(false));
        //        Guid SubStatus = Helper.GetGuidbyName(msdyn_workordersubstatus.EntityLogicalName, "msdyn_name", "Product Replacement in Progress", service);
        //        if(SubStatus != Guid.Empty)
        //        {
        //            Wo.msdyn_SubStatus = new EntityReference(msdyn_workordersubstatus.EntityLogicalName, SubStatus);
        //            service.Update(Wo);
        //        }
        //        if (Asst.hil_ProductSubcategory != null)
        //        {
        //            Product Pdt = (Product)service.Retrieve(Product.EntityLogicalName, Asst.hil_ProductSubcategory.Id, new ColumnSet("hil_serviceable"));
        //            if(Pdt.hil_serviceable != null)
        //            {
        //                ProductReplacement(service, Inc, Asst, Pdt.hil_serviceable.Value);
        //            }
        //        }
        //    }
        //}
        //public static void ProductReplacement(IOrganizationService service, msdyn_workorderincident Wo, msdyn_customerasset Asst, int Servicable)
        //{
        //    if (Servicable == 2)
        //    {
        //        if (Wo.hil_warrantystatus.Value == 1)
        //        {
        //            CreateProductRequestRecord(service, Wo, Asst);//, Servicable);
        //        }
        //        else if (Wo.hil_warrantystatus.Value == 2 || Wo.hil_warrantystatus.Value == 3)
        //        {
        //            CreateProductRequestRecord(service, Wo, Asst);//, Servicable);
        //        }
        //    }
        //    //else if(Servicable == 1)
        //    //{
        //    //    if (Wo.hil_warrantystatus.Value == 1)
        //    //    {
        //    //        CreateProductRequestRecord(service, Wo, Asst);//, Servicable);
        //    //    }
        //    //    else if (Wo.hil_warrantystatus.Value == 2 || Wo.hil_warrantystatus.Value == 3)
        //    //    {
        //    //        CreateProductRequestRecord(service, Wo, Asst);//, Servicable);
        //    //    }
        //    //}
        //}
        //public static void ExecuteConditionalAssign(IOrganizationService service, Guid Category, Guid ProdReq, int Diff, int WarrantyStatus)
        //{
        //    int agediffernt = 0;
        //    if (Diff < 180)
        //    {
        //        agediffernt = 1;
        //    }
        //    else if (Diff >= 180 && Diff <= 365)
        //    {
        //        agediffernt = 2;
        //    }
        //    else if (Diff < 365)
        //    {
        //        agediffernt = 3;
        //    }
        //    QueryByAttribute Query = new QueryByAttribute(hil_integrationconfiguration.EntityLogicalName);
        //    Query.ColumnSet = new ColumnSet(true);
        //    Query.AddAttributeValue("hil_division", Category);
        //    Query.AddAttributeValue("new_warrantystatus", WarrantyStatus);
        //    Query.AddAttributeValue("hil_ageofproduct", agediffernt);
        //    EntityCollection Found = service.RetrieveMultiple(Query);
        //    if (Found.Entities.Count >= 1)
        //    {
        //        foreach (hil_integrationconfiguration Conf in Found.Entities)
        //        {
        //            if(Conf.hil_approvername != null)
        //            {
        //                Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, Conf.hil_approvername.Id, ProdReq, service);
        //            }
        //        }
        //    }
        //}
        //public static void AssignToLevelOneApprover(IOrganizationService service, Guid ProdCat, Guid ProdReq)
        //{
        //    QueryByAttribute Query = new QueryByAttribute(hil_integrationconfiguration.EntityLogicalName);
        //    Query.ColumnSet = new ColumnSet(true);
        //    Query.AddAttributeValue("hil_division", ProdCat);
        //    Query.AddAttributeValue("new_warrantystatus", 2);
        //    EntityCollection Found = service.RetrieveMultiple(Query);
        //    if(Found.Entities.Count >= 1)
        //    {
        //        foreach(hil_integrationconfiguration Conf in Found.Entities)
        //        {
        //            if(Conf.hil_levelofapprover != null && Conf.hil_approvername != null)
        //            {
        //                if(Conf.hil_levelofapprover.Value == 1)
        //                {
        //                    Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, Conf.hil_approvername.Id, ProdReq, service);
        //                }
        //            }
        //        }
        //    }
        //}
        #endregion
        #region Product Replacement 
        //public static void ProductReplacementProcessInitiation(IOrganizationService service, Guid IncId)
        //{
        //    msdyn_workorderincident Incident = (msdyn_workorderincident)service.Retrieve(msdyn_workorderincident.EntityLogicalName, IncId, new ColumnSet(true));
        //    if (Incident.msdyn_CustomerAsset != null && Incident.msdyn_WorkOrder != null)
        //    {
        //        msdyn_customerasset Asst = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, Incident.msdyn_CustomerAsset.Id, new ColumnSet(true));
        //        msdyn_workorder Wo = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, Incident.msdyn_WorkOrder.Id, new ColumnSet(false));
        //        Guid SubStatus = Helper.GetGuidbyName(msdyn_workordersubstatus.EntityLogicalName, "msdyn_name", "Product Replacement in Progress", service);
        //        if (SubStatus != Guid.Empty)
        //        {
        //            Wo.msdyn_SubStatus = new EntityReference(msdyn_workordersubstatus.EntityLogicalName, SubStatus);
        //            service.Update(Wo);
        //        }
        //        if (Asst.hil_ProductSubcategory != null)
        //        {
        //            Product Pdt = (Product)service.Retrieve(Product.EntityLogicalName, Asst.hil_ProductSubcategory.Id, new ColumnSet("hil_serviceable"));
        //            //if (Pdt.hil_serviceable != null)
        //            //{
        //            RouteAsPerWarrantyProductReplacement(service, Incident, Asst);//, Pdt.hil_serviceable.Value);
        //            //}
        //        }
        //    }
        //}
        //public static void RouteAsPerWarrantyProductReplacement(IOrganizationService service, msdyn_workorderincident Wo, msdyn_customerasset Asst)//, int Servicable)
        //{
        //    //if (Servicable == 2)
        //    //{
        //        if (Wo.hil_warrantystatus.Value == 1)
        //        {
        //        CreateProductRequestRecord(service, Wo, Asst);//, Servicable);
        //        }
        //        else if (Wo.hil_warrantystatus.Value == 2 || Wo.hil_warrantystatus.Value == 3)
        //        {
        //        CreateProductRequestRecord(service, Wo, Asst);//, Servicable);
        //        }
        //    //}
        //    //else if (Servicable == 1)
        //    //{
        //        //if (Wo.hil_warrantystatus.Value == 1)
        //        //{
        //        //CreateProductRequestRecord(service, Wo, Asst);//, Servicable);
        //        //}
        //        //else if (Wo.hil_warrantystatus.Value == 2 || Wo.hil_warrantystatus.Value == 3)
        //        //{
        //        //CreateProductRequestRecord(service, Wo, Asst);//, Servicable);
        //        //}
        //    //}
        //}
        //public static void CreateProductRequestRecord(IOrganizationService service, msdyn_workorderincident Incident, msdyn_customerasset Asst)//, int Servicable)
        //{
        //    OptionSetValue PrType = new OptionSetValue(910590002);
        //    Guid HeaderId = new Guid();
        //    Guid _poLine = new Guid();
        //    Account RelatedFranchiseeAcc = new Account();
        //    msdyn_workorder _wrkOd = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, Incident.msdyn_WorkOrder.Id, new ColumnSet("hil_owneraccount"));
        //    Product Division = (Product)service.Retrieve(Product.EntityLogicalName, Incident.hil_ProductCategory.Id, new ColumnSet("hil_sapcode", "name"));
        //    if(_wrkOd.hil_OwnerAccount != null)
        //    {
        //        RelatedFranchiseeAcc = (Account)service.Retrieve(Account.EntityLogicalName, _wrkOd.hil_OwnerAccount.Id, new ColumnSet("ownerid", "customertypecode", "hil_inwarrantycustomersapcode", "hil_outwarrantycustomersapcode", "name"));
        //        if (RelatedFranchiseeAcc.hil_InWarrantyCustomerSAPCode != null && RelatedFranchiseeAcc.hil_OutWarrantyCustomerSAPCode != null)
        //        {
        //            if(Division.hil_SAPCode != null)
        //            {
        //                if(Incident.hil_warrantystatus.Value == 1)
        //                {
        //                    HeaderId = HelperPO.CreatePOHeader(service, Incident.msdyn_WorkOrder, RelatedFranchiseeAcc.OwnerId, _wrkOd.hil_OwnerAccount, Incident.hil_ProductCategory, Division.hil_SAPCode, PrType.Value, RelatedFranchiseeAcc.hil_InWarrantyCustomerSAPCode, Incident.hil_warrantystatus);
        //                    if(HeaderId != Guid.Empty)
        //                    {
        //                        string DistributionChannel = HelperPO.getDistributionChannel(Incident.hil_warrantystatus, RelatedFranchiseeAcc.CustomerTypeCode, service);
        //                        _poLine = HelperPO.CreatePO(service, RelatedFranchiseeAcc.OwnerId.Id, Incident.msdyn_WorkOrder.Id, HeaderId, 1, _wrkOd.hil_OwnerAccount.Id, Asst.msdyn_Product.Id, PrType.Value, RelatedFranchiseeAcc.CustomerTypeCode, RelatedFranchiseeAcc.hil_InWarrantyCustomerSAPCode, DistributionChannel, Incident.hil_warrantystatus);
        //                        if(_poLine != Guid.Empty)
        //                        {
        //                            hil_productrequest _thisPO = (hil_productrequest)service.Retrieve(hil_productrequest.EntityLogicalName, _poLine, new ColumnSet(false));
        //                            hil_productrequestheader _thisHeader = (hil_productrequestheader)service.Retrieve(hil_productrequestheader.EntityLogicalName, HeaderId, new ColumnSet(false));
        //                            //if (Servicable == 1)//Yes
        //                            //{
        //                                _thisPO.statuscode = new OptionSetValue(910590004);
        //                                service.Update(_thisPO);
        //                                _thisHeader.statuscode = new OptionSetValue(910590004);
        //                                service.Update(_thisHeader);
        //                                ExecuteAssign(service, Asst, _thisPO, Incident.hil_warrantystatus.Value, _thisHeader);
        //                            //, Servicable
        //                            //}
        //                            //else if (Servicable == 2)//No
        //                            //{
        //                            //    _thisPO.statuscode = new OptionSetValue(910590003);
        //                            //    service.Update(_thisPO);
        //                            //    _thisHeader.statuscode = new OptionSetValue(910590003);
        //                            //    service.Update(_thisHeader);

        //                            //}
        //                        }
        //                    }
        //                }
        //                else if(Incident.hil_warrantystatus.Value == 2 || Incident.hil_warrantystatus.Value == 3)
        //                {
        //                    HeaderId = HelperPO.CreatePOHeader(service, Incident.msdyn_WorkOrder, RelatedFranchiseeAcc.OwnerId, _wrkOd.hil_OwnerAccount, Incident.hil_ProductCategory, Division.hil_SAPCode, PrType.Value, RelatedFranchiseeAcc.hil_OutWarrantyCustomerSAPCode, Incident.hil_warrantystatus);
        //                    if (HeaderId != Guid.Empty)
        //                    {
        //                        string DistributionChannel = HelperPO.getDistributionChannel(Incident.hil_warrantystatus, RelatedFranchiseeAcc.CustomerTypeCode, service);
        //                        _poLine = HelperPO.CreatePO(service, RelatedFranchiseeAcc.OwnerId.Id, Incident.msdyn_WorkOrder.Id, HeaderId, 1, _wrkOd.hil_OwnerAccount.Id, Asst.msdyn_Product.Id, PrType.Value, RelatedFranchiseeAcc.CustomerTypeCode, RelatedFranchiseeAcc.hil_OutWarrantyCustomerSAPCode, DistributionChannel, Incident.hil_warrantystatus);
        //                        if (_poLine != Guid.Empty)
        //                        {
        //                            hil_productrequest _thisPO = (hil_productrequest)service.Retrieve(hil_productrequest.EntityLogicalName, _poLine, new ColumnSet(false));
        //                            hil_productrequestheader _thisHeader = (hil_productrequestheader)service.Retrieve(hil_productrequestheader.EntityLogicalName, HeaderId, new ColumnSet(false));
        //                            //if (Servicable == 1)//Yes
        //                            //{
        //                                _thisPO.statuscode = new OptionSetValue(910590004);
        //                                service.Update(_thisPO);
        //                                _thisHeader.statuscode = new OptionSetValue(910590004);
        //                                service.Update(_thisHeader);
        //                                ExecuteAssign(service, Asst, _thisPO, Incident.hil_warrantystatus.Value, _thisHeader);
        //                            //, Servicable
        //                            //}
        //                            //else if (Servicable == 2)//No
        //                            //{
        //                            //    _thisPO.statuscode = new OptionSetValue(910590004);
        //                            //    service.Update(_thisPO);
        //                            //    _thisHeader.statuscode = new OptionSetValue(910590004);
        //                            //    service.Update(_thisHeader);
        //                            //    ExecuteAssign(service, Asst, _thisPO, Servicable, Incident.hil_warrantystatus.Value, _thisHeader);
        //                            //}
        //                        }
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                throw new InvalidPluginExecutionException("Division SAP Code Blank for Division : " + Division.Name);
        //            }
        //        }
        //        else
        //        {
        //            throw new InvalidPluginExecutionException("Customer SAP Code blank for : " + RelatedFranchiseeAcc.Name);
        //        }
        //    }
        //}
        //public static void ExecuteAssign(IOrganizationService service, msdyn_customerasset Asst, hil_productrequest _thisPO, int WarrantyStatus, hil_productrequestheader _thisHeader)
        //{
        //    int age = 0;
        //    if(WarrantyStatus == 1)
        //    {
        //        if (Asst.hil_InvoiceDate != null)
        //        {
        //            age = DateTime.Now.Subtract(Convert.ToDateTime(Asst.hil_InvoiceDate)).Days;
        //        }
        //        if(age <= 180)
        //        {
        //            QueryExpression Query = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
        //            Query.ColumnSet = new ColumnSet("hil_approvername");
        //            Query.Criteria.AddFilter(LogicalOperator.And);
        //            Query.Criteria.AddCondition("hil_division", ConditionOperator.Equal, Asst.hil_ProductCategory.Id);
        //            Query.Criteria.AddCondition("hil_levelofapprover", ConditionOperator.Equal, 1);
        //            EntityCollection Found = service.RetrieveMultiple(Query);
        //            if (Found.Entities.Count > 0)
        //            {
        //                hil_integrationconfiguration Conf = (hil_integrationconfiguration)Found.Entities[0];
        //                Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, Conf.hil_approvername.Id, _thisPO.Id, service);
        //                Helper.Assign(SystemUser.EntityLogicalName, hil_productrequestheader.EntityLogicalName, Conf.hil_approvername.Id, _thisHeader.Id, service);
        //            }
        //        }
        //        else if (age > 180 && age <= 365)
        //        {
        //            QueryExpression Query = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
        //            Query.ColumnSet = new ColumnSet("hil_approvername");
        //            Query.Criteria.AddFilter(LogicalOperator.And);
        //            Query.Criteria.AddCondition("hil_division", ConditionOperator.Equal, Asst.hil_ProductCategory.Id);
        //            Query.Criteria.AddCondition("hil_levelofapprover", ConditionOperator.Equal, 2);
        //            EntityCollection Found = service.RetrieveMultiple(Query);
        //            if (Found.Entities.Count > 0)
        //            {
        //                hil_integrationconfiguration Conf = (hil_integrationconfiguration)Found.Entities[0];
        //                Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, Conf.hil_approvername.Id, _thisPO.Id, service);
        //                Helper.Assign(SystemUser.EntityLogicalName, hil_productrequestheader.EntityLogicalName, Conf.hil_approvername.Id, _thisHeader.Id, service);
        //            }
        //        }
        //        else if (age > 365)
        //        {
        //            QueryExpression Query = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
        //            Query.ColumnSet = new ColumnSet("hil_approvername");
        //            Query.Criteria.AddFilter(LogicalOperator.And);
        //            Query.Criteria.AddCondition("hil_division", ConditionOperator.Equal, Asst.hil_ProductCategory.Id);
        //            Query.Criteria.AddCondition("hil_levelofapprover", ConditionOperator.Equal, 3);
        //            EntityCollection Found = service.RetrieveMultiple(Query);
        //            if (Found.Entities.Count > 0)
        //            {
        //                hil_integrationconfiguration Conf = (hil_integrationconfiguration)Found.Entities[0];
        //                Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, Conf.hil_approvername.Id, _thisPO.Id, service);
        //                Helper.Assign(SystemUser.EntityLogicalName, hil_productrequestheader.EntityLogicalName, Conf.hil_approvername.Id, _thisHeader.Id, service);
        //            }
        //        }
        //    }
        //    else if(WarrantyStatus == 2 || WarrantyStatus == 3)
        //    {
        //        if (Asst.hil_InvoiceDate != null)
        //        {
        //            age = DateTime.Now.Subtract(Convert.ToDateTime(Asst.hil_InvoiceDate)).Days;
        //        }
        //        if (age <= 180)
        //        {
        //            QueryExpression Query = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
        //            Query.ColumnSet = new ColumnSet("hil_approvername");
        //            Query.Criteria.AddFilter(LogicalOperator.And);
        //            Query.Criteria.AddCondition("hil_division", ConditionOperator.Equal, Asst.hil_ProductCategory.Id);
        //            Query.Criteria.AddCondition("hil_warrantystatus", ConditionOperator.Equal, WarrantyStatus);
        //            Query.Criteria.AddCondition("hil_levelofapprover", ConditionOperator.Equal, 1);
        //            EntityCollection Found = service.RetrieveMultiple(Query);
        //            if (Found.Entities.Count > 0)
        //            {
        //                hil_integrationconfiguration Conf = (hil_integrationconfiguration)Found.Entities[0];
        //                Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, Conf.hil_approvername.Id, _thisPO.Id, service);
        //                Helper.Assign(SystemUser.EntityLogicalName, hil_productrequestheader.EntityLogicalName, Conf.hil_approvername.Id, _thisHeader.Id, service);
        //            }
        //        }
        //        else if (age > 180 && age <= 365)
        //        {
        //            QueryExpression Query = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
        //            Query.ColumnSet = new ColumnSet("hil_approvername");
        //            Query.Criteria.AddFilter(LogicalOperator.And);
        //            Query.Criteria.AddCondition("hil_division", ConditionOperator.Equal, Asst.hil_ProductCategory.Id);
        //            Query.Criteria.AddCondition("hil_warrantystatus", ConditionOperator.Equal, WarrantyStatus);
        //            Query.Criteria.AddCondition("hil_levelofapprover", ConditionOperator.Equal, 2);
        //            EntityCollection Found = service.RetrieveMultiple(Query);
        //            if (Found.Entities.Count > 0)
        //            {
        //                hil_integrationconfiguration Conf = (hil_integrationconfiguration)Found.Entities[0];
        //                Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, Conf.hil_approvername.Id, _thisPO.Id, service);
        //                Helper.Assign(SystemUser.EntityLogicalName, hil_productrequestheader.EntityLogicalName, Conf.hil_approvername.Id, _thisHeader.Id, service);
        //            }
        //        }
        //        else if (age > 365)
        //        {
        //            QueryExpression Query = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
        //            Query.ColumnSet = new ColumnSet("hil_approvername");
        //            Query.Criteria.AddFilter(LogicalOperator.And);
        //            Query.Criteria.AddCondition("hil_division", ConditionOperator.Equal, Asst.hil_ProductCategory.Id);
        //            Query.Criteria.AddCondition("hil_warrantystatus", ConditionOperator.Equal, WarrantyStatus);
        //            Query.Criteria.AddCondition("hil_levelofapprover", ConditionOperator.Equal, 3);
        //            EntityCollection Found = service.RetrieveMultiple(Query);
        //            if (Found.Entities.Count > 0)
        //            {
        //                hil_integrationconfiguration Conf = (hil_integrationconfiguration)Found.Entities[0];
        //                Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, Conf.hil_approvername.Id, _thisPO.Id, service);
        //                Helper.Assign(SystemUser.EntityLogicalName, hil_productrequestheader.EntityLogicalName, Conf.hil_approvername.Id, _thisHeader.Id, service);
        //            }
        //        }
        //    }
        //    //if (Servicable == 2 && (WarrantyStatus == 2 || WarrantyStatus == 3))
        //    //{
        //    //    QueryExpression Query = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
        //    //    Query.ColumnSet = new ColumnSet("hil_approvername");
        //    //    Query.Criteria.AddFilter(LogicalOperator.And);
        //    //    Query.Criteria.AddCondition("hil_division", ConditionOperator.Equal, Asst.hil_ProductCategory.Id);
        //    //    Query.Criteria.AddCondition("hil_warrantystatus", ConditionOperator.Equal, WarrantyStatus);
        //    //    Query.Criteria.AddCondition("hil_levelofapprover", ConditionOperator.Equal, 1);
        //    //    EntityCollection Found = service.RetrieveMultiple(Query);
        //    //    if (Found.Entities.Count > 0)
        //    //    {
        //    //        hil_integrationconfiguration Conf = (hil_integrationconfiguration)Found.Entities[0];
        //    //        Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, Conf.hil_approvername.Id, _thisPO.Id, service);
        //    //        Helper.Assign(SystemUser.EntityLogicalName, hil_productrequestheader.EntityLogicalName, Conf.hil_approvername.Id, _thisHeader.Id, service);
        //    //    }
        //    //}
        //}
        #endregion
        #region Product Replacement Updated 4/12/2018
        public static void ProductReplacementProcessInitiation(IOrganizationService service, Guid IncId, ITracingService tracingService)
        {
            //Product Division = new Product();
            tracingService.Trace("3.6");
            EntityReference iFall = new EntityReference();
            OptionSetValue opBrand = new OptionSetValue();
            msdyn_workorderincident Incident = (msdyn_workorderincident)service.Retrieve(msdyn_workorderincident.EntityLogicalName, IncId, new ColumnSet(true));
            msdyn_workorder Wo = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, Incident.msdyn_WorkOrder.Id, new ColumnSet("hil_brand", "hil_regardingfallback"));
            if (Wo.hil_Brand != null)
            {
                opBrand = Wo.hil_Brand;
            }
            else
            {
                throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXX - Brand Can't be null in Job - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
            }
            if (Wo.Contains("hil_regardingfallback") && Wo.Attributes.Contains("hil_regardingfallback"))
            {
                iFall = Wo.GetAttributeValue<EntityReference>("hil_regardingfallback");
            }
            else
            {
                throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXX - Fall Back Can't be Null in Job - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
            }

            if (Incident.msdyn_CustomerAsset != null && Incident.msdyn_WorkOrder != null && opBrand != null)
            {
                tracingService.Trace("3.7");
                msdyn_customerasset Asst = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, Incident.msdyn_CustomerAsset.Id, new ColumnSet("hil_invoicedate", "msdyn_product", "hil_productsubcategory"));
                //(msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, Incident.msdyn_WorkOrder.Id, new ColumnSet(false));
                Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Product Replacement in Progress", Wo.Id, service);
                tracingService.Trace("3.8");
                if (Asst.hil_ProductSubcategory != null)
                {
                    tracingService.Trace("3.9");
                    Product Pdt = (Product)service.Retrieve(Product.EntityLogicalName, Asst.hil_ProductSubcategory.Id, new ColumnSet("hil_serviceable"));
                    if (Pdt.hil_serviceable != null)
                    {
                        tracingService.Trace("3.10");
                        RouteAsPerWarrantyProductReplacement(service, Incident, Asst, Pdt.hil_serviceable.Value, opBrand.Value, iFall, tracingService);
                    }
                }
            }
        }
        public static void RouteAsPerWarrantyProductReplacement(IOrganizationService service, msdyn_workorderincident Wo, msdyn_customerasset Asst, int Servicable, int Brand, EntityReference iFall, ITracingService tracingService)
        {
            tracingService.Trace("3.23");
            #region Create PO Header and Line
            Guid iHeaderId = new Guid();
            Guid ipoLine = new Guid();
            string iDistributionChannel = string.Empty;
            string iSAPCode = string.Empty;
            double iAge = 0;
            int iAssignLevel = 0;
            Account iFranchisee = new Account();
            OptionSetValue iPoType = new OptionSetValue(910590002);
            msdyn_workorder iJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, Wo.msdyn_WorkOrder.Id, new ColumnSet("hil_owneraccount"));
            tracingService.Trace("3.21");
            Product iDivision = (Product)service.Retrieve(Product.EntityLogicalName, Wo.hil_ProductCategory.Id, new ColumnSet("hil_sapcode", "name"));
            tracingService.Trace("3.22");
            if (iJob.hil_OwnerAccount != null)
            {
                tracingService.Trace("3.1");
                iFranchisee = (Account)service.Retrieve(Account.EntityLogicalName, iJob.hil_OwnerAccount.Id, new ColumnSet("ownerid", "customertypecode", "hil_inwarrantycustomersapcode", "hil_outwarrantycustomersapcode", "name"));
                if (iFranchisee.CustomerTypeCode != null)
                {
                    tracingService.Trace("3.2");
                    iSAPCode = iFranchisee.hil_InWarrantyCustomerSAPCode;
                    tracingService.Trace("3.3");
                    iDistributionChannel = HelperPO.getDistributionChannel(new OptionSetValue(1), iFranchisee.CustomerTypeCode, service);
                    tracingService.Trace("3.4");
                }
            }
            #endregion
            if (Wo.hil_warrantystatus.Value == 2 || Wo.hil_warrantystatus.Value == 3 || Wo.hil_warrantystatus.Value == 4) // OUT WARRANTY or WARRANTY VOID
            {
                tracingService.Trace("3.5");
                iHeaderId = HelperPO.CreatePOHeader(service, Wo.msdyn_WorkOrder, iFranchisee.OwnerId, iJob.hil_OwnerAccount, Wo.hil_ProductCategory, iDivision.hil_SAPCode, iPoType.Value, iFranchisee.hil_InWarrantyCustomerSAPCode, Wo.hil_warrantystatus);
                if (iHeaderId != Guid.Empty)
                {
                    ipoLine = HelperPO.CreatePO(service, iFranchisee.OwnerId.Id, Wo.msdyn_WorkOrder.Id, iHeaderId, 1, iJob.hil_OwnerAccount.Id, Asst.msdyn_Product.Id, iPoType.Value, iFranchisee.CustomerTypeCode, iFranchisee.hil_InWarrantyCustomerSAPCode, iDistributionChannel, Wo.hil_warrantystatus);
                    if (ipoLine != Guid.Empty)
                    {
                        iAssignLevel = 1;
                        hil_productrequestheader Header = new hil_productrequestheader();
                        hil_productrequest Line = new hil_productrequest();
                        Line.Id = ipoLine;
                        Header.Id = iHeaderId;
                        Header.statuscode = new OptionSetValue(910590004);
                        Header["hil_pendingat"] = new OptionSetValue(iAssignLevel);
                        Header["hil_finallevel"] = new OptionSetValue(3);
                        Line.statuscode = new OptionSetValue(910590004);
                        service.Update(Header);
                        service.Update(Line);
                        ExecuteAssign(service, iDivision.Id, ipoLine, iHeaderId, iAssignLevel, iFall);
                    }
                }
            }
            else if (Wo.hil_warrantystatus.Value == 1) // IN WARRANTY
            {
                iHeaderId = HelperPO.CreatePOHeader(service, Wo.msdyn_WorkOrder, iFranchisee.OwnerId, iJob.hil_OwnerAccount, Wo.hil_ProductCategory, iDivision.hil_SAPCode, iPoType.Value, iFranchisee.hil_InWarrantyCustomerSAPCode, Wo.hil_warrantystatus);
                if (iHeaderId != Guid.Empty)
                {
                    ipoLine = HelperPO.CreatePO(service, iFranchisee.OwnerId.Id, Wo.msdyn_WorkOrder.Id, iHeaderId, 1, iJob.hil_OwnerAccount.Id, Asst.msdyn_Product.Id, iPoType.Value, iFranchisee.CustomerTypeCode, iFranchisee.hil_InWarrantyCustomerSAPCode, iDistributionChannel, Wo.hil_warrantystatus);
                    //Brand Enum {1-Havells,2-Lloyd,3-WaterPurifier}
                    if (ipoLine != Guid.Empty && Brand == 1)
                    {
                        if (Servicable == 2)//IN WARRANTY NON-SERVICABLE //Straight Approve, No Approval required
                        {
                            hil_productrequestheader Header = new hil_productrequestheader();
                            hil_productrequest Line = new hil_productrequest();
                            Line.Id = ipoLine;
                            Header.Id = iHeaderId;
                            Header.statuscode = new OptionSetValue(910590003);
                            Header.hil_SyncStatus = new OptionSetValue(1);
                            Line.statuscode = new OptionSetValue(910590003);
                            service.Update(Header);
                            service.Update(Line);
                        }
                        else if (Servicable == 1)// IN WARRANTY SERVICABLE
                        {
                            if (Asst.hil_InvoiceDate != null)
                            {
                                DateTime iPurchaseDate = Convert.ToDateTime(Asst.hil_InvoiceDate);
                                DateTime iSysDate = DateTime.Now;
                                iAge = iSysDate.Subtract(iPurchaseDate).Days / (365.25 / 12);
                                if (iAge < 6) //Assign to Last Level Straight
                                {
                                    iAssignLevel = 1;
                                    hil_productrequestheader Header = new hil_productrequestheader();
                                    hil_productrequest Line = new hil_productrequest();
                                    Line.Id = ipoLine;
                                    Header.Id = iHeaderId;
                                    Header.statuscode = new OptionSetValue(910590004);
                                    Header["hil_pendingat"] = new OptionSetValue(iAssignLevel);
                                    Header["hil_finallevel"] = new OptionSetValue(1);
                                    Line.statuscode = new OptionSetValue(910590004);
                                    service.Update(Header);
                                    service.Update(Line);
                                }
                                else //Assign to Middle Level Straight
                                {
                                    iAssignLevel = 1;
                                    hil_productrequestheader Header = new hil_productrequestheader();
                                    hil_productrequest Line = new hil_productrequest();
                                    Line.Id = ipoLine;
                                    Header.Id = iHeaderId;
                                    Header.statuscode = new OptionSetValue(910590004);
                                    Header["hil_pendingat"] = new OptionSetValue(iAssignLevel);
                                    Header["hil_finallevel"] = new OptionSetValue(2);
                                    Line.statuscode = new OptionSetValue(910590004);
                                    service.Update(Header);
                                    service.Update(Line);
                                }
                                if (iAssignLevel > 0)
                                    ExecuteAssign(service, iDivision.Id, ipoLine, iHeaderId, iAssignLevel, iFall);
                            }
                        }
                    }
                    else if (ipoLine != Guid.Empty && Brand == 2) // LLOYD BRAND
                    {
                        iAssignLevel = 1;
                        hil_productrequestheader Header = new hil_productrequestheader();
                        hil_productrequest Line = new hil_productrequest();
                        Line.Id = ipoLine;
                        Header.Id = iHeaderId;
                        Header.statuscode = new OptionSetValue(910590004);
                        Header["hil_pendingat"] = new OptionSetValue(iAssignLevel);
                        Header["hil_finallevel"] = new OptionSetValue(3);
                        Line.statuscode = new OptionSetValue(910590004);
                        service.Update(Header);
                        service.Update(Line);
                        ExecuteAssign(service, iDivision.Id, ipoLine, iHeaderId, iAssignLevel, iFall);
                    }
                    else if (ipoLine != Guid.Empty && Brand == 3)
                    {
                        iAssignLevel = 1;
                        hil_productrequestheader Header = new hil_productrequestheader();
                        hil_productrequest Line = new hil_productrequest();
                        Line.Id = ipoLine;
                        Header.Id = iHeaderId;
                        Header.statuscode = new OptionSetValue(910590004);
                        Header["hil_pendingat"] = new OptionSetValue(iAssignLevel);
                        Header["hil_finallevel"] = new OptionSetValue(3);
                        Line.statuscode = new OptionSetValue(910590004);
                        service.Update(Header);
                        service.Update(Line);
                        ExecuteAssign(service, iDivision.Id, ipoLine, iHeaderId, iAssignLevel, iFall);
                    }
                }
            }
        }
        public static void ExecuteAssign(IOrganizationService service, Guid iDiv, Guid iPoLineId, Guid iHeaderId, int iAssignLevel, EntityReference iFall)
        {
            if (iFall != null)
            {
                Entity SBUDiv = service.Retrieve("hil_sbubranchmapping", iFall.Id, new ColumnSet(true));
                if (iAssignLevel == 1)
                {
                    if (SBUDiv.Contains("hil_branchheaduser") && SBUDiv.Attributes.Contains("hil_branchheaduser"))
                    {
                        EntityReference Branch = SBUDiv.GetAttributeValue<EntityReference>("hil_branchheaduser");
                        Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, Branch.Id, iPoLineId, service);
                        Helper.Assign(SystemUser.EntityLogicalName, hil_productrequestheader.EntityLogicalName, Branch.Id, iHeaderId, service);
                    }
                }
                if (iAssignLevel == 2)
                {
                    if (SBUDiv.Contains("hil_nph") && SBUDiv.Attributes.Contains("hil_nph"))
                    {
                        EntityReference Branch = SBUDiv.GetAttributeValue<EntityReference>("hil_nph");
                        Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, Branch.Id, iPoLineId, service);
                        Helper.Assign(SystemUser.EntityLogicalName, hil_productrequestheader.EntityLogicalName, Branch.Id, iHeaderId, service);
                    }
                }
                if (iAssignLevel == 3)
                {
                    if (SBUDiv.Contains("hil_nsh") && SBUDiv.Attributes.Contains("hil_nsh"))
                    {
                        EntityReference Branch = SBUDiv.GetAttributeValue<EntityReference>("hil_nsh");
                        Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, Branch.Id, iPoLineId, service);
                        Helper.Assign(SystemUser.EntityLogicalName, hil_productrequestheader.EntityLogicalName, Branch.Id, iHeaderId, service);
                    }
                }
            }
            //QueryExpression Query = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
            //Query.ColumnSet = new ColumnSet(true);
            //Query.Criteria.AddFilter(LogicalOperator.And);
            //Query.Criteria.AddCondition("hil_division", ConditionOperator.Equal, iDiv);
            //Query.Criteria.AddCondition("hil_levelofapprover", ConditionOperator.Equal, iAssignLevel);
            //EntityCollection Found = service.RetrieveMultiple(Query);
            //if (Found.Entities.Count > 0)
            //{
            //    hil_integrationconfiguration Conf = (hil_integrationconfiguration)Found.Entities[0];
            //    Helper.Assign(SystemUser.EntityLogicalName, hil_productrequest.EntityLogicalName, Conf.hil_approvername.Id, iPoLineId, service);
            //    Helper.Assign(SystemUser.EntityLogicalName, hil_productrequestheader.EntityLogicalName, Conf.hil_approvername.Id, iHeaderId, service);
            //}
        }
        #endregion
    }
}