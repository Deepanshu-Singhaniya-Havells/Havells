using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.WorkOrderIncident
{
    public class PostCreate : IPlugin
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
                    && context.PrimaryEntityName.ToLower() == msdyn_workorderincident.EntityLogicalName
                    && context.MessageName.ToUpper() == "CREATE" && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    msdyn_workorderincident _entWrkInc = entity.ToEntity<msdyn_workorderincident>();
                    msdyn_workorder enJob = new msdyn_workorder();
                    tracingService.Trace("-1");
                    enJob.Id = _entWrkInc.msdyn_WorkOrder.Id;
                    enJob = SetWarrantyStatus_JobIncident(entity, service, enJob);
                    enJob = SetWorkStartedOnJob(_entWrkInc, service, enJob);
                    tracingService.Trace("0");
                    service.Update(enJob);
                    tracingService.Trace("1");
                    Guid Model = Common.SetModeldetail(entity, service);
                    tracingService.Trace("2");
                    CheckIfReplacedProduct(service, _entWrkInc);
                    tracingService.Trace("3");

                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrderIncident.PostCreate" + ex.Message);
            }
        }
        #region Check if Replaced Product
        public static void CheckIfReplacedProduct(IOrganizationService service, msdyn_workorderincident iJobLine)
        {
            Guid iOldAsset = new Guid();
            Guid iNewAsset = new Guid();
            QueryExpression Query = new QueryExpression(msdyn_workorderincident.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("msdyn_customerasset");
            Query.Criteria.AddFilter(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_productreplacement", ConditionOperator.Equal, 1);
            Query.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, iJobLine.msdyn_WorkOrder.Id);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                msdyn_workorderincident iReplace = Found.Entities[0].ToEntity<msdyn_workorderincident>();
                iOldAsset = iReplace.msdyn_CustomerAsset.Id;
                iNewAsset = iJobLine.msdyn_CustomerAsset.Id;
                QueryExpression Query1 = new QueryExpression(hil_unitwarranty.EntityLogicalName);
                Query1.ColumnSet = new ColumnSet("hil_customerasset");
                Query1.Criteria.AddFilter(LogicalOperator.And);
                Query1.Criteria.AddCondition("hil_customerasset", ConditionOperator.Equal, iOldAsset);
                EntityCollection Found1 = service.RetrieveMultiple(Query1);
                if(Found1.Entities.Count > 0)
                {
                    foreach(hil_unitwarranty iWarranty in Found1.Entities)
                    {
                        iWarranty.hil_CustomerAsset = new EntityReference(msdyn_customerasset.EntityLogicalName, iNewAsset);
                        service.Update(iWarranty);
                    }
                }
            }
        }
        #endregion
        #region Set Work Started On Job
        public msdyn_workorder SetWorkStartedOnJob(msdyn_workorderincident _WoInc, IOrganizationService service, msdyn_workorder iUpdateWo)
        {
            if(_WoInc.msdyn_WorkOrder != null)
            {
                msdyn_workorder _enJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, iUpdateWo.Id, new ColumnSet("hil_workstarted"));
                if(_enJob.hil_WorkStarted != true)
                {
                    if(_WoInc.Contains("msdyn_customerasset") && _WoInc.msdyn_CustomerAsset != null)
                    {
                        iUpdateWo.msdyn_CustomerAsset = _WoInc.msdyn_CustomerAsset;
                    }
                    iUpdateWo.hil_WorkStarted = true;
                    Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Work Initiated", _enJob.Id, service);
                }
            }
            return iUpdateWo;
        }
        #endregion
        #region Warranty Status
        public msdyn_workorder SetWarrantyStatus_JobIncident(Entity entity, IOrganizationService service, msdyn_workorder upWo)
        {
            try
            {
                msdyn_workorderincident woIncident = (msdyn_workorderincident)service.Retrieve(entity.LogicalName, entity.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("msdyn_customerasset", "msdyn_workorder"));
                if (woIncident.msdyn_CustomerAsset != null)
                {
                    msdyn_workorder iWo = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, woIncident.msdyn_WorkOrder.Id, new ColumnSet("createdon"));
                    msdyn_customerasset enAsset = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, woIncident.msdyn_CustomerAsset.Id, new ColumnSet("hil_warrantytilldate", "hil_warrantystatus", "hil_invoicedate"));
                    msdyn_workorderincident upWOIncident = new msdyn_workorderincident();
                    hil_inventorytype result = hil_inventorytype.NANotfound;
                    OptionSetValue _asWty = new OptionSetValue();
                    if (enAsset.Attributes.Contains("hil_warrantystatus"))
                    {
                        _asWty = (OptionSetValue)enAsset["hil_warrantystatus"];
                    }
                    QueryExpression Query = new QueryExpression(msdyn_workorderincident.EntityLogicalName);
                    Query.ColumnSet = new ColumnSet(new string[] { "hil_warrantystatus" });
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition(new ConditionExpression("msdyn_workorder", ConditionOperator.Equal, woIncident.msdyn_WorkOrder.Id));
                    Query.Criteria.AddCondition(new ConditionExpression("msdyn_workorderincidentid", ConditionOperator.NotEqual, woIncident.Id));
                    Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                    EntityCollection Found = service.RetrieveMultiple(Query);

                    if (_asWty.Value != 3)
                    {
                        if (enAsset.hil_InvoiceDate != null)
                        {
                            //hil_inventorytype result = HelperWarranty.GetWarrantyStatus(service, woIncident.msdyn_CustomerAsset.Id);
                            DateTime dtEndDate = (DateTime)enAsset["hil_warrantytilldate"];//HelperWarranty.GetWarrantyEndDate(service, woIncident.msdyn_CustomerAsset.Id);
                            DateTime today = (DateTime)iWo.CreatedOn;
                            if (dtEndDate != null)
                            {
                                upWOIncident["new_warrantyenddate"] = dtEndDate;
                                if (dtEndDate.Date >= today.Date)
                                {
                                    result = hil_inventorytype.InWarranty;
                                    upWOIncident.hil_warrantystatus = new OptionSetValue(1);//in
                                    if (Found.Entities.Count > 0 && Found.Entities[0].Contains("hil_warrantystatus")
                                        && Found.Entities[0].GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value != 1)
                                    {
                                        throw new InvalidPluginExecutionException("  ***Only " + Found.Entities[0].FormattedValues["hil_warrantystatus"] + " incident(s) are allowed for this job.***  ");
                                    }
                                }
                                else
                                {
                                    if (Found.Entities.Count > 0 && Found.Entities[0].Contains("hil_warrantystatus")
                                           && Found.Entities[0].GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value != 2)
                                    {
                                        throw new InvalidPluginExecutionException("  ***Only " + Found.Entities[0].FormattedValues["hil_warrantystatus"] + " incident(s) are allowed for this job.***  ");
                                    }
                                    upWOIncident.hil_warrantystatus = new OptionSetValue(2);//out
                                }
                            }
                            else
                            {
                                if (Found.Entities.Count > 0 && Found.Entities[0].Contains("hil_warrantystatus")
                                              && Found.Entities[0].GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value != 2)
                                {
                                    throw new InvalidPluginExecutionException("  ***Only In Warranty incident(s) are allowed for this job.***  ");
                                }
                                upWOIncident.hil_warrantystatus = new OptionSetValue(2);//NA not found
                            }
                        }
                        else
                        {
                            if (Found.Entities.Count > 0 && Found.Entities[0].Contains("hil_warrantystatus")
                                              && Found.Entities[0].GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value != 2)
                            {
                                throw new InvalidPluginExecutionException("  ***Only Out Warranty incident(s) are allowed for this job.***  ");
                            }
                            upWOIncident.hil_warrantystatus = new OptionSetValue(2);//out
                        }
                    }
                    else
                    {
                        if (Found.Entities.Count > 0 && Found.Entities[0].Contains("hil_warrantystatus")
                                              && Found.Entities[0].GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value != 2)
                        {
                            throw new InvalidPluginExecutionException("  ***Only Out Warranty incident(s) are allowed for this job.***  ");
                        }
                        upWOIncident.hil_warrantystatus = new OptionSetValue(2);//Warranty Void
                    }

                    upWOIncident.msdyn_workorderincidentId = woIncident.msdyn_workorderincidentId.Value;
                    service.Update(upWOIncident);
                    #region WorkOrderWarrantyStatusChange
                    if (result == hil_inventorytype.InWarranty)
                    {
                        if (woIncident.msdyn_WorkOrder != null)
                        {
                            upWo.hil_WarrantyStatus = new OptionSetValue(1); //In warranty
                        }
                    }
                    #endregion
                }
                return upWo;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.HelperWOProduct.SetWarrantyStatus" + ex.Message);
            }
        }
        #endregion
    }
    #region UN-USED
    //tracingService.Trace("WorkOrderIncident.PostCreate 1");
    //tracingService.Trace("WorkOrderIncident.PostCreate 2");
    //Set warranty Status
    //tracingService.Trace("WorkOrderIncident.PostCreate 3");
    //HelperCallAllocationRouting.CallAllocationRoute(_entWrkInc, service);
    //Job class half day-full day
    //Common.getandSetJobClass(entity, context, service);
    //tracingService.Trace("WorkOrderIncident.PostCreate 4");
    //tracingService.Trace("WorkOrderIncident.PostCreate 6");
    //tracingService.Trace("WorkOrderIncident.PostCreate 5" + Model);

    //tracingService.Trace("WorkOrderIncident.PostCreate 7");
    #endregion
}