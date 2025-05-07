using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.WorkOrder
{

    public class PreUpdate_SubStatus : IPlugin
    {
        #region MAIN REGION
        private static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == msdyn_workorder.EntityLogicalName && context.MessageName.ToUpper() == "UPDATE" && context.Depth <= 2)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    msdyn_workorder enWorkorder = entity.ToEntity<msdyn_workorder>();

                    //string CurrentStatusName = GetSubStatucName(service, enWorkorder.msdyn_SubStatus.Id).ToUpper();
                    //bool jobCancel = entity.GetAttributeValue<bool>("msdyn_substatus");

                    Entity enWorkorderTemp = service.Retrieve(msdyn_workorder.EntityLogicalName, entity.Id, new ColumnSet("hil_isocr", "hil_kkgcode_sms", "hil_kkgcode", "hil_modeofpayment", "hil_callsubtype"));
                    bool? isOCR = null;
                    tracingService.Trace("1.0 " + enWorkorder.Attributes.Contains("hil_isocr").ToString());
                    tracingService.Trace("1.01 " + enWorkorder.Contains("hil_isocr").ToString());
                    if (enWorkorder.Attributes.Contains("hil_isocr") || enWorkorder.Contains("hil_isocr"))
                    {
                        tracingService.Trace("1.1 " + enWorkorder.GetAttributeValue<bool>("hil_isocr").ToString());
                        isOCR = Convert.ToBoolean(entity["hil_isocr"]);
                    }
                    else if (enWorkorderTemp.Attributes.Contains("hil_isocr") || enWorkorderTemp.Contains("hil_isocr"))
                    {
                        isOCR = Convert.ToBoolean(enWorkorderTemp["hil_isocr"]);
                    }
                    if (isOCR == false) { isOCR = null; }

                    int? kkgDisposition = null;
                    if (enWorkorder.Attributes.Contains("hil_kkgcode_sms") || enWorkorder.Contains("hil_kkgcode_sms"))
                    {
                        tracingService.Trace("1.3 " + enWorkorder.Attributes.Contains("hil_kkgcode_sms").ToString());
                        kkgDisposition = enWorkorder.GetAttributeValue<OptionSetValue>("hil_kkgcode_sms").Value;
                        tracingService.Trace("1.4 " + enWorkorder.Attributes.Contains("hil_kkgcode_sms").ToString());
                    }
                    else if (enWorkorderTemp.Attributes.Contains("hil_kkgcode_sms") || enWorkorderTemp.Contains("hil_kkgcode_sms"))
                    {
                        kkgDisposition = enWorkorderTemp.GetAttributeValue<OptionSetValue>("hil_kkgcode_sms").Value;
                    }

                    string kkgCode = null;
                    if (enWorkorder.Attributes.Contains("hil_kkgcode") || enWorkorder.Contains("hil_kkgcode"))
                    {
                        kkgCode = enWorkorder.GetAttributeValue<string>("hil_kkgcode");
                    }
                    else if (enWorkorderTemp.Attributes.Contains("hil_kkgcode") || enWorkorderTemp.Contains("hil_kkgcode"))
                    {
                        kkgCode = enWorkorderTemp.GetAttributeValue<string>("hil_kkgcode");
                    }

                    Entity enPreEntity = (Entity)context.PreEntityImages["WorkOrderStatusChanePreEntity"];
                    msdyn_workorder enPreWorkOrder = enPreEntity.ToEntity<msdyn_workorder>();
                    tracingService.Trace("1 " + enPreWorkOrder.msdyn_SubStatus.Name);

                    string CurrentStatusName = GetSubStatucName(service, enWorkorder.msdyn_SubStatus.Id).ToUpper();
                    if (enWorkorder.msdyn_SubStatus != null && enPreWorkOrder.msdyn_SubStatus != null)
                    {
                        tracingService.Trace("3.1 " + enWorkorder.msdyn_SubStatus.Id.ToString());
                        tracingService.Trace("3.2 " + enPreWorkOrder.msdyn_SubStatus.Id.ToString());
                        tracingService.Trace("3.3 " + enWorkorder.hil_CloseTicket.ToString());
                        string OldStatusName = enPreWorkOrder.msdyn_SubStatus.Name.ToUpper();
                        tracingService.Trace("3.4 " + CurrentStatusName);
                        tracingService.Trace("3.5 " + OldStatusName);
                        tracingService.Trace("3.6 " + isOCR.ToString());
                        tracingService.Trace("3.7 " + kkgDisposition.ToString());
                        tracingService.Trace("3.8 " + kkgCode);
                        if (OldStatusName == "WORK DONE" && CurrentStatusName == "CLOSED" && isOCR == null && kkgDisposition == null && kkgCode == null)
                        {
                            throw new InvalidPluginExecutionException("STATUS CAN'T BE CHANGED");
                        }
                        else
                        {
                            if (isOCR == true && (OldStatusName == "PARTS FULFILLED" || OldStatusName == "WORK INITIATED" || OldStatusName == "PENDING FOR ALLOCATION"))
                            {
                                //do nothing
                            }
                            else
                            {
                                StatusTransitionValidation(service, enWorkorder.msdyn_SubStatus, enPreWorkOrder.msdyn_SubStatus);
                            }
                        }
                    }
                    else if (enWorkorder.msdyn_SubStatus != null && enPreWorkOrder.msdyn_SubStatus == null)
                    {
                        tracingService.Trace("3");
                        StatusTransitionValidation(service, enWorkorder.msdyn_SubStatus, new EntityReference());
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("SUB-STATUS CAN'T BE CHANGED");
                    }
                    if (CurrentStatusName == "CLOSED")
                    {
                        if (enWorkorderTemp.Attributes.Contains("hil_modeofpayment"))
                        {
                            if (enWorkorderTemp.GetAttributeValue<OptionSetValue>("hil_modeofpayment").Value == 6) // Online PayU
                            {
                                if (enWorkorderTemp.GetAttributeValue<EntityReference>("hil_callsubtype").Id == new Guid("55a71a52-3c0b-e911-a94e-000d3af06cd4")) // AMC Jobs
                                {
                                    if (!Common.ValidateOnlinePayment(service, enWorkorder.Id))
                                    {
                                        throw new InvalidPluginExecutionException("Payment for this AMC is pending, please cross check with customer & try again.");
                                    }
                                }
                            }
                        }
                    }
                    if (CurrentStatusName == "CANCELED") {
                        string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='incident'>
                            <attribute name='ticketnumber' />
                            <attribute name='incidentid' />
                            <filter type='and'>
                                <condition attribute='hil_job' operator='eq' value='{enWorkorder.Id}' />
                                <condition attribute='statecode' operator='ne' value='1' />
                            </filter>
                            </entity>
                            </fetch>";

                        EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entCol.Entities.Count > 0)
                        {
                            throw new InvalidPluginExecutionException($"Open Grievance {entCol.Entities[0].GetAttributeValue<string>("ticketnumber")} is found against the Job. Please first Resolve the Grievance.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message.ToUpper());
            }
        }
        #endregion
        #region ON JOB STATUS PRE UPDATE
        public static string GetSubStatucName(IOrganizationService service, Guid statusGuid)
        {
            string _subStatusName = string.Empty;
            QueryExpression qeObj = new QueryExpression("msdyn_workordersubstatus");
            qeObj.ColumnSet = new ColumnSet("msdyn_name");
            qeObj.Criteria = new FilterExpression(LogicalOperator.And);
            qeObj.Criteria.AddCondition(new ConditionExpression("msdyn_workordersubstatusid", ConditionOperator.Equal, statusGuid));
            EntityCollection ecObj = service.RetrieveMultiple(qeObj);
            if (ecObj.Entities.Count == 1)
            {
                _subStatusName = ecObj.Entities[0].GetAttributeValue<string>("msdyn_name");
            }
            return _subStatusName;
        }
        public static void StatusTransitionValidation(IOrganizationService service, EntityReference erNewStatus, EntityReference erBaseStatus)
        {
            try
            {
                QueryExpression iQuery = new QueryExpression("hil_statustransitionmatrix");
                iQuery.ColumnSet = new ColumnSet(false);
                //iQuery.Criteria = new FilterExpression(LogicalOperator.And);
                //iQuery.Criteria.AddCondition("hil_basestatus", ConditionOperator.Equal, erBaseStatus.Id);
                FilterExpression filter = new FilterExpression(LogicalOperator.And);
                FilterExpression filter1 = new FilterExpression(LogicalOperator.And);
                filter1.Conditions.Add(new ConditionExpression("hil_basestatus", ConditionOperator.Equal, erBaseStatus.Id));
                FilterExpression filter2 = new FilterExpression(LogicalOperator.Or);
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition1", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition2", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition3", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition4", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition5", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition6", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition7", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition8", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition9", ConditionOperator.Equal, erNewStatus.Id));
                filter2.Conditions.Add(new ConditionExpression("hil_statustransition10", ConditionOperator.Equal, erNewStatus.Id));
                filter.AddFilter(filter1);
                filter.AddFilter(filter2);
                iQuery.Criteria = filter;
                EntityCollection ecCollection = service.RetrieveMultiple(iQuery);
                if (ecCollection.Entities.Count == 0)
                {
                    tracingService.Trace("4");
                    throw new InvalidPluginExecutionException("STATUS CAN'T BE CHANGED");
                    #region UN-USED CONDITIONS
                    //Entity enTransitionMatrix = (Entity)ecCollection.Entities[0];
                    //if (enTransitionMatrix.Contains("hil_statustransition1") && enTransitionMatrix.Attributes.Contains("hil_statustransition1"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition1");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition2") && enTransitionMatrix.Attributes.Contains("hil_statustransition2"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition2");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition3") && enTransitionMatrix.Attributes.Contains("hil_statustransition3"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition3");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition4") && enTransitionMatrix.Attributes.Contains("hil_statustransition4"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition4");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition5") && enTransitionMatrix.Attributes.Contains("hil_statustransition5"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition5");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition6") && enTransitionMatrix.Attributes.Contains("hil_statustransition6"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition6");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition7") && enTransitionMatrix.Attributes.Contains("hil_statustransition7"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition7");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition8") && enTransitionMatrix.Attributes.Contains("hil_statustransition8"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition8");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition9") && enTransitionMatrix.Attributes.Contains("hil_statustransition9"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition9");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //else if (enTransitionMatrix.Contains("hil_statustransition10") && enTransitionMatrix.Attributes.Contains("hil_statustransition10"))
                    //{
                    //    erToBe = enTransitionMatrix.GetAttributeValue<EntityReference>("hil_statustransition10");
                    //    if (erToBe.Id == erNewStatus.Id)
                    //        IfMatch = true;
                    //}
                    //if (!IfMatch)
                    //{
                    //    throw new InvalidPluginExecutionException("STATUS CAN'T BE CHANGED");
                    //}
                    #endregion
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        #endregion
    }
}