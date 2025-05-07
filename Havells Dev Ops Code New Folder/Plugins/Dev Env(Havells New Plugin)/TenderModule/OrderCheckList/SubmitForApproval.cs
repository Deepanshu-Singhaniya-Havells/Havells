using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.OrderCheckList
{
    public class SubmitForApproval : IPlugin
    {
        public static ITracingService tracingService = null;

        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace(entity.LogicalName);
                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));
                    //hil_approval approval = entity.ToEntity<hil_approval>();
                    int status = entity.GetAttributeValue<OptionSetValue>("hil_approvalstatus").Value;
                    string department = entity.GetAttributeValue<EntityReference>("hil_department").Name.ToString();
                    tracingService.Trace("department " + department);

                    if (status == 3)
                    {
                        #region autoCorrect Serial No.
                        QueryExpression qsChecklist = new QueryExpression("hil_orderchecklistproduct");
                        qsChecklist.ColumnSet = new ColumnSet("hil_name", "hil_orderchecklistid");
                        qsChecklist.Criteria = new FilterExpression(LogicalOperator.And);
                        qsChecklist.Criteria.AddCondition("hil_orderchecklistid", ConditionOperator.Equal, entity.Id);
                        qsChecklist.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                        EntityCollection entCol = service.RetrieveMultiple(qsChecklist);
                        int count = entCol.Entities.Count;
                        int i = 1;
                        foreach (Entity prod in entCol.Entities)
                        {
                            string _tend = prod.GetAttributeValue<EntityReference>("hil_orderchecklistid").Name;
                            string name = prod.GetAttributeValue<string>("hil_name");
                            Entity prd = new Entity(prod.LogicalName);
                            prd.Id = prod.Id;
                            prd["hil_name"] = _tend + "_" + i.ToString().PadLeft(3, '0');
                            service.Update(prd);
                            i++;
                        }
                        #endregion
                        qsChecklist = new QueryExpression("hil_orderchecklistproduct");
                        qsChecklist.ColumnSet = new ColumnSet(true);
                        qsChecklist.NoLock = true;
                        qsChecklist.Criteria = new FilterExpression(LogicalOperator.And);
                        qsChecklist.Criteria.AddCondition("hil_orderchecklistid", ConditionOperator.Equal, entity.Id);
                        qsChecklist.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                        entCol = service.RetrieveMultiple(qsChecklist);
                        decimal poquantity = 0;
                        foreach (Entity product in entCol.Entities)
                        {
                            if (product.Contains("hil_name"))
                            {
                                if (product.Contains("hil_tenderproductid"))
                                {
                                    tracingService.Trace("With Tender");
                                    var oclPrdName = Convert.ToString(product.GetAttributeValue<string>("hil_name"));
                                    if (product.Contains("hil_poqty"))
                                    {
                                        poquantity = Convert.ToDecimal(product.GetAttributeValue<decimal>("hil_poqty"));
                                    }
                                    if (!product.Contains("hil_poqty"))
                                    {
                                        throw new InvalidPluginExecutionException("PO Qty is missing in " + oclPrdName);
                                    }
                                    if (!product.Contains("hil_basicpriceinrsmtr"))
                                    {
                                        throw new InvalidPluginExecutionException("HO Price or LP is missing in " + oclPrdName);
                                    }
                                    else if (!product.Contains("hil_porate"))
                                    {
                                        throw new InvalidPluginExecutionException("PO Rate is missing in " + oclPrdName);
                                    }
                                    else if (!product.Contains("hil_poamount"))
                                    {
                                        throw new InvalidPluginExecutionException("PO Amount is missing in " + oclPrdName);
                                    }
                                    else if (!product.Contains("hil_product"))
                                    {
                                        throw new InvalidPluginExecutionException("Product code is missing in " + oclPrdName);
                                    }
                                    else if (!product.Contains("hil_plantcode"))
                                    {
                                        throw new InvalidPluginExecutionException("Product code is missing in " + oclPrdName);
                                    }
                                    else if (!product.Contains("hil_inspectiontype"))
                                    {
                                        throw new InvalidPluginExecutionException("Inspection type is missing of " + oclPrdName);
                                    }
                                }
                                else if (!product.Contains("hil_tenderproductid"))
                                {
                                    tracingService.Trace("With out Tender");
                                    var TenderLine = Convert.ToString(product.GetAttributeValue<string>("hil_name"));
                                    if (product.Contains("hil_quantity"))
                                    {
                                        tracingService.Trace("OrderchecklistProduct-3");
                                        poquantity = Convert.ToDecimal(product.GetAttributeValue<decimal>("hil_quantity"));
                                    }
                                    if (!product.Contains("hil_quantity"))
                                    {
                                        throw new InvalidPluginExecutionException("PO Qty is missing in " + TenderLine);
                                    }
                                    if (!product.Contains("hil_basicpriceinrsmtr"))
                                    {
                                        throw new InvalidPluginExecutionException("HO Price or LP is missing in " + TenderLine);
                                    }
                                    else if (!product.Contains("hil_lprsmtr"))
                                    {
                                        throw new InvalidPluginExecutionException("PO Rate is missing in " + TenderLine);
                                    }
                                    else if (!product.Contains("hil_product"))
                                    {
                                        throw new InvalidPluginExecutionException("Product code is missing in " + TenderLine);
                                    }
                                    else if (!product.Contains("hil_plantcode"))
                                    {
                                        throw new InvalidPluginExecutionException("Product code is missing in " + TenderLine);
                                    }
                                    else if (!product.Contains("hil_inspectiontype"))
                                    {
                                        throw new InvalidPluginExecutionException("Inspection type is missing of " + TenderLine);
                                    }
                                    tracingService.Trace("OrderchecklistProduct-4");
                                }
                            }
                        }

                        if (department == "Cable")
                        {
                            if (entity.Contains("hil_approveddatasheetgtp"))
                            {
                                if (entity.GetAttributeValue<bool>("hil_approveddatasheetgtp"))
                                {
                                    if (!attachmentAttached(service, "Approved GTP", entity.ToEntityReference()))
                                        throw new InvalidPluginExecutionException("Please Attach Approved GTP");
                                }
                            }
                            if (entity.Contains("hil_drumlengthschedule"))
                            {
                                if (entity.GetAttributeValue<bool>("hil_drumlengthschedule"))
                                {
                                    if (!attachmentAttached(service, "Drum Schedule", entity.ToEntityReference()))
                                        throw new InvalidPluginExecutionException("Please Attach Drum Schedule");
                                }
                            }
                            if (entity.Contains("hil_qap"))
                            {
                                if (entity.GetAttributeValue<bool>("hil_qap"))
                                {
                                    if (!attachmentAttached(service, "Approved QAP", entity.ToEntityReference()))
                                        throw new InvalidPluginExecutionException("Please Attach Approved QAP");
                                }
                            }
                            if (entity.Contains("hil_manufacturingclearance"))
                            {
                                if (entity.GetAttributeValue<bool>("hil_manufacturingclearance"))
                                {
                                    if (!attachmentAttached(service, "Manufacturing Clearance Certificate", entity.ToEntityReference()))
                                        throw new InvalidPluginExecutionException("Please Attach Manufacturing Clearance Certificate");
                                }
                            }
                            createApproval(entity, service, "Order Check List Approval");
                        }
                        else if (department == "Motor")
                        {
                            tracingService.Trace("Motor Order Check List Approval");
                            createApproval(entity, service, "Motor Order Check List Approval");
                        }
                        validateLP(service, entity);
                    }

                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("|| SubmitForApproval Error: " + ex.Message);
            }
        }
        public static void validityCheck(IOrganizationService service, Entity entity)
        {
            EntityReference department = entity.GetAttributeValue<EntityReference>("hil_department");
            Entity _DepartmentEnt = service.Retrieve(department.LogicalName, department.Id, new ColumnSet("hil_ocllifecycleindays"));
            DateTime createdOn = entity.GetAttributeValue<DateTime>("createdon");
            TimeSpan difference = DateTime.Now - createdOn; //create TimeSpan object

            if (difference.Days > _DepartmentEnt.GetAttributeValue<int>("hil_ocllifecycleindays") &&
                _DepartmentEnt.GetAttributeValue<int>("hil_ocllifecycleindays") != 0)
            {
                throw new InvalidPluginExecutionException("OCL No. " + entity.GetAttributeValue<string>("hil_name") 
                    + " Validity has been expired. \nYou are requested to kindly recreate the OCL and trigger for approval within 24 hours of creation.");
            }
        }
        public static void validateLP(IOrganizationService service, Entity entity)
        {
            validityCheck(service, entity);
            string _rowIds = string.Empty;
            QueryExpression query = new QueryExpression("hil_orderchecklistproduct");
            query.ColumnSet = new ColumnSet("hil_orderchecklistproductid", "hil_product", "hil_name", "hil_lprsmtr", "hil_unit");
            query.Distinct = true;
            query.Criteria.AddCondition("hil_orderchecklistid", ConditionOperator.Equal, entity.Id);
            LinkEntity EntityP = new LinkEntity("hil_orderchecklistproduct", "productpricelevel", "hil_product", "productid", JoinOperator.LeftOuter);
            EntityP.Columns = new ColumnSet("amount");
            EntityP.EntityAlias = "prd";
            query.LinkEntities.Add(EntityP);
            EntityCollection _entColl = service.RetrieveMultiple(query);
            //tracingService.Trace("_entColl.Entities.Count " + _entColl.Entities.Count);
            if (_entColl.Entities.Count > 0)
            {
                Decimal _tndProdLP;
                Decimal _prodLP;
                decimal newLP;
                bool _flg = false;
                _rowIds = "LP has been changed for Product# ";
                foreach (Entity ent in _entColl.Entities)
                {
                    _tndProdLP = 0;
                    _prodLP = 0;
                    _tndProdLP = ent.GetAttributeValue<Money>("hil_lprsmtr").Value;
                    //tracingService.Trace("_tndProdLP " + _tndProdLP);
                    if (ent.Contains("prd.amount"))
                    {
                        newLP = ((Money)ent.GetAttributeValue<AliasedValue>("prd.amount").Value).Value;
                        //tracingService.Trace("newLP " + newLP);
                        _prodLP = Math.Round(newLP, 2, MidpointRounding.AwayFromZero);
                    }
                    if (_tndProdLP != _prodLP)
                    {
                        //tracingService.Trace("_tndProdLP " + _tndProdLP);
                        //tracingService.Trace("_prodLP " + _prodLP);
                        _rowIds = _rowIds + ent.GetAttributeValue<EntityReference>("hil_product").Name + ",";
                        _flg = true;
                    }
                }
                if (_flg)
                    throw new InvalidPluginExecutionException(_rowIds);
            }
        }
        public bool attachmentAttached(IOrganizationService service, String Documenttype, EntityReference regarding)
        {
            bool isAttached = false;
            string fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='hil_attachment'>
                                <attribute name='activityid' />
                                <attribute name='subject' />
                                <attribute name='createdon' />
                                <order attribute='subject' descending='false' />
                                <filter type='and'>
                                  <condition attribute='regardingobjectid' operator='eq' uiname='" + regarding.Name + @"' uitype='" + regarding.LogicalName + @"' value='" + regarding.Id + @"' />
                                </filter>
                                <link-entity name='hil_attachmentdocumenttype' from='hil_attachmentdocumenttypeid' to='hil_documenttype' link-type='inner' alias='aa'>
                                  <filter type='and'>
                                    <condition attribute='hil_name' operator='eq' value='" + Documenttype + @"' />
                                  </filter>
                                </link-entity>
                              </entity>
                            </fetch>";
            EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(fetch));
            if (entCol.Entities.Count > 0)
                isAttached = true;
            return isAttached;
        }
        public void createApproval(Entity entity, IOrganizationService service, string purpose)
        {
            try
            {
                OrganizationRequest req = new OrganizationRequest("hil_CreateApprovals");
                req["EntityName"] = entity.LogicalName;
                req["EntityID"] = entity.Id.ToString();
                req["Purpose"] = purpose;//"Order Check List Approval";
                req["Target"] = new EntityReference("hil_approval", new Guid("a59b4a18-080b-ec11-b6e6-6045bd72f2f7"));

                OrganizationResponse response = service.Execute(req);
                tracingService.Trace("Create approval request");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("|| SubmitForApproval.createApproval Error " + ex.Message);
            }
        }
        //public void createApproval(Entity entity, IOrganizationService service)
        //{
        //    try
        //    {
        //        OrganizationRequest req = new OrganizationRequest("hil_CreateApprovals");
        //        req["EntityName"] = entity.LogicalName;
        //        req["EntityID"] = entity.Id.ToString();
        //        req["Purpose"] = "Order Check List Approval";
        //        req["Target"] = new EntityReference("hil_approval", new Guid("a59b4a18-080b-ec11-b6e6-6045bd72f2f7"));

        //        OrganizationResponse response = service.Execute(req);
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new InvalidPluginExecutionException("|| SubmitForApproval.createApproval Error " + ex.Message);
        //    }
        //}
    }
}

