using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;

namespace Havells_Plugin.WorkOrder
{
    public class PostUpdate_Asynch : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #region MainRegion
            try
            {
                tracingService.Trace("Depth: " + context.Depth.ToString());
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "msdyn_workorder" && context.MessageName.ToUpper() == "CREATE"
                    && context.Depth < 2)
                {
                    tracingService.Trace($@"Step 1: Work Order Assignment Engine Starts . {DateTime.Now.ToString()}");
                    Entity entity = (Entity)context.InputParameters["Target"];
                    //ExecuteCallAllocation(service, entity);
                    OptionSetValue Source = new OptionSetValue();

                    if (entity.Contains("hil_sourceofjob") && entity.Attributes.Contains("hil_sourceofjob"))
                    {
                        Source = entity.GetAttributeValue<OptionSetValue>("hil_sourceofjob");
                    }
                    if(Source != null)
                    {
                        if (Source.Value != 5)
                        {
                            CallAllocation(service, entity, tracingService, entity);
                        }
                    }
                    else
                    {
                        CallAllocation(service, entity, tracingService, entity);
                    }

                    tracingService.Trace($@"Step 2: Work Order Assignment Engine Ends . {DateTime.Now.ToString()}");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrder.PostUpdate_Asynch.Execute: " + ex.Message.ToUpper());
            }
            #endregion
        }
        #region Call Allocation
        public static void CallAllocation(IOrganizationService service, Entity entity, ITracingService tracingService, Entity iJob)
        {
            //1 - Both
            //2 - Direct Engineer
            //3 - Franchisee
            if (!(iJob.Attributes.Contains("hil_pincode") && iJob.Attributes["hil_pincode"] != null
                && iJob.Attributes.Contains("hil_productcategory") && iJob.Attributes["hil_productcategory"] != null
                && iJob.Attributes.Contains("hil_callsubtype") && iJob.Attributes["hil_callsubtype"] != null
                && iJob.Attributes.Contains("hil_salesoffice") && iJob.Attributes["hil_salesoffice"] != null))
            {
                throw new InvalidPluginExecutionException("Mandatory field(s) is/are missing." + System.Environment.NewLine
                    + "Please make sure Pincode, Product Category Call Sub Type and Sales Office should have the value.");
            }
            //tracingService.Trace("1 - " + iJob.GetAttributeValue<string>("msdyn_name"));

            if (true)
            {
                //tracingService.Trace("6");

                Guid iSalesOffice = ((EntityReference)iJob["hil_salesoffice"]).Id;
                Guid iDivision = iJob.GetAttributeValue<EntityReference>("hil_productcategory").Id;
                Guid iPin = iJob.GetAttributeValue<EntityReference>("hil_pincode").Id;
                Guid iCallStype = iJob.GetAttributeValue<EntityReference>("hil_callsubtype").Id;
                //tracingService.Trace("7");
                Product pdt = (Product)service.Retrieve(Product.EntityLogicalName, iDivision, new ColumnSet("hil_repairresource"));
                if (pdt.hil_repairresource != null)
                {
                    //tracingService.Trace("8");
                    EntityReference iAssign = new EntityReference(SystemUser.EntityLogicalName);
                    switch (pdt.hil_repairresource.Value)
                    {
                        case 1: //Both
                            iAssign = GetAssigneeForBoth(service, iPin, iDivision, iCallStype, entity.Id, iJob, iJob.GetAttributeValue<EntityReference>("ownerid"), iSalesOffice, tracingService);
                            if (iAssign.Id != Guid.Empty)
                            {
                                Helper.Assign(iAssign.LogicalName, entity.LogicalName, iAssign.Id, entity.Id, service);
                            }
                            break;
                        case 2: //Direct Engineer
                            iAssign = GetAssignee(service, iPin, iDivision, iCallStype, pdt.hil_repairresource.Value, entity.Id, iJob, iJob.GetAttributeValue<EntityReference>("ownerid"), iSalesOffice, tracingService);
                            if (iAssign.Id != Guid.Empty)
                            {
                                Helper.Assign(iAssign.LogicalName, entity.LogicalName, iAssign.Id, entity.Id, service);
                            }
                            break;
                        case 3: //Franchisee
                            iAssign = GetAssignee(service, iPin, iDivision, iCallStype, pdt.hil_repairresource.Value, entity.Id, iJob, iJob.GetAttributeValue<EntityReference>("ownerid"), iSalesOffice, tracingService);
                            if (iAssign.Id != Guid.Empty)
                            {
                                Helper.Assign(iAssign.LogicalName, entity.LogicalName, iAssign.Id, entity.Id, service);
                            }
                            break;
                        default:
                            GetFallBackRecord(service, iDivision, iSalesOffice, iJob.GetAttributeValue<EntityReference>("ownerid").Id, iJob, 2);
                            break;
                    }
                }
                else
                {
                    GetFallBackRecord(service, iDivision, iSalesOffice, iJob.GetAttributeValue<EntityReference>("ownerid").Id, iJob, 2);
                }
            }
        }
        public static EntityReference GetAssigneeForBoth(IOrganizationService service, Guid iPin, Guid iDivision, Guid iCallStype, Guid JobId, Entity iJob, EntityReference Default, Guid SalesOfficeId, ITracingService tracingService)
        {
            //tracingService.Trace("GetAssigneeForBoth 1");
            EntityReference Assignee = new EntityReference();
            msdyn_workorder eJob = new msdyn_workorder();
            eJob.Id = JobId;
            hil_assignmentmatrix iMatx = new hil_assignmentmatrix();
            Guid iFallBackId = new Guid();
            QueryExpression Query = new QueryExpression(hil_assignmentmatrix.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("hil_franchiseedirectengineer", "ownerid");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, iPin));
            Query.Criteria.AddCondition(new ConditionExpression("hil_division", ConditionOperator.Equal, iDivision));
            Query.Criteria.AddCondition(new ConditionExpression("hil_callsubtype", ConditionOperator.Equal, iCallStype));
            Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count == 1)
            {
                //tracingService.Trace("GetAssigneeForBoth 2");
                iMatx = Found.Entities[0].ToEntity<hil_assignmentmatrix>();
                if (iMatx.Attributes.Contains("hil_franchiseedirectengineer"))
                {
                    //tracingService.Trace("GetAssigneeForBoth 3");
                    EntityReference iRef = (EntityReference)iMatx["hil_franchiseedirectengineer"];
                    Account iFown = (Account)service.Retrieve(Account.EntityLogicalName, iRef.Id, new ColumnSet("ownerid", "customertypecode"));
                    if (iFown.CustomerTypeCode.Value == 9) // DSE
                    {
                        //tracingService.Trace("GetAssigneeForBoth 4");
                        bool IfDEPresent = CheckIfDirectEngineerPresent(service, iFown.OwnerId.Id);
                        if (IfDEPresent)
                        {
                            //tracingService.Trace("GetAssigneeForBoth 5");
                            if (iFown.OwnerId.Id != new Guid("08074320-FCEE-E811-A949-000D3AF03089"))
                            {
                                Assignee = iFown.OwnerId;
                                eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
                                iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 1);
                                if (iFallBackId != Guid.Empty)
                                    eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                                //eJob.hil_AutomaticAssign = new OptionSetValue(1);
                                service.Update(eJob);
                                Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Work Allocated", eJob.Id, service);
                            }
                            else
                            {
                                iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 2);
                                if (iFallBackId != Guid.Empty)
                                    eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                                //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                                service.Update(eJob);
                                Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                            }
                        }
                        else
                        {
                            //tracingService.Trace("GetAssigneeForBoth 6");
                            Assignee = iMatx.OwnerId;
                            //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                            eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
                            iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 1);
                            if (iFallBackId != Guid.Empty)
                                eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                            service.Update(eJob);
                            Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                            //Changed on 21st Jan
                        }
                    }
                    else
                    {
                        //tracingService.Trace("GetAssigneeForBoth 7");
                        if (iFown.OwnerId.Id != new Guid("08074320-FCEE-E811-A949-000D3AF03089"))
                        {
                            Assignee = iFown.OwnerId;
                            eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
                            iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob,1);
                            if (iFallBackId != Guid.Empty)
                                eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                            //eJob.hil_AutomaticAssign = new OptionSetValue(1);
                            service.Update(eJob);
                            Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Work Allocated", eJob.Id, service);
                        }
                        else
                        {
                            iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob ,2);
                            if (iFallBackId != Guid.Empty)
                                eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                            //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                            service.Update(eJob);
                            Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                        }
                    }
                }
                else
                {
                    //tracingService.Trace("GetAssigneeForBoth 8 - Not Assigned");
                    iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 2);
                    if (iFallBackId != Guid.Empty)
                        eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                    //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                    service.Update(eJob);
                    Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                    //Assignee = iMatx.OwnerId;
                }
            }
            else if (Found.Entities.Count > 1)
            {
                //tracingService.Trace("GetAssigneeForBoth 9");
                iMatx = Found.Entities[0].ToEntity<hil_assignmentmatrix>();
                Assignee = iMatx.OwnerId;
                eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
                iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 1);
                //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                if (iFallBackId != Guid.Empty)
                    eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                service.Update(eJob);
                Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                //Changed on 21st Jan
            }
            else if (Found.Entities.Count == 0)
            {
                //tracingService.Trace("GetAssigneeForBoth 10");

                //XXXXXXXXXXXXXXXXXXXXXXXXXXXX---FALL BACK---XXXXXXXXXXXXXXXXXXXXXXXXXXXXX//
                QueryExpression Query1 = new QueryExpression("hil_sbubranchmapping");
                Query1.ColumnSet = new ColumnSet("hil_branchheaduser");
                Query1.Criteria = new FilterExpression(LogicalOperator.And);
                Query1.Criteria.AddCondition(new ConditionExpression("hil_salesoffice", ConditionOperator.Equal, SalesOfficeId));
                Query1.Criteria.AddCondition(new ConditionExpression("hil_productdivision", ConditionOperator.Equal, iDivision));
                Query1.Criteria.AddCondition(new ConditionExpression("hil_branchheaduser", ConditionOperator.NotNull));
                EntityCollection Found1 = service.RetrieveMultiple(Query1);
                if (Found1.Entities.Count > 0)
                {
                    //tracingService.Trace("GetAssigneeForBoth 11");
                    Entity iSbuBranch = Found1.Entities[0];
                    //iMatx = Found1.Entities[0].ToEntity<hil_assignmentmatrix>();
                    Assignee = (EntityReference)iSbuBranch["hil_branchheaduser"];
                    eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iSbuBranch.Id);
                    //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                    service.Update(eJob);
                    Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                    //Changed on 21st Jan
                }
                else
                {
                    //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                    //tracingService.Trace("GetAssigneeForBoth 12 - No Assignment Identified");
                    msdyn_workorder iWo = iJob.ToEntity<msdyn_workorder>();
                    if (iWo.hil_Brand != null)
                    {
                        QueryExpression Query2 = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
                        Query2.ColumnSet = new ColumnSet("hil_brand", "hil_approvername");
                        Query2.Criteria = new FilterExpression(LogicalOperator.And);
                        Query2.Criteria.AddCondition(new ConditionExpression("hil_brand", ConditionOperator.Equal, iWo.hil_Brand.Value));
                        Query2.Criteria.AddCondition(new ConditionExpression("hil_approvername", ConditionOperator.NotNull));
                        EntityCollection iColl = service.RetrieveMultiple(Query2);
                        if (iColl.Entities.Count > 0)
                        {
                            hil_integrationconfiguration iConf = iColl.Entities[0].ToEntity<hil_integrationconfiguration>();
                            Helper.Assign(SystemUser.EntityLogicalName, msdyn_workorder.EntityLogicalName, iConf.hil_approvername.Id, iWo.Id, service);
                            Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", iWo.Id, service);
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException("SBU Branch is not defined for respective Sales Office. Please contact to Administrator");
                        }
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("SBU Branch is not defined for respective Sales Office. Please contact to Administrator");
                    }
                    //Changed on 21st Jan
                    //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                    service.Update(eJob);
                }
            }
            if(Assignee.Id != Guid.Empty)
            {
                bool IfActiveUser = CheckIfUserActive(service, Assignee);
                if (IfActiveUser == true)
                {
                    return Assignee;
                }
                else
                {
                    GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 2);
                    EntityReference AssignNull = new EntityReference();
                    AssignNull.Id = Guid.Empty;
                    return AssignNull;
                }
            }
            return Assignee;
        }
        public static EntityReference GetAssignee(IOrganizationService service, Guid iPin, Guid iDivision, Guid iCallStype, int iResource, Guid JobId, Entity iJob, EntityReference Default, Guid SalesOfficeId, ITracingService tracingService)
        {
            //tracingService.Trace("GetAssignee 1");
            EntityReference Assignee = new EntityReference();
            msdyn_workorder eJob = new msdyn_workorder();
            eJob.Id = JobId;
            //string fetchxml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //                  <entity name='hil_assignmentmatrix'>
            //                  <attribute name='hil_assignmentmatrixid' />
            //                  <attribute name='hil_name' />
            //                  <attribute name='createdon' />
            //                  <attribute name='hil_franchiseedirectengineer' />
            //                  <order attribute='hil_name' descending='false' />
            //                  <filter type='and'>
            //                    <condition attribute='statecode' operator='ne' value='1' />
            //                  </filter>
            //                  <link-entity name='account' from='accountid' to='hil_franchiseedirectengineer' link-type='inner' alias='ae'>
            //                    <link-entity name='systemuser' from='systemuserid' to='owninguser' link-type='inner' alias='af'>
            //                      <filter type='and'>
            //                        <condition attribute='isdisabled' operator='ne' value='1' />
            //                      </filter>
            //                    </link-entity>
            //                  </link-entity>
            //                </entity>
            //              </fetch>";
            //RetrieveMultipleRequest fetchRequest1 = new RetrieveMultipleRequest
            //{
            //    Query = new FetchExpression(fetchxml)
            //};
            //EntityCollection Found = ((RetrieveMultipleResponse)
            //service.Execute(fetchRequest1)).EntityCollection;
            //< condition attribute = 'systemuserid' operator= 'ne-userid' />
            hil_assignmentmatrix iMatx = new hil_assignmentmatrix();
            Guid iFallBackId = new Guid();
            QueryExpression Query = new QueryExpression(hil_assignmentmatrix.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("hil_franchiseedirectengineer", "ownerid");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, iPin));
            Query.Criteria.AddCondition(new ConditionExpression("hil_division", ConditionOperator.Equal, iDivision));
            Query.Criteria.AddCondition(new ConditionExpression("hil_callsubtype", ConditionOperator.Equal, iCallStype));
            Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 1)
            {
                //tracingService.Trace("GetAssignee 2");
                foreach (hil_assignmentmatrix iMatx1 in Found.Entities)
                {
                    if (iMatx1.Attributes.Contains("hil_franchiseedirectengineer"))
                    {
                        //tracingService.Trace("GetAssignee 3");
                        EntityReference iRef = (EntityReference)iMatx1["hil_franchiseedirectengineer"];
                        Account iFown = (Account)service.Retrieve(Account.EntityLogicalName, iRef.Id, new ColumnSet("ownerid"));
                        SystemUser iSys = (SystemUser)service.Retrieve(SystemUser.EntityLogicalName, iFown.OwnerId.Id, new ColumnSet("positionid"));
                        if (iSys.PositionId != null)
                        {
                            //tracingService.Trace("GetAssignee 4");
                            if ((iResource == 2) && (iSys.PositionId.Name == "DSE"))
                            {
                                //tracingService.Trace("GetAssignee 5");
                                bool IfDEPresent = CheckIfDirectEngineerPresent(service, iFown.OwnerId.Id);
                                if (IfDEPresent)
                                {
                                    //Changed on 21st Jan
                                    //tracingService.Trace("GetAssignee 6");
                                    if (iFown.OwnerId.Id != new Guid("08074320-FCEE-E811-A949-000D3AF03089"))
                                    {
                                        Assignee = iFown.OwnerId;
                                        eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx1.Id);
                                        iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 1);
                                        if (iFallBackId != Guid.Empty)
                                            eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                                        //eJob.hil_AutomaticAssign = new OptionSetValue(1);
                                        service.Update(eJob);
                                        Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Work Allocated", eJob.Id, service);
                                    }
                                    else
                                    {
                                        iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 1);
                                        if (iFallBackId != Guid.Empty)
                                            eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                                        //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                                        service.Update(eJob);
                                        Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                                    }
                                    break;
                                }
                                else
                                {
                                    //Changed on 21st Jan
                                    //tracingService.Trace("GetAssignee 7");
                                    Assignee = iMatx1.OwnerId;
                                    eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx1.Id);
                                    iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 2);
                                    //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                                    if (iFallBackId != Guid.Empty)
                                        eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                                    service.Update(eJob);
                                    Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                                    break;
                                }
                            }
                            else if ((iResource == 3) && (iSys.PositionId.Name == "Franchise"))
                            {
                                //Changed on 21st Jan
                                //tracingService.Trace("GetAssignee 8");
                                if (iFown.OwnerId.Id != new Guid("08074320-FCEE-E811-A949-000D3AF03089"))
                                {
                                    Assignee = iFown.OwnerId;
                                    eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx1.Id);
                                    iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 1);
                                    if (iFallBackId != Guid.Empty)
                                        eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                                    //eJob.hil_AutomaticAssign = new OptionSetValue(1);
                                    service.Update(eJob);
                                    Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Work Allocated", eJob.Id, service);
                                }
                                else
                                {
                                    iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 2);
                                    if (iFallBackId != Guid.Empty)
                                        eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                                    //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                                    service.Update(eJob);
                                    Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                                }
                                break;
                            }
                            else
                            {
                                //Changed on 21st Jan
                                //tracingService.Trace("GetAssignee 9");
                                Assignee = iMatx1.OwnerId;  //Branch Head
                                eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx1.Id);
                                iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 1);
                                if (iFallBackId != Guid.Empty)
                                    eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                                //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                                service.Update(eJob);
                                Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                            }
                        }
                        else
                        {
                            //Changed on 21st Jan
                            //tracingService.Trace("GetAssignee 10");
                            Assignee = iMatx1.OwnerId;  //Branch Head
                            eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx1.Id);
                            iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 1);
                            if (iFallBackId != Guid.Empty)
                                eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                            //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                            service.Update(eJob);
                            Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                        }
                    }
                    else
                    {
                        //Changed on 21st Jan
                        //tracingService.Trace("GetAssignee 11");
                        Assignee = iMatx1.OwnerId;  //Branch Head
                        eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx1.Id);
                        iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 1);
                        if (iFallBackId != Guid.Empty)
                            eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                        //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                        service.Update(eJob);
                        Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                    }
                }
            }
            else if (Found.Entities.Count == 1)
            {
                //Changed on 21st Jan
                //tracingService.Trace("GetAssignee 12");
                iMatx = Found.Entities[0].ToEntity<hil_assignmentmatrix>();
                if (iMatx.Attributes.Contains("hil_franchiseedirectengineer"))
                {
                    //tracingService.Trace("GetAssignee 13");
                    EntityReference iRef = (EntityReference)iMatx["hil_franchiseedirectengineer"];
                    if (iRef != null)
                    {
                        //tracingService.Trace("GetAssignee 14");
                        Account iFown = (Account)service.Retrieve(Account.EntityLogicalName, iRef.Id, new ColumnSet("ownerid"));
                        SystemUser iSys = (SystemUser)service.Retrieve(SystemUser.EntityLogicalName, iFown.OwnerId.Id, new ColumnSet("positionid"));
                        if (iSys.PositionId != null)
                        {
                            //tracingService.Trace("GetAssignee 14");
                            if ((iResource == 2) && (iSys.PositionId.Name == "DSE"))
                            {
                                bool IfDEPresent = CheckIfDirectEngineerPresent(service, iFown.OwnerId.Id);
                                if (IfDEPresent)
                                {
                                    //Changed on 21st Jan
                                    //tracingService.Trace("GetAssignee 15");
                                    if (iFown.OwnerId.Id != new Guid("08074320-FCEE-E811-A949-000D3AF03089"))
                                    {
                                        Assignee = iFown.OwnerId;
                                        eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
                                        iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 1);
                                        if (iFallBackId != Guid.Empty)
                                            eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                                        //eJob.hil_AutomaticAssign = new OptionSetValue(1);
                                        service.Update(eJob);
                                        Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Work Allocated", eJob.Id, service);
                                    }
                                    else
                                    {
                                        iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 1);
                                        if (iFallBackId != Guid.Empty)
                                            eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                                        //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                                        service.Update(eJob);
                                        Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                                    }
                                }
                                else
                                {
                                    //Changed on 21st Jan
                                    //tracingService.Trace("GetAssignee 7");
                                    Assignee = iMatx.OwnerId;
                                    eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
                                    iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 2);
                                    //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                                    if (iFallBackId != Guid.Empty)
                                        eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                                    service.Update(eJob);
                                    Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                                }
                            }
                            else if ((iResource == 3) && (iSys.PositionId.Name == "Franchisee"))
                            {
                                //Changed on 21st Jan
                                //tracingService.Trace("GetAssignee 16");
                                if(iFown.OwnerId.Id != new Guid("08074320-FCEE-E811-A949-000D3AF03089"))
                                {
                                    Assignee = iFown.OwnerId;
                                    eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
                                    iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 1);
                                    if (iFallBackId != Guid.Empty)
                                        eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                                    //eJob.hil_AutomaticAssign = new OptionSetValue(1);
                                    service.Update(eJob);
                                    Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Work Allocated", eJob.Id, service);
                                }
                                else
                                {
                                    iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 2);
                                    if (iFallBackId != Guid.Empty)
                                        eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                                    //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                                    service.Update(eJob);
                                    Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                                }
                            }
                            else
                            {
                                //Changed on 21st Jan
                                //tracingService.Trace("GetAssignee 17");
                                Assignee = iMatx.OwnerId; //Branch Head
                                eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
                                iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 1);
                                if (iFallBackId != Guid.Empty)
                                    eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                                //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                                service.Update(eJob);
                                Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                            }
                        }
                        else
                        {
                            //Changed on 21st Jan
                            //tracingService.Trace("GetAssignee 18");
                            Assignee = iMatx.OwnerId; //Branch Head
                            eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
                            iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 1);
                            if (iFallBackId != Guid.Empty)
                                eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                            //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                            service.Update(eJob);
                            Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                        }
                    }
                    else
                    {
                        //Changed on 21st Jan
                        //tracingService.Trace("GetAssignee 19");
                        Assignee = iMatx.OwnerId; //Branch Head
                        eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
                        iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 1);
                        if (iFallBackId != Guid.Empty)
                            eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                        //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                        service.Update(eJob);
                        Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                    }
                }
                else
                {
                    //Changed on 21st Jan
                    //tracingService.Trace("GetAssignee 20");
                    Assignee = iMatx.OwnerId; //Branch Head
                    eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
                    iFallBackId = GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 1);
                    if (iFallBackId != Guid.Empty)
                        eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iFallBackId);
                    //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                    service.Update(eJob);
                    Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                }
            }
            else if (Found.Entities.Count == 0)
            {
                //Changed on 21st Jan
                //tracingService.Trace("GetAssignee 21");
                QueryExpression Query1 = new QueryExpression("hil_sbubranchmapping");
                Query1.ColumnSet = new ColumnSet("hil_branchheaduser");
                Query1.Criteria = new FilterExpression(LogicalOperator.And);
                Query1.Criteria.AddCondition(new ConditionExpression("hil_salesoffice", ConditionOperator.Equal, SalesOfficeId));
                Query1.Criteria.AddCondition(new ConditionExpression("hil_productdivision", ConditionOperator.Equal, iDivision));
                Query1.Criteria.AddCondition(new ConditionExpression("hil_branchheaduser", ConditionOperator.NotNull));
                EntityCollection Found1 = service.RetrieveMultiple(Query1);
                if (Found1.Entities.Count > 0)
                {
                    //Changed on 21st Jan
                    //tracingService.Trace("GetAssignee 22");
                    Entity iSbuBranch = Found1.Entities[0];
                    //iMatx = Found1.Entities[0].ToEntity<hil_assignmentmatrix>();
                    Assignee = (EntityReference)iSbuBranch["hil_branchheaduser"];
                    eJob["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", iSbuBranch.Id);
                    //eJob.hil_AutomaticAssign = new OptionSetValue(2);
                    service.Update(eJob);
                    Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", eJob.Id, service);
                }
                //XXXXXXXXXXXXXXXXXXXXXXXXXXXX---FALL BACK---XXXXXXXXXXXXXXXXXXXXXXXXXXXXX//
                //QueryExpression Query1 = new QueryExpression(hil_assignmentmatrix.EntityLogicalName);
                //Query1.ColumnSet = new ColumnSet("hil_franchiseedirectengineer", "ownerid");
                //Query1.Criteria = new FilterExpression(LogicalOperator.And);
                //Query1.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, iPin));
                //Query1.Criteria.AddCondition(new ConditionExpression("hil_division", ConditionOperator.Equal, iDivision));
                //EntityCollection Found1 = service.RetrieveMultiple(Query1);
                //if (Found1.Entities.Count > 0)
                //{
                //    iMatx = Found1.Entities[0].ToEntity<hil_assignmentmatrix>();
                //    Assignee = iMatx.OwnerId;
                //    eJob.hil_AssignmentMatrix = new EntityReference(hil_assignmentmatrix.EntityLogicalName, iMatx.Id);
                //    service.Update(eJob);
                //}
                else
                {
                    msdyn_workorder iWo = iJob.ToEntity<msdyn_workorder>();
                    if (iWo.hil_Brand != null)
                    {
                        QueryExpression Query2 = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
                        Query2.ColumnSet = new ColumnSet("hil_brand", "hil_approvername");
                        Query2.Criteria = new FilterExpression(LogicalOperator.And);
                        Query2.Criteria.AddCondition(new ConditionExpression("hil_brand", ConditionOperator.Equal, iWo.hil_Brand.Value));
                        Query2.Criteria.AddCondition(new ConditionExpression("hil_approvername", ConditionOperator.NotNull));
                        EntityCollection iColl = service.RetrieveMultiple(Query2);
                        if (iColl.Entities.Count > 0)
                        {
                            hil_integrationconfiguration iConf = iColl.Entities[0].ToEntity<hil_integrationconfiguration>();
                            Helper.Assign(SystemUser.EntityLogicalName, msdyn_workorder.EntityLogicalName, iConf.hil_approvername.Id, iWo.Id, service);
                            Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", iWo.Id, service);
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException("SBU Branch is not defined for respective Sales Office. Please contact to Administrator");
                        }
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("SBU Branch is not defined for respective Sales Office. Please contact to Administrator");
                    }
                }
            }
            if(Assignee.Id != Guid.Empty)
            {
                bool IfActiveUser = CheckIfUserActive(service, Assignee);
                if (IfActiveUser == true)
                {
                    return Assignee;
                }
                else
                {
                    GetFallBackRecord(service, iDivision, SalesOfficeId, Default.Id, iJob, 2);
                    EntityReference AssignNull = new EntityReference();
                    AssignNull.Id = Guid.Empty;
                    return AssignNull;
                }
            }
            return Assignee;
        }
        #endregion
        #region Check If Direct Engineer Present
        public static bool CheckIfDirectEngineerPresent(IOrganizationService service, Guid iEngineer)
        {
            bool IfPresent = false;
            QueryExpression Query = new QueryExpression(msdyn_timeoffrequest.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("ownerid", ConditionOperator.Equal, iEngineer));
            Query.Criteria.AddCondition(new ConditionExpression("msdyn_starttime", ConditionOperator.Today));
            //Query.Criteria.AddCondition(new ConditionExpression("msdyn_endtime", ConditionOperator.OnOrAfter, DateTime.Now.ToUniversalTime()));
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                IfPresent = true;
            }
            return IfPresent;
        }
        #endregion
        #region Get Fall Back Record
        public static Guid GetFallBackRecord(IOrganizationService service, Guid iDivision, Guid iSalesOffice, Guid ownerId, Entity entity, int OperationCode)
        {
            SBUBranchMapping sBUBranchMappingObj = new SBUBranchMapping();
            msdyn_workorder iWo = new msdyn_workorder();
            QueryExpression Query1 = new QueryExpression("hil_sbubranchmapping");
            Query1.ColumnSet = new ColumnSet(new string[] { "hil_branchheaduser" });
            Query1.Criteria = new FilterExpression(LogicalOperator.And);
            Query1.Criteria.AddCondition(new ConditionExpression("hil_salesoffice", ConditionOperator.Equal, iSalesOffice));
            Query1.Criteria.AddCondition(new ConditionExpression("hil_productdivision", ConditionOperator.Equal, iDivision));
            Query1.Criteria.AddCondition(new ConditionExpression("hil_branchheaduser", ConditionOperator.NotNull));
            EntityCollection Found1 = service.RetrieveMultiple(Query1);
            if (Found1.Entities.Count > 0)
            {
                sBUBranchMappingObj.SBUBranchMappingId = Found1.Entities[0].Id;

                if (Found1.Entities[0].Contains("hil_branchheaduser"))
                {
                    sBUBranchMappingObj.BranchHeadUser = Found1.Entities[0].GetAttributeValue<EntityReference>("hil_branchheaduser");
                }
            }
            if(OperationCode == 2)
            {
                if (sBUBranchMappingObj.SBUBranchMappingId != Guid.Empty)
                {
                    msdyn_workorder msdyn_WorkorderUpdateObj = new msdyn_workorder();
                    msdyn_WorkorderUpdateObj.Id = entity.Id;
                    msdyn_WorkorderUpdateObj["hil_regardingfallback"] = new EntityReference("hil_sbubranchmapping", sBUBranchMappingObj.SBUBranchMappingId);
                    //msdyn_WorkorderUpdateObj.hil_AutomaticAssign = new OptionSetValue(2);
                    Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", msdyn_WorkorderUpdateObj.Id, service);
                    service.Update(msdyn_WorkorderUpdateObj);
                    //Changed on 21st Jan
                    if (sBUBranchMappingObj.BranchHeadUser != null)
                    {
                        Helper.Assign(sBUBranchMappingObj.BranchHeadUser.LogicalName, msdyn_workorder.EntityLogicalName, sBUBranchMappingObj.BranchHeadUser.Id, entity.Id, service);
                        Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", entity.Id, service);
                    }
                    else
                    {
                        iWo = entity.ToEntity<msdyn_workorder>();
                        if (iWo.hil_Brand != null)
                        {
                            QueryExpression Query = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
                            Query.ColumnSet = new ColumnSet("hil_brand", "hil_approvername");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition(new ConditionExpression("hil_brand", ConditionOperator.Equal, iWo.hil_Brand.Value));
                            Query.Criteria.AddCondition(new ConditionExpression("hil_approvername", ConditionOperator.NotNull));
                            EntityCollection iColl = service.RetrieveMultiple(Query);
                            if (iColl.Entities.Count > 0)
                            {
                                hil_integrationconfiguration iConf = iColl.Entities[0].ToEntity<hil_integrationconfiguration>();
                                Helper.Assign(sBUBranchMappingObj.BranchHeadUser.LogicalName, msdyn_workorder.EntityLogicalName, iConf.hil_approvername.Id, entity.Id, service);
                                Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", entity.Id, service);
                            }
                            else
                            {
                                throw new InvalidPluginExecutionException("SBU Branch is not defined for respective Sales Office. Please contact to Administrator");
                            }
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException("SBU Branch is not defined for respective Sales Office. Please contact to Administrator");
                        }
                    }
                }
                else
                {
                    iWo = entity.ToEntity<msdyn_workorder>();
                    if (iWo.hil_Brand != null)
                    {
                        QueryExpression Query = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
                        Query.ColumnSet = new ColumnSet("hil_brand", "hil_approvername");
                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                        Query.Criteria.AddCondition(new ConditionExpression("hil_brand", ConditionOperator.Equal, iWo.hil_Brand.Value));
                        Query.Criteria.AddCondition(new ConditionExpression("hil_approvername", ConditionOperator.NotNull));
                        EntityCollection iColl = service.RetrieveMultiple(Query);
                        if (iColl.Entities.Count > 0)
                        {
                            hil_integrationconfiguration iConf = iColl.Entities[0].ToEntity<hil_integrationconfiguration>();
                            Helper.Assign(sBUBranchMappingObj.BranchHeadUser.LogicalName, msdyn_workorder.EntityLogicalName, iConf.hil_approvername.Id, entity.Id, service);
                            Havells_Plugin.JobStatuses.StatusTransition.setWoSubStatusRecord("Pending for Allocation", entity.Id, service);
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException("SBU Branch is not defined for respective Sales Office. Please contact to Administrator");
                        }
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("SBU Branch is not defined for respective Sales Office. Please contact to Administrator");
                    }
                }
            }
            return sBUBranchMappingObj.SBUBranchMappingId;
        }
        #endregion
        #region If Active User
        public static bool CheckIfUserActive(IOrganizationService service, EntityReference Assignee)
        {
            bool IfActive = false;
            SystemUser iUser = (SystemUser)service.Retrieve(SystemUser.EntityLogicalName, Assignee.Id, new ColumnSet("isdisabled"));
            if(iUser.IsDisabled == false)
            {
                IfActive = true;
            }
            return IfActive;
        }
        #endregion
    }
    public class SBUBranchMapping
    {
        public Guid SBUBranchMappingId { get; set; }
        public EntityReference BranchHeadUser { get; set; }
    }
}