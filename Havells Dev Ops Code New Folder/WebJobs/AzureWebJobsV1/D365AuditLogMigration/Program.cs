using Microsoft.Crm.Sdk.Messages;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Auth;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace D365AuditLogMigration
{
    public class Program
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

        #region Global Varialble declaration
        static IOrganizationService _service;
        static IOrganizationService _serviceFSM;
        #endregion
        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            _serviceFSM = ConnectToCRM(string.Format(connStr, "https://havellsfsm.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                //D365AuditLogMigration();
                //D365AuditLogMigration_ActiveSystemUser*/();
                //D365AuditLogMigration_DeactiveSystemUser();
                //D365AuditLogMigration_Customer();
                //D365AuditLogMigration_Customer();
                //D365AuditLogMigrationInventoryRequest();
                //<attribute name='msdyn_name' />
                //DeleteAuditLog();
                //DeleteAsyncOperations();
                //DeletePlugInTraces();
                //DeleteWorkFlowBase();
                //DeleteProcessSessions();
                //D365AuditLogMigrationProductRequest();
                //D365AuditLogMigration_Jobs();
                //GetDistinctEntityInAuditLog();
                //GetEntityMinDateOfAuditLog("10196");
                //DeleteAuditLog_Entitywise();
                //DeleteAuditLog_DeleteEntry();
                //DeleteImportFile();

                //DeleteEmails();
                //D365AuditLogMigration_Delta("hil_claimline", "2020-01-01", "2020-12-31");
                //D365AuditLogMigration_Delta("msdyn_workorder", "2022-10-10", "2022-10-15");
                //D365AuditLogMigration_Delta("msdyn_customerasset", "2021-03-01", "2021-03-31");
                //D365AuditLogMigration_Delta("msdyn_customerasset", "2021-04-01", "2021-04-30");
                //D365AuditLogMigration_Delta("msdyn_customerasset", "2021-05-01", "2021-05-31");
                //D365AuditLogMigration_Delta("msdyn_customerasset", "2021-06-01", "2021-06-30");
                //D365AuditLogMigration_Delta("msdyn_customerasset", "2021-07-01", "2021-07-31");

                //D365AuditLogMigrationDaily("msdyn_workorder");

                DeleteDuplicateClaimLines(_service);

                //FSM App Master Data Migration
                //FSMDataMigration();

                //IotServiceCall _obj = new IotServiceCall();
                //AuthenticateConsumer _retVal;
                //_retVal = _obj.AuthenticateConsumerAMC(new AuthenticateConsumer() {LoginUserId= "8285906486",SourceType="5" }, _service);

                //PrincipalObjectAttributeAccess();

                //Entity iJob = _service.Retrieve(msdyn_workorder.EntityLogicalName, new Guid("0896cd84-2e88-ee11-8179-6045bdac5292"), new ColumnSet(true));
                //ITracingService tracingService = null;
                //CallAllocation(_service, iJob, iJob, tracingService);
            }
        }

        static void CallAllocation(IOrganizationService service, Entity entity, Entity iJob, ITracingService tracingService)
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
        static EntityReference GetAssigneeForBoth(IOrganizationService service, Guid iPin, Guid iDivision, Guid iCallStype, Guid JobId, Entity iJob, EntityReference Default, Guid SalesOfficeId, ITracingService tracingService)
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
            if (Assignee.Id != Guid.Empty)
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
        static bool CheckIfUserActive(IOrganizationService service, EntityReference Assignee)
        {
            bool IfActive = false;
            SystemUser iUser = (SystemUser)service.Retrieve(SystemUser.EntityLogicalName, Assignee.Id, new ColumnSet("isdisabled"));
            if (iUser.IsDisabled == false)
            {
                IfActive = true;
            }
            return IfActive;
        }
        static EntityReference GetAssignee(IOrganizationService service, Guid iPin, Guid iDivision, Guid iCallStype, int iResource, Guid JobId, Entity iJob, EntityReference Default, Guid SalesOfficeId, ITracingService tracingService)
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
            if (Assignee.Id != Guid.Empty)
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

        static bool CheckIfDirectEngineerPresent(IOrganizationService service, Guid iEngineer)
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
        public class SBUBranchMapping
        {
            public Guid SBUBranchMappingId { get; set; }
            public EntityReference BranchHeadUser { get; set; }
        }
        static Guid GetFallBackRecord(IOrganizationService service, Guid iDivision, Guid iSalesOffice, Guid ownerId, Entity entity, int OperationCode)
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
            if (OperationCode == 2)
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

        static void FSMDataMigration()
        {
            try
            {
                int i = 0;
                int j = 0;
                int pagenumber = 1;
                while (true)
                {
                    string fetchXML = $@"<fetch version='1.0' page='{pagenumber}' output-format='xml-platform' mapping='logical' distinct='true'>
                    <entity name='product'>
                    <attribute name='name' />
                    <attribute name='productnumber' />
                    <attribute name='description' />
                    <attribute name='statecode' />
                    <attribute name='productstructure' />
                    <attribute name='productid' />
                    <attribute name='hil_materialgroup' />
                    <attribute name='hil_division' />
                    <attribute name='hil_amount' />
                    <order attribute='productnumber' descending='false' />
                    <link-entity name='hil_servicebom' from='hil_productcategory' to='productid' link-type='inner' alias='ac' />
                    </entity>
                    </fetch>";
                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    j = j + EntityList.Entities.Count;
                    Entity ent = null;
                    QueryExpression query = null;
                    EntityCollection entList = null;
                    foreach (var record in EntityList.Entities)
                    {
                        try
                        {
                            query = new QueryExpression("product");
                            query.ColumnSet = new ColumnSet(false);
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition("productid", ConditionOperator.Equal, record.Id);
                            entList = _serviceFSM.RetrieveMultiple(query);
                            if (entList.Entities.Count>0) {
                                i += 1;
                                Console.WriteLine(record.GetAttributeValue<string>("productnumber") + ": " + i.ToString() + "/" + j.ToString());
                                continue;
                            }

                            ent = new Entity(record.LogicalName, record.Id);
                            ent["name"] = record.GetAttributeValue<string>("name");
                            ent["productnumber"] = record.GetAttributeValue<string>("productnumber");
                            ent["description"] = record.GetAttributeValue<string>("description");
                            ent["price"] = record.GetAttributeValue<Money>("hil_amount");
                            ent["producttypecode"] = new OptionSetValue(5);//MODEL
                            if (record.Contains("hil_materialgroup"))
                            {
                                ent["hil_assetcategory"] = new EntityReference("msdyn_customerassetcategory", record.GetAttributeValue<EntityReference>("hil_materialgroup").Id);
                            }
                            ent["defaultuomscheduleid"] = new EntityReference("uomschedule", new Guid("ca3d0dcc-1332-4b87-87af-dd4d495c9fd6"));
                            ent["defaultuomid"] = new EntityReference("uom", new Guid("8368d3d7-1e1a-4db4-ac6a-a9ae6fac531a"));
                            ent["pricelevelid"] = new EntityReference("pricelevel", new Guid("3c111099-1b8c-ed11-81ac-6045bdaaae69"));
                            ent["quantitydecimal"] = 0;
                            ent["msdyn_fieldserviceproducttype"] = new OptionSetValue(690970000); // Field Service Inventory Type : Inventory
                            Guid recID = _serviceFSM.Create(ent);

                            Entity updateProd = new Entity(record.LogicalName, recID);
                            updateProd["pricelevelid"] = new EntityReference("pricelevel", new Guid("3c111099-1b8c-ed11-81ac-6045bdaaae69"));
                            _serviceFSM.Update(updateProd);

                            _serviceFSM.Execute(new SetStateRequest
                            {
                                EntityMoniker = new EntityReference(record.LogicalName, record.Id),
                                State = new OptionSetValue(0), //Status
                                Status = new OptionSetValue(1) //Status reason
                            });

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        i += 1;
                        Console.WriteLine(record.GetAttributeValue<string>("productnumber") + ": " + i.ToString() + "/" + j.ToString());
                    }
                    pagenumber = pagenumber + 1;
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }

        static void DeleteDuplicateClaimLines(IOrganizationService service)
        {
            string filePath = @"C:\Kuldeep khare\DuplicateClaimLine.xlsx";
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();

            Microsoft.Office.Interop.Excel.Application excelApp = new Excel.Application();
            if (excelApp != null)
            {
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];

                Excel.Range excelRange = excelWorksheet.UsedRange;
                Excel.Range range;
                string claimheader, claimcategory, jobid;
                QueryExpression Query1;
                EntityCollection entcoll1;

                for (int i = 2; i <= excelRange.Rows.Count; i++)
                {
                    try
                    {
                        Console.WriteLine(i.ToString() + "/" + excelRange.Rows.Count.ToString());
                        range = (excelWorksheet.Cells[i, 1] as Excel.Range);
                        claimheader = range.Value.ToString();
                        range = (excelWorksheet.Cells[i, 2] as Excel.Range);
                        claimcategory = range.Value.ToString();
                        range = (excelWorksheet.Cells[i, 3] as Excel.Range);
                        if (range.Value != null)
                        {
                            jobid = range.Value.ToString();
                            Query1 = new QueryExpression("hil_claimline");
                            Query1.ColumnSet = new ColumnSet("hil_name", "hil_claimheader", "hil_claimcategory", "hil_jobid", "statecode", "statuscode");
                            Query1.Criteria = new FilterExpression(LogicalOperator.And);
                            Query1.Criteria.AddCondition("hil_claimheader", ConditionOperator.Equal, new Guid(claimheader));
                            Query1.Criteria.AddCondition("hil_claimcategory", ConditionOperator.Equal, new Guid(claimcategory));
                            Query1.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, new Guid(jobid));
                            Query1.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            Query1.AddOrder("createdon", OrderType.Descending);
                            entcoll1 = service.RetrieveMultiple(Query1);
                            if (entcoll1.Entities.Count > 1)
                            {
                                Entity ent = entcoll1.Entities[0];
                                Console.WriteLine("Found.. Claim Header# " + ent.GetAttributeValue<EntityReference>("hil_claimheader").Name + " Claimcategory# " + ent.GetAttributeValue<EntityReference>("hil_claimcategory").Name + " Job# " + ent.GetAttributeValue<EntityReference>("hil_jobid").Name);
                                ent["hil_name"] = "Found Duplicate";
                                ent["statecode"] = new OptionSetValue(1);
                                ent["statuscode"] = new OptionSetValue(2);
                                service.Update(ent);
                            }
                        }
                        Console.WriteLine(i.ToString() + "/" + excelRange.Rows.Count.ToString());
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(i.ToString() + "/" + excelRange.Rows.Count.ToString() + ex.Message);
                    }
                }
            }
        }
        static void DeleteEmails()
        {
            try
            {
                int i = 0;
                int j = 0;

                /*
                <value>hil_tender -> 10681</value>
                <value>hil_tenderproduct -> 10682</value>
                <value>hil_orderchecklist -> 11116</value>
                <value>hil_orderchecklistproduct -> 11165</value>
                <value>hil_tenderbankguarantee -> 10800</value>
                <value>hil_oaheader -> 11114</value>
                <value>hil_oaproduct -> 11115</value> 
                <condition attribute='statuscode' operator='in'>
                            <value>3</value>
                            <value>5</value>
                            <value>2</value>
                            <value>8</value>
                            <value>4</value>
                          </condition>

                */
                string _fromDate = ConfigurationManager.AppSettings["FromDate"].ToString();
                string _toDate = ConfigurationManager.AppSettings["ToDate"].ToString();
                while (true)
                {
                    //regardingobjecttypecode

                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='email'>
                        <attribute name='subject' />
                        <attribute name='regardingobjecttypecode' />
                        <attribute name='regardingobjectid' />
                        <attribute name='createdon' />
                        <order attribute='createdon' descending='true' />
                        <filter type='and'>
                            <condition attribute='createdon' operator='on-or-after' value='2022-12-11' />
                            <condition attribute='createdon' operator='on-or-before' value='2022-12-20' />
                            <condition attribute='statuscode' operator='in'>
                                <value>6</value>
                                <value>2</value>
                                <value>4</value>
                                <value>8</value>
                            </condition>
                        </filter>
                      </entity>
                    </fetch>";

                    //string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    //  <entity name='email'>
                    //    <attribute name='subject' />
                    //    <attribute name='regardingobjecttypecode' />
                    //    <attribute name='regardingobjectid' />
                    //    <attribute name='createdon' />
                    //    <order attribute='createdon' descending='true' />
                    //    <filter type='and'>
                    //     <condition attribute='statuscode' operator='in'>
                    //        <value>3</value>
                    //        <value>5</value>
                    //        <value>2</value>
                    //        <value>8</value>
                    //        <value>4</value>
                    //      </condition>
                    //      <condition attribute='createdon' operator='olderthan-x-days' value='3' />
                    //      <condition attribute='regardingobjecttypecode' operator='not-in'>
                    //            <value>10681</value>
                    //            <value>10682</value>
                    //            <value>11116</value>
                    //            <value>11165</value>
                    //            <value>10800</value>
                    //            <value>11114</value>
                    //            <value>11115</value>
                    //      </condition>
                    //    </filter>
                    //  </entity>
                    //</fetch>";

                    //string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    //  <entity name='email'>
                    //    <attribute name='subject' />
                    //    <attribute name='regardingobjecttypecode' />
                    //    <attribute name='regardingobjectid' />
                    //    <attribute name='createdon' />
                    //    <order attribute='createdon' descending='false' />
                    //    <filter type='and'>
                    //     <condition attribute='statuscode' operator='in'>
                    //        <value>6</value>
                    //        <value>1</value>
                    //      </condition>
                    //      <condition attribute='createdon' operator='on-or-after' value='" + _fromDate  + @"' />
                    //      <condition attribute='createdon' operator='on-or-before' value='" + _toDate + @"' />
                    //      <condition attribute='regardingobjecttypecode' operator='not-in'>
                    //            <value>10681</value>
                    //            <value>10682</value>
                    //            <value>11116</value>
                    //            <value>11165</value>
                    //            <value>10800</value>
                    //            <value>11114</value>
                    //            <value>11115</value>
                    //        </condition>
                    //    </filter>
                    //  </entity>
                    //</fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    j = j + EntityList.Entities.Count;
                    foreach (var record in EntityList.Entities)
                    {
                        EntityReference ent = record.GetAttributeValue<EntityReference>("regardingobjectid");
                        if (ent != null)
                        {
                            try
                            {
                                if (ent.LogicalName != "hil_orderchecklist" && ent.LogicalName != "hil_orderchecklistproduct" && ent.LogicalName != "hil_tender" && ent.LogicalName != "hil_tenderproduct" && ent.LogicalName != "hil_oaheader" && ent.LogicalName != "hil_oaproduct" && ent.LogicalName != "hil_tenderbankguarantee")
                                    _service.Delete("email", record.Id);
                            }
                            catch (Exception ex) { Console.WriteLine(ex.Message); }
                            Console.WriteLine(record.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString() + ": " + record.GetAttributeValue<EntityReference>("regardingobjectid").LogicalName + ":" + record.GetAttributeValue<string>("subject") + ":" +  i.ToString() + "/" + j.ToString());
                        }
                        else
                        {
                            if (record.Contains("subject"))
                            {
                                if (record.GetAttributeValue<string>("subject").IndexOf("TEND") < 0)
                                {
                                    _service.Delete("email", record.Id);
                                    Console.WriteLine(record.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString() + ": " + record.GetAttributeValue<string>("subject") + ": " + i.ToString() + "/" + j.ToString());
                                }
                            }
                            else {
                                _service.Delete("email", record.Id);
                                Console.WriteLine(record.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString() + ": " + i.ToString() + "/" + j.ToString());
                            }
                        }
                        i += 1;
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }

        static void DeleteAuditLog_DeleteEntry()
        {
            try
            {
                //https://havells.api.crm8.dynamics.com/api/data/v9.1/EntityDefinitions?$select=LogicalName,ObjectTypeCode
                //<condition attribute='objecttypecode' operator='eq' value='10196' />
                /*<condition attribute='action' operator='not-in'>
                <value>1</value> Create
                <value>2</value> Update
                <value>4</value> Activate
                <value>5</value> Deactivate
                <value>13</value> Assign
                </condition>
                 */
                int i = 0;
                int j = 0;
                while (true)
                {
                    string fetchXML = @"<fetch top='1000'>
                    <entity name='audit'>
                    <attribute name='objectid' />
                    <attribute name='action' />
                    <attribute name='objecttypecode' />
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='createdon' operator='on-or-after' value='2021-01-01'/>
                        <condition attribute='createdon' operator='on-or-before' value='2021-01-05'/>
                        <condition attribute='action' operator='in'>
                        <value>3</value>
                        <value>106</value>
                        <value>107</value>
                        <value>108</value>
                        <value>109</value>
                        <value>110</value>
                        <value>111</value>
                        </condition>
                    </filter>
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    j = j + EntityList.Entities.Count;

                    foreach (var record in EntityList.Entities)
                    {
                        var EntityLogReference = ((EntityReference)record.Attributes["objectid"]);

                        var DeleteAudit = new DeleteRecordChangeHistoryRequest();
                        Console.WriteLine("Object Type: " + EntityLogReference.LogicalName + "::" + record.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString());
                        if (EntityLogReference.Id != Guid.Empty && EntityLogReference.LogicalName != "adx_webpage" && EntityLogReference.LogicalName != "adx_webrole")
                        {
                            try
                            {
                                Console.WriteLine("Deleting Log: " + record.GetAttributeValue<OptionSetValue>("action").Value + " / " + EntityLogReference.LogicalName + " / " + EntityLogReference.Name);
                                DeleteAudit.Target = (EntityReference)record.Attributes["objectid"];
                                _service.Execute(DeleteAudit);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                        }
                        i += 1;
                        Console.WriteLine(i.ToString() + "/" + j.ToString());
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }
        static void DeletePlugInTraces()
        {
            try
            {
                int i = 0;
                int j = 0;
                //<condition attribute='createdon' operator='on' value='2020-09-01' />
                //<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>

                while (true)
                {
                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='plugintracelog'>
                    <attribute name='plugintracelogid' />
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='false' />
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    j = j + EntityList.Entities.Count;
                    foreach (var record in EntityList.Entities)
                    {
                        try
                        {
                            _service.Delete("plugintracelog", record.Id);
                        }
                        catch (Exception ex) { Console.WriteLine(ex.Message); }
                        i += 1;
                        Console.WriteLine(i.ToString() + "/" + j.ToString() + ": " + record.GetAttributeValue<DateTime>("createdon").ToString());
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }

        static void PrincipalObjectAttributeAccess()
        {
            try
            {
                int i = 0;
                int j = 0;

                while (true)
                {
                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='principalobjectaccess'>
                    <all-attributes/>
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    j = j + EntityList.Entities.Count;
                    foreach (var record in EntityList.Entities)
                    {
                        try
                        {
                            _service.Delete("principalobjectaccess", record.Id);
                        }
                        catch (Exception ex) { Console.WriteLine(ex.Message); }
                        i += 1;
                        Console.WriteLine(i.ToString() + "/" + j.ToString() + ": " + record.GetAttributeValue<DateTime>("changedon").ToString());
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }
        static void DeleteAsyncOperations() {
            try
            {
                int i = 0;
                int j = 0;
                //<condition attribute='createdon' operator='on' value='2020-09-01' />
                //<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>

                while (true)
                {
                    //string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    //<entity name='asyncoperation'>
                    //<attribute name='asyncoperationid' />
                    //<attribute name='name' />
                    //<attribute name='regardingobjectid' />
                    //<attribute name='operationtype' />
                    //<attribute name='statuscode' />
                    //<attribute name='ownerid' />
                    //<attribute name='startedon' />
                    //<attribute name='statecode' />
                    //<attribute name='createdon' />
                    //<attribute name='modifiedon' />
                    //<attribute name='message' />
                    //<attribute name='friendlymessage' />
                    //<order attribute='createdon' descending='false' />
                    //<filter type='and'>
                    //    <condition attribute='statuscode' operator='in'>
                    //    <value>31</value>
                    //    <value>30</value>
                    //    <value>32</value>
                    //    </condition>
                    //    <condition attribute='recurrencestarttime' operator='null' />
                    //</filter>
                    //</entity>
                    //</fetch>";

                    //string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    //<entity name='asyncoperation'>
                    //<attribute name='asyncoperationid' />
                    //<attribute name='name' />
                    //<attribute name='regardingobjectid' />
                    //<attribute name='operationtype' />
                    //<attribute name='statuscode' />
                    //<attribute name='ownerid' />
                    //<attribute name='startedon' />
                    //<attribute name='statecode' />
                    //<attribute name='createdon' />
                    //<attribute name='modifiedon' />
                    //<attribute name='message' />
                    //<attribute name='friendlymessage' />
                    //<order attribute='createdon' descending='false' />
                    //<filter type='and'>
                    //    <condition attribute='statuscode' operator='eq' value='10' />
                    //    <condition attribute='recurrencepattern' operator='null' />
                    //    <condition attribute='operationtype' operator='eq' value='10' />
                    //    <condition attribute='createdon' operator='on' value='2020-10-14' />
                    //</filter>
                    //</entity>
                    //</fetch>";

                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='asyncoperation'>
                    <attribute name='name' />
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='recurrencepattern' operator='null' />
                        <condition attribute='statuscode' operator='in'>
                            <value>32</value>
                            <value>31</value>
                            <value>30</value>
                        </condition>
                    </filter>
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    //if (EntityList.Entities.Count == 0) { break; }
                    j = j + EntityList.Entities.Count;
                    foreach (var record in EntityList.Entities)
                    {
                        try
                        {
                            _service.Delete("asyncoperation", record.Id);
                        }
                        catch (Exception ex) { Console.WriteLine(ex.Message); }
                        i += 1;
                        Console.WriteLine(record.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString() + ": " + record.GetAttributeValue<string>("name") + ": " + i.ToString() + "/" + j.ToString());
                        //Console.WriteLine(i.ToString() + "/" + j.ToString() + ": " + record.GetAttributeValue<DateTime>("createdon").ToString());
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }
        static void DeleteWorkFlowBase()
        {
            try
            {
                int i = 0;
                int j = 0;
                /*
                    <condition attribute='createdon' operator='on-or-before' value='2022-06-25' />
                    <condition attribute='statuscode' operator='in'>
                    <value>31</value>
                    <value>30</value>
                <filter type='and'>
                        <condition attribute='createdon' operator='on-or-after' value='2023-03-01' />
                    </filter>
                    </condition>
                 */
                while (true)
                {
                    //string fetchXML = @"<fetch top='5000'>
                    //    <entity name='asyncoperation'>
                    //        <attribute name='asyncoperationid' />
                    //        <attribute name='createdon' />
                    //        <attribute name='name' />
                    //    <order attribute='createdon' descending='true' />
                    //    <filter type='and'>
                    //        <condition attribute='statuscode' operator='eq' value='10' />
                    //        <condition attribute='createdon' operator='olderthan-x-days' value='1' />
                    //        <condition attribute='recurrencepattern' operator='null' />
                    //    </filter>
                    //    </entity>
                    //    </fetch>";
                    string fetchXML = @"<fetch top='5000'>
                        <entity name='workflowlog'>
                            <attribute name='createdon' />
                            <attribute name='name' />
                        <order attribute='createdon' descending='true' />
                        </entity>
                        </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    j = j + EntityList.Entities.Count;
                    foreach (var record in EntityList.Entities)
                    {
                        try
                        {
                            _service.Delete("asyncoperation", record.Id);
                            //Console.WriteLine("Success");
                        }
                        catch (Exception ex) { Console.WriteLine(ex.Message); }
                        i += 1;
                        Console.WriteLine(record.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString() + ": " + record.GetAttributeValue<string>("name") + ": " + i.ToString() + "/" + j.ToString());
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }

        static void DeleteProcessSessions()
        {
            try
            {
                int i = 0;
                int j = 0;

                while (true)
                {
                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='processsession'>
                    <attribute name='name' />
                    <attribute name='processsessionid' />
                    <attribute name='statuscode' />
                    <attribute name='startedon' />
                    <attribute name='startedby' />
                    <attribute name='regardingobjectid' />
                    <attribute name='ownerid' />
                    <attribute name='stepname' />
                    <attribute name='statecode' />
                    <attribute name='createdon' />
                    <attribute name='createdby' />
                    <attribute name='completedon' />
                    <attribute name='completedby' />
                    <attribute name='activityname' />
                    <order attribute='createdon' descending='false' />
                    <link-entity name='workflow' from='workflowid' to='processid' visible='false' link-type='outer' alias='wf'>
                        <attribute name='name' />
                        <attribute name='category' />
                    </link-entity>
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    j = j + EntityList.Entities.Count;
                    foreach (var record in EntityList.Entities)
                    {
                        try
                        {
                            _service.Delete("processsession", record.Id);
                        }
                        catch (Exception ex) { Console.WriteLine(ex.Message); }
                        i += 1;
                        Console.WriteLine(record.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString() + ": " + record.GetAttributeValue<string>("name") + ": " + i.ToString() + "/" + j.ToString());
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }

        static void DeleteImportFile()
        {
            try
            {
                int i = 0;
                int j = 0;

                while (true)
                {
                    try
                    {
                        QueryExpression qryExp = new QueryExpression("importfile");
                        qryExp.ColumnSet = new ColumnSet("name", "statuscode", "createdon");
                        qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qryExp.Criteria.AddCondition("statuscode", ConditionOperator.In, new object[] { 0, 4, 5 });
                        qryExp.AddOrder("createdon", OrderType.Ascending);
                        qryExp.TopCount = 1000;
                        EntityCollection EntityList = _service.RetrieveMultiple(qryExp);
                        if (EntityList.Entities.Count == 0) { break; }
                        j = j + EntityList.Entities.Count;
                        if (EntityList.Entities.Count > 0)
                        {
                            foreach (Entity ent in EntityList.Entities)
                            {
                                Console.WriteLine(i++.ToString() + "/" + j.ToString() + ent.GetAttributeValue<DateTime>("createdon").ToString() + "/" + ent.GetAttributeValue<string>("name"));
                                _service.Delete(ent.LogicalName, ent.Id);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("ERROR!!! "+ ex.Message);
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }
        static void DeleteAuditLog()
        {
            try
            {
                //https://havells.api.crm8.dynamics.com/api/data/v9.1/EntityDefinitions?$select=LogicalName,ObjectTypeCode
                /*
                 * <condition attribute='createdon' operator='on-or-after' value='2020-01-01' />
                   <condition attribute='createdon' operator='on-or-before' value='2020-09-01' />
                    <condition attribute='objectid' operator='ne' value='{00000000-0000-0000-0000-000000000000}' />       
                   Consumer -24/Sep/2020
                && EntityLogReference.LogicalName !="msdyn_workorder"
                  <condition attribute='objecttypecode' operator='not-in'>
                        <value>10196</value>
                        <value>10215</value>
                        <value>10216</value>
                        </condition>

                  <condition attribute='objecttypecode' operator='not-in'>
                        <value>10196</value>
                        <value>10216</value>
                        <value>10215</value>
                        </condition>

                 */
                int i = 0;
                int j = 0;
                while (true)
                {
                    string fetchXML =
                    @"<fetch top='1000'>
                    <entity name='audit'>
                    <attribute name='objectid' />
                    <attribute name='action' />
                    <attribute name='objecttypecode' />
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='objecttypecode' operator='eq' value='11165' />                   
                    </filter>
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    j = j + EntityList.Entities.Count;

                    foreach (var record in EntityList.Entities)
                    {
                        var EntityLogReference = ((EntityReference)record.Attributes["objectid"]);
                        var DeleteAudit = new DeleteRecordChangeHistoryRequest();
                        Console.WriteLine("Object Type: " + EntityLogReference.LogicalName + "::" + record.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString());
                        if (EntityLogReference.Id != Guid.Empty)
                        {
                            try
                            {
                                Console.WriteLine("Deleting Log: " + record.GetAttributeValue<OptionSetValue>("action").Value + " / " + EntityLogReference.LogicalName + " / " + EntityLogReference.Name);
                                DeleteAudit.Target = (EntityReference)record.Attributes["objectid"];
                                _service.Execute(DeleteAudit);
                            }
                            catch (Exception ex)
                            { 
                                Console.WriteLine(ex.Message); 
                            }
                        }
                        i += 1;
                        Console.WriteLine(i.ToString() + "/" + j.ToString());
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }
        static void DeleteCurrentAuditLog()
        {
            try
            {
                int i = 0;
                int j = 0;
                while (true)
                {
                    string fetchXML = @"<fetch top='1000'>
                    <entity name='audit'>
                    <attribute name='objectid' />
                    <attribute name='action' />
                    <attribute name='objecttypecode' />
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='objecttypecode' operator='in'>
                        <value>8</value>
                        </condition>
                        <condition attribute='createdon' operator='on-or-after' value='2020-01-01' />
                        <condition attribute='createdon' operator='on-or-before' value='2020-01-01' />
                        <condition attribute='objectid' operator='ne' value='{00000000-0000-0000-0000-000000000000}' />       
                    </filter>
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    j = j + EntityList.Entities.Count;

                    foreach (var record in EntityList.Entities)
                    {
                        var EntityLogReference = ((EntityReference)record.Attributes["objectid"]);

                        var DeleteAudit = new DeleteRecordChangeHistoryRequest();
                        Console.WriteLine("Object Type: " + EntityLogReference.LogicalName + "::" + record.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString());
                        if (EntityLogReference.Id != Guid.Empty && EntityLogReference.LogicalName != "msdyn_workorder" && EntityLogReference.LogicalName != "msdyn_customerasset")
                        {
                            try
                            {
                                Console.WriteLine("Deleting Log: " + record.GetAttributeValue<OptionSetValue>("action").Value + " / " + EntityLogReference.LogicalName + " / " + EntityLogReference.Name);
                                DeleteAudit.Target = (EntityReference)record.Attributes["objectid"];
                                _service.Execute(DeleteAudit);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                        }
                        i += 1;
                        Console.WriteLine(i.ToString() + "/" + j.ToString());
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }
        static void DeleteAuditLog_Entitywise()
        {
            try
            {
                //https://havells.api.crm8.dynamics.com/api/data/v9.1/EntityDefinitions?$select=LogicalName,ObjectTypeCode
                //<condition attribute='createdon' operator='on-or-after' value='2020-01-01' />
                //<condition attribute='objectid' operator='ne' value='{00000000-0000-0000-0000-000000000000}' /> 
                //< condition attribute = 'createdon' operator= 'on-or-after' value = '2019-01-01' />
                //< condition attribute = 'createdon' operator= 'on-or-before' value = '2021-02-22' />
                /*
                contact	2
                systemuser	8
                product	1024
                cardtype	9983
                msdyn_approval	10052
                msdyn_customerasset	10144
                msdyn_timeoffrequest	10193
                msdyn_workorder	10196
                msdyn_workorderincident	10199
                msdyn_workorderproduct	10200
                adx_webpage	10215
                adx_webrole	10216
                hil_integrationconfiguration 10375
                hil_assignmentmatrix	10456
                hil_inventory	10469
                hil_inventoryjournal	10470
                hil_inventoryrequest	10471
                hil_sbubranchmapping	10488
                hil_technician	10514
                hil_postatustracker	10564
                hil_alternatepart	10568
                hil_partnerdepartment	10598
                msdyn_workorderservice	10202
                <condition attribute='objecttypecode' operator='in'>
                <value>10144</value>
                </condition>
                 */

                //string _fromDate = ConfigurationManager.AppSettings["FromDate"].ToString();
                //string _toDate = ConfigurationManager.AppSettings["ToDate"].ToString();
                //string _objectTypeCode = ConfigurationManager.AppSettings["ObjectTypeCode"].ToString();
                //string _topCount = ConfigurationManager.AppSettings["TopCount"].ToString();
                int i = 0;
                int j = 0;
                while (true)
                {
                    string fetchXML = @"<fetch top='1000'>
                    <entity name='audit'>
                    <attribute name='objectid' />
                    <attribute name='action' />
                    <attribute name='objecttypecode' />
                    <attribute name='createdon' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                        <condition attribute='createdon' operator='on-or-after' value='2021-01-01' />
                        <condition attribute='createdon' operator='on-or-before' value='2021-01-03' />
                        <condition attribute='objectid' operator='ne' value='{00000000-0000-0000-0000-000000000000}' /> 
                    </filter>
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) {break;}
                    j = j + EntityList.Entities.Count;

                    foreach (var record in EntityList.Entities)
                    {
                        var EntityLogReference = ((EntityReference)record.Attributes["objectid"]);

                        var DeleteAudit = new DeleteRecordChangeHistoryRequest();
                        Console.WriteLine("Object Type: " + EntityLogReference.LogicalName + "::" + record.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString());
                        if (EntityLogReference.Id != Guid.Empty && EntityLogReference.LogicalName != "adx_webpage" && EntityLogReference.LogicalName != "adx_webrole")
                        {
                            try
                            {
                                Console.WriteLine("Deleting Log: " + record.FormattedValues["action"].ToString() + " / " + EntityLogReference.LogicalName + " / " + EntityLogReference.Name);
                                DeleteAudit.Target = (EntityReference)record.Attributes["objectid"];
                                _service.Execute(DeleteAudit);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                        }
                        i += 1;
                        Console.WriteLine(i.ToString() + "/" + j.ToString());
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }
        static string GetObjectTypeCode(string entitylogicalname)
        {
            Entity entity = new Entity(entitylogicalname);
            RetrieveEntityRequest EntityRequest = new RetrieveEntityRequest();
            EntityRequest.LogicalName = entity.LogicalName;
            EntityRequest.EntityFilters = EntityFilters.Entity;
            RetrieveEntityResponse responseent = (RetrieveEntityResponse)_service.Execute(EntityRequest);
            EntityMetadata ent = (EntityMetadata)responseent.EntityMetadata;
            string ObjectTypeCode = ent.ObjectTypeCode.ToString();
            return ObjectTypeCode;
        }
        static void updateAppconfig(string ent)
        {
            ConfigurationManager.RefreshSection("appSettings");
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings["FromDate"].Value = ent;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }
        static void GetDistinctEntityInAuditLog()
        {
            try
            {
                //https://havells.api.crm8.dynamics.com/api/data/v9.1/EntityDefinitions?$select=LogicalName,ObjectTypeCode
                //2021-02-19 T13:00:00
                //<condition attribute='objecttypecode' operator='eq' value='10375' />
                //10144 -Asset, 10196- WorkOrder
                //<condition attribute='objecttypecode' operator='eq' value='10144' />
                //<condition attribute='createdon' operator='last-x-hours' value='1' />

                int i = 1;
                //while (true)
                //{
                //string fetchXML = @"<fetch mapping='logical' distinct='true' version='1.0'>
                //<entity name='audit'>
                //<attribute name='objecttypecode'/>
                //<filter type='and'>
                //    <condition attribute='createdon' operator='yesterday' />
                //    <condition attribute='objecttypecode' operator='not-in'>
                //    <value>10144</value>
                //    <value>10196</value>
                //    <value>2</value>
                //    </condition>
                //    <condition attribute='objectid' operator='ne' value='{00000000-0000-0000-0000-000000000000}' />       
                //</filter>
                //</entity>
                //</fetch>";

                string fetchXML = @"<fetch mapping='logical' version='1.0'>
                <entity name='audit'>
                <attribute name='objectid' />
                <attribute name='action' />
                <attribute name='objecttypecode' />
                <attribute name='createdon' />
                <filter type='and'>
                    <condition attribute='createdon' operator='yesterday' />
                    <condition attribute='objecttypecode' operator='in'>
                        <value>10144</value>
                        </condition>
                    <condition attribute='action' operator='ne' value='111' />
                </filter>
                </entity>
                </fetch>";

                EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    //if (EntityList.Entities.Count == 0) { break; }
                    foreach (var record in EntityList.Entities)
                    {
                        string _objectTypeCode = record.Attributes["objecttypecode"].ToString();
                        string _action = ((OptionSetValue)record.Attributes["action"]).Value.ToString();
                    Console.WriteLine("Audit Log Migration of Entity: " + _action + ":" + _objectTypeCode);
                        //DateTime _createdOnMin = GetEntityMinDateOfAuditLog(_objectTypeCode);
                        //D365AuditLogMigration_Delta(_objectTypeCode);
                    }
                //}
            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
            }
        }
        static DateTime GetEntityMinDateOfAuditLog(string _objectType)
        {
            try
            {
                //<condition attribute='createdon' operator='last-x-hours' value='1' />
                string _objectTypeCode = GetObjectTypeCode(_objectType);
                string fetchXML = @"<fetch top='1'>
                <entity name='audit'>
                <attribute name='createdon' />
                <order attribute='createdon' descending='false' />
                <filter type='and'>
                    <condition attribute='objecttypecode' operator='eq' value='" + _objectTypeCode + @"' />
                    <condition attribute='objectid' operator='ne' value='{00000000-0000-0000-0000-000000000000}' />       
                </filter>
                </entity>
                </fetch>";

                EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                if (EntityList.Entities.Count == 0) { return new DateTime(1900, 1, 1); }
                return EntityList.Entities[0].GetAttributeValue<DateTime>("createdon").AddMinutes(330);
            }
            catch (Exception ex)
            {
                return new DateTime(1900, 1, 1);
            }
        }
        static List<EntityAttribute> GetEntityAttributes(string _entityLogicalName)
        {
            List<EntityAttribute> lstAttribute = new List<EntityAttribute>();

            RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.Attributes,
                LogicalName = _entityLogicalName,
                RetrieveAsIfPublished = true
            };

            RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)_service.Execute(retrieveEntityRequest);
            
            EntityMetadata AccountEntity = retrieveAccountEntityResponse.EntityMetadata;

            foreach (object attribute in AccountEntity.Attributes)
            {
                AttributeMetadata a = (AttributeMetadata)attribute;
                if (a.AttributeTypeName.Value != "VirtualType")
                {
                    EntityAttribute entAttr = new EntityAttribute();
                    entAttr.LogicalName = a.LogicalName;
                    if (a.DisplayName.LocalizedLabels.Count > 0)
                    { entAttr.DisplayName = a.DisplayName.LocalizedLabels[0].Label; }
                    else
                    { entAttr.DisplayName = a.LogicalName; }
                    entAttr.AttributeType = a.AttributeTypeName.Value.ToString();
                    if (entAttr.AttributeType.IndexOf("Pick") >= 0 || entAttr.AttributeType.IndexOf("State") >= 0 || entAttr.AttributeType.IndexOf("Status") >= 0)
                    {
                        entAttr.Optionset = new List<EntityOptionSet>();
                        var retrieveAttributeRequest = new RetrieveAttributeRequest
                        {
                            EntityLogicalName = _entityLogicalName,
                            LogicalName = entAttr.LogicalName
                        };
                        var retrieveAttributeResponse = (RetrieveAttributeResponse)_service.Execute(retrieveAttributeRequest);
                        var retrievedPicklistAttributeMetadata = (EnumAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
                        foreach (OptionMetadata opt in retrievedPicklistAttributeMetadata.OptionSet.Options)
                        {
                            entAttr.Optionset.Add(new EntityOptionSet() { OptionName = opt.Label.LocalizedLabels[0].Label, OptionValue = opt.Value });
                        }
                    }
                    lstAttribute.Add(entAttr);
                }
            }

            return lstAttribute;
        }
        static void D365AuditLogMigration() {
            int i = 1;
            Console.WriteLine("Fetching Entity Attribute Metadata..." + DateTime.Now.ToString());
            List<EntityAttribute> lstEntityAttribute = GetEntityAttributes("msdyn_workorder");
            Console.WriteLine("Entity Attribute Metadata successfully fetched..." + DateTime.Now.ToString());
            Entity entTemp = null;
            string entityLogicalName = "msdyn_workorder";
            string[] primaryKeySchemaName = { "msdyn_workorderid", "msdyn_name" };

            QueryExpression Query = new QueryExpression(entityLogicalName);
            Query.ColumnSet = new ColumnSet(primaryKeySchemaName);
            //Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, "2018-12-31");
            Query.Criteria.AddCondition("hil_archivelog", ConditionOperator.NotEqual, true);
            Query.AddOrder("createdon", OrderType.Ascending);
            Query.PageInfo = new PagingInfo();
            Query.PageInfo.Count = 5000;
            Query.PageInfo.PageNumber = 1;
            Query.PageInfo.ReturnTotalRecordCount = true;

            try
            {
                EntityCollection ec = _service.RetrieveMultiple(Query);
                do
                {
                    if (ec.Entities.Count > 0)
                    {
                        string accName = "d365storagesa";
                        string accKey = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
                        //bool _retValue = false;
                        // Implement the accout, set true for https for SSL.  
                        StorageCredentials creds = new StorageCredentials(accName, accKey);
                        CloudStorageAccount strAcc = new CloudStorageAccount(creds, true);
                        // Create the blob client.  
                        CloudBlobClient blobClient = strAcc.CreateCloudBlobClient();
                        // Retrieve a reference to a container.   
                        CloudBlobContainer container = blobClient.GetContainerReference("d365auditlogarchive");
                        // Create the container if it doesn't already exist.  
                        container.CreateIfNotExistsAsync();

                        foreach (Entity ent in ec.Entities)
                        {
                            if (CheckIfAzureBlobExist(ent.Id.ToString(), container))
                            {
                                Console.WriteLine("Job # " + i.ToString() + " " + ent.GetAttributeValue<string>("msdyn_name"));
                                entTemp = new Entity("msdyn_workorder", ent.Id);
                                entTemp["hil_archivelog"] = true;
                                _service.Update(entTemp);
                                i += 1;
                                continue;
                            }
                            Console.WriteLine("Job # " + i.ToString() + " " + ent.GetAttributeValue<string>("msdyn_name") + " Start..." + DateTime.Now.ToString());
                            RetrieveRecordChangeHistoryRequest changeRequest = new RetrieveRecordChangeHistoryRequest();
                            changeRequest.Target = new EntityReference(entityLogicalName, ent.Id);
                            RetrieveRecordChangeHistoryResponse changeResponse =
                            (RetrieveRecordChangeHistoryResponse)_service.Execute(changeRequest);

                            AuditDetailCollection auditDetailCollection = changeResponse.AuditDetailCollection;
                            StringBuilder strContent = new StringBuilder();
                            strContent.AppendLine("auditid,createdon,action,actionname,objectid,objecttypecode,objecttypecodename,operation,operationname,userid,useridname,fieldname,oldvalue,oldvaluename,oldvaluetype,newvalue,newvaluename,newvaluetype");
                            string strLine;
                            string strAttributes;
                            string dateString;
                            foreach (var attrAuditDetail in auditDetailCollection.AuditDetails)
                            {
                                var auditRecord = attrAuditDetail.AuditRecord;
                                strLine = string.Empty;
                                dateString = string.Empty;
                                if (auditRecord.Attributes.Contains("auditid"))
                                {
                                    strLine = auditRecord.Attributes["auditid"].ToString();
                                }
                                else
                                {
                                    strLine = null;
                                }
                                if (auditRecord.Attributes.Contains("createdon"))
                                {
                                    DateTime dt = ((DateTime)auditRecord.Attributes["createdon"]).AddMinutes(330);
                                    strLine += "," + dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0') + dt.Hour.ToString().PadLeft(2, '0') + dt.Minute.ToString().PadLeft(2, '0') + dt.Second.ToString().PadLeft(2, '0');
                                }
                                else
                                {
                                    strLine += "," + string.Empty;
                                }
                                if (auditRecord.Attributes.Contains("action"))
                                {
                                    strLine += "," + ((OptionSetValue)auditRecord.Attributes["action"]).Value.ToString();
                                    strLine += "," + auditRecord.FormattedValues["action"];
                                }
                                else
                                {

                                    strLine += "," + string.Empty + "," + string.Empty;
                                }
                                if (auditRecord.Attributes.Contains("objectid"))
                                {
                                    strLine += "," + ((EntityReference)auditRecord.Attributes["objectid"]).Id.ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("objecttypecode"))
                                {
                                    strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                    strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("operation"))
                                {
                                    strLine += "," + ((OptionSetValue)auditRecord.Attributes["operation"]).Value.ToString();
                                    strLine += "," + auditRecord.FormattedValues["operation"];
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("userid"))
                                {
                                    EntityReference er = auditRecord.GetAttributeValue<EntityReference>("userid");
                                    strLine += "," + er.Id.ToString();
                                    strLine += "," + er.Name.ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }
                                strLine += ",";
                                strAttributes = string.Empty;
                                if (attrAuditDetail.GetType().ToString() != "Microsoft.Crm.Sdk.Messages.AuditDetail")
                                {
                                    Dictionary<string, Object> oldValues = new Dictionary<string, object>();
                                    Dictionary<string, Object> newValues = new Dictionary<string, object>();
                                    object _OldValue;
                                    //object _NewValue;
                                    var newValueEntity = ((AttributeAuditDetail)attrAuditDetail).NewValue;
                                    if (newValueEntity != null)
                                    {
                                        newValues = newValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                    }

                                    var oldValueEntity = ((AttributeAuditDetail)attrAuditDetail).OldValue;
                                    if (oldValueEntity != null)
                                    {
                                        oldValues = oldValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                    }
                                    if (newValues.Count > 0)
                                    {
                                        foreach (KeyValuePair<string, object> de in newValues)
                                        {
                                            strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                            //strAttributes = de.Key.ToString() + ",";
                                            //Console.WriteLine("Attribute Name :" + strAttributes);
                                            if (oldValues.ContainsKey(de.Key))
                                            {
                                                _OldValue = oldValues[de.Key.ToString()];
                                                if (_OldValue.GetType() == typeof(Decimal))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Double))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(System.DateTime))
                                                {
                                                    DateTime dtValue = ((DateTime)_OldValue).AddMinutes(330);
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                {
                                                    EntityReference entRef = (EntityReference)_OldValue;
                                                    if (entRef != null)
                                                    {
                                                        strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                    }
                                                    else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                }
                                                else if (_OldValue.GetType() == typeof(System.Boolean))
                                                {
                                                    string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                    strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(System.Int32))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                {
                                                    Money money = (Money)_OldValue;
                                                    strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                {
                                                    OptionSetValue optValue = (OptionSetValue)_OldValue;
                                                    EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                    EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                    strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(System.String))
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                }

                                                oldValues.Remove(de.Key.ToString());
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                            }

                                            if (de.Value.GetType() == typeof(Decimal))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Double))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.DateTime))
                                            {
                                                DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                            {
                                                EntityReference entRef = (EntityReference)de.Value;
                                                if (entRef != null)
                                                {
                                                    strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                }
                                                else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                            }
                                            else if (de.Value.GetType() == typeof(System.Boolean))
                                            {
                                                string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.Int32))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                            {
                                                Money money = (Money)de.Value;
                                                strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                            {
                                                OptionSetValue optValue = (OptionSetValue)de.Value;
                                                //string name = string.Empty;
                                                //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.String))
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }

                                            if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                            strContent.AppendLine(strLine + strAttributes);
                                        }
                                    }
                                    if (oldValues.Count > 0)
                                    {
                                        foreach (KeyValuePair<string, object> de in oldValues)
                                        {
                                            //strAttributes = de.Key.ToString() + ",";
                                            strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                            if (de.Value.GetType() == typeof(Decimal))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Double))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.DateTime))
                                            {
                                                DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                            {
                                                EntityReference entRef = (EntityReference)de.Value;
                                                if (entRef != null)
                                                {
                                                    strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                }
                                                else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                            }
                                            else if (de.Value.GetType() == typeof(System.Boolean))
                                            {
                                                string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.Int32))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                            {
                                                Money money = (Money)de.Value;
                                                strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                            {
                                                OptionSetValue optValue = (OptionSetValue)de.Value;
                                                //string name = string.Empty;
                                                //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.String))
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                            if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                            strContent.AppendLine(strLine + strAttributes);
                                        }
                                    }
                                }
                                else
                                {
                                    strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty;
                                    strContent.AppendLine(strLine + strAttributes);
                                }
                            }
                            //UploadAzureBlob(ent.Id.ToString(), strContent.ToString());
                            UploadAzureBlob(ent.Id.ToString(), strContent.ToString(), container);

                            var DeleteAudit = new DeleteRecordChangeHistoryRequest();
                            DeleteAudit.Target = (EntityReference)ent.ToEntityReference();
                            try
                            {
                                _service.Execute(DeleteAudit);
                                entTemp = new Entity("msdyn_workorder", ent.Id);
                                entTemp["hil_archivelog"] = true;
                                _service.Update(entTemp);
                            }
                            catch { }
                            Console.WriteLine("Job # " + i.ToString() + " End... " + DateTime.Now.ToString());
                            i += 1;
                        }

                        #region Delete Audit Log

                        #endregion
                    }
                    Query.PageInfo.PageNumber += 1;
                    Query.PageInfo.PagingCookie = ec.PagingCookie;
                    ec = _service.RetrieveMultiple(Query);
                } while (ec.MoreRecords);
                Console.WriteLine("DONE !!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        static void D365AuditLogMigration_ActiveSystemUser()
        {
            int i = 1;
            Console.WriteLine("Fetching Entity Attribute Metadata..." + DateTime.Now.ToString());
            List<EntityAttribute> lstEntityAttribute = GetEntityAttributes("systemuser");
            Console.WriteLine("Entity Attribute Metadata successfully fetched..." + DateTime.Now.ToString());
            Entity entTemp = null;
            string entityLogicalName = "systemuser";
            string[] primaryKeySchemaName = { "systemuserid", "internalemailaddress" };
            DateTime _createdOn = new DateTime(1900, 1, 1);

            string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='audit'>
            <attribute name='objecttypecode' />
            <filter type='and'>
                <condition attribute='objecttypecode' operator='eq' value='8' />
                <condition attribute='createdon' operator='on-or-before' value='2020-09-01' />
            </filter>
            </entity>
            </fetch>";

            EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
            if (EntityList.Entities.Count > 0) {
                _createdOn = EntityList.Entities[0].GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                //_createdOn = EntityList.Entities[0].GetAttributeValue<DateTime>("ad.createdon").AddMinutes(330);
            }

            QueryExpression Query = new QueryExpression(entityLogicalName);
            Query.ColumnSet = new ColumnSet(primaryKeySchemaName);
            Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, _createdOn.Year.ToString() + "-" + _createdOn.Month.ToString().PadLeft(2, '0') + "-" + _createdOn.Day.ToString().PadLeft(2, '0'));
            Query.Criteria.AddCondition("isdisabled", ConditionOperator.Equal, false);
            Query.Criteria.AddCondition("accessmode", ConditionOperator.NotEqual, 3);
            Query.Criteria.AddCondition("accessmode", ConditionOperator.NotEqual, 5);
            Query.Criteria.AddCondition("fullname", ConditionOperator.NotEqual, "SYSTEM");
            Query.Criteria.AddCondition("fullname", ConditionOperator.NotEqual, "INTEGRATION");
            //Query.Criteria.AddCondition("hil_archivelog", ConditionOperator.NotEqual, true);
            Query.AddOrder("createdon", OrderType.Ascending);
            Query.PageInfo = new PagingInfo();
            Query.PageInfo.Count = 5000;
            Query.PageInfo.PageNumber = 1;
            Query.PageInfo.ReturnTotalRecordCount = true;

            try
            {
                EntityCollection ec = _service.RetrieveMultiple(Query);
                do
                {
                    if (ec.Entities.Count > 0)
                    {
                        string accName = "d365storagesa";
                        string accKey = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
                        //bool _retValue = false;
                        // Implement the accout, set true for https for SSL.  
                        StorageCredentials creds = new StorageCredentials(accName, accKey);
                        CloudStorageAccount strAcc = new CloudStorageAccount(creds, true);
                        // Create the blob client.  
                        CloudBlobClient blobClient = strAcc.CreateCloudBlobClient();
                        // Retrieve a reference to a container.   
                        CloudBlobContainer container = blobClient.GetContainerReference("d365auditlogarchive");
                        // Create the container if it doesn't already exist.  
                        container.CreateIfNotExistsAsync();

                        foreach (Entity ent in ec.Entities)
                        {
                            //CheckIfAzureBlobExist(ent.Id.ToString(), container) || 
                            if (ent.GetAttributeValue<string>("internalemailaddress") == "havellconnect@havells.com")
                            {
                                Console.WriteLine("User # " + i.ToString() + " " + ent.GetAttributeValue<string>("internalemailaddress"));
                                //try
                                //{
                                //    entTemp = new Entity("systemuser", ent.Id);
                                //    entTemp["hil_archivelog"] = true;
                                //    _service.Update(entTemp);
                                //}
                                //catch  { }
                                i += 1;
                                continue;
                            }
                            Console.WriteLine("User # " + i.ToString() + " " + ent.GetAttributeValue<string>("internalemailaddress") + " Start..." + DateTime.Now.ToString());
                            RetrieveRecordChangeHistoryRequest changeRequest = new RetrieveRecordChangeHistoryRequest();
                            changeRequest.Target = new EntityReference(entityLogicalName, ent.Id);
                            RetrieveRecordChangeHistoryResponse changeResponse =
                            (RetrieveRecordChangeHistoryResponse)_service.Execute(changeRequest);

                            AuditDetailCollection auditDetailCollection = changeResponse.AuditDetailCollection;
                            StringBuilder strContent = new StringBuilder();
                            strContent.AppendLine("auditid,createdon,action,actionname,objectid,objecttypecode,objecttypecodename,operation,operationname,userid,useridname,fieldname,oldvalue,oldvaluename,oldvaluetype,newvalue,newvaluename,newvaluetype");
                            string strLine;
                            string strAttributes;
                            string dateString;
                            foreach (var attrAuditDetail in auditDetailCollection.AuditDetails)
                            {
                                var auditRecord = attrAuditDetail.AuditRecord;
                                strLine = string.Empty;
                                dateString = string.Empty;
                                if (auditRecord.Attributes.Contains("auditid"))
                                {
                                    strLine = auditRecord.Attributes["auditid"].ToString();
                                }
                                else
                                {
                                    strLine = null;
                                }
                                if (auditRecord.Attributes.Contains("createdon"))
                                {
                                    DateTime dt = ((DateTime)auditRecord.Attributes["createdon"]).AddMinutes(330);
                                    strLine += "," + dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0') + dt.Hour.ToString().PadLeft(2, '0') + dt.Minute.ToString().PadLeft(2, '0') + dt.Second.ToString().PadLeft(2, '0');
                                }
                                else
                                {
                                    strLine += "," + string.Empty;
                                }
                                if (auditRecord.Attributes.Contains("action"))
                                {
                                    strLine += "," + ((OptionSetValue)auditRecord.Attributes["action"]).Value.ToString();
                                    strLine += "," + auditRecord.FormattedValues["action"];
                                }
                                else
                                {

                                    strLine += "," + string.Empty + "," + string.Empty;
                                }
                                if (auditRecord.Attributes.Contains("objectid"))
                                {
                                    strLine += "," + ((EntityReference)auditRecord.Attributes["objectid"]).Id.ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("objecttypecode"))
                                {
                                    strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                    strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("operation"))
                                {
                                    strLine += "," + ((OptionSetValue)auditRecord.Attributes["operation"]).Value.ToString();
                                    strLine += "," + auditRecord.FormattedValues["operation"];
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("userid"))
                                {
                                    EntityReference er = auditRecord.GetAttributeValue<EntityReference>("userid");
                                    strLine += "," + er.Id.ToString();
                                    strLine += "," + er.Name.ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }
                                strLine += ",";
                                strAttributes = string.Empty;
                                if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.AttributeAuditDetail")
                                {
                                    Dictionary<string, Object> oldValues = new Dictionary<string, object>();
                                    Dictionary<string, Object> newValues = new Dictionary<string, object>();
                                    object _OldValue;
                                    //object _NewValue;
                                    var newValueEntity = ((AttributeAuditDetail)attrAuditDetail).NewValue;
                                    if (newValueEntity != null)
                                    {
                                        newValues = newValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                    }

                                    var oldValueEntity = ((AttributeAuditDetail)attrAuditDetail).OldValue;
                                    if (oldValueEntity != null)
                                    {
                                        oldValues = oldValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                    }
                                    if (newValues.Count > 0)
                                    {
                                        foreach (KeyValuePair<string, object> de in newValues)
                                        {
                                            strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                            //strAttributes = de.Key.ToString() + ",";
                                            //Console.WriteLine("Attribute Name :" + strAttributes);
                                            if (oldValues.ContainsKey(de.Key))
                                            {
                                                _OldValue = oldValues[de.Key.ToString()];
                                                if (_OldValue.GetType() == typeof(Decimal))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Double))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(System.DateTime))
                                                {
                                                    DateTime dtValue = ((DateTime)_OldValue).AddMinutes(330);
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                {
                                                    EntityReference entRef = (EntityReference)_OldValue;
                                                    if (entRef != null)
                                                    {
                                                        strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                    }
                                                    else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                }
                                                else if (_OldValue.GetType() == typeof(System.Boolean))
                                                {
                                                    string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                    strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(System.Int32))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                {
                                                    Money money = (Money)_OldValue;
                                                    strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                {
                                                    OptionSetValue optValue = (OptionSetValue)_OldValue;
                                                    EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                    EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                    strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(System.String))
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                }

                                                oldValues.Remove(de.Key.ToString());
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                            }

                                            if (de.Value.GetType() == typeof(Decimal))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Double))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.DateTime))
                                            {
                                                DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                            {
                                                EntityReference entRef = (EntityReference)de.Value;
                                                if (entRef != null)
                                                {
                                                    strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                }
                                                else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                            }
                                            else if (de.Value.GetType() == typeof(System.Boolean))
                                            {
                                                string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.Int32))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                            {
                                                Money money = (Money)de.Value;
                                                strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                            {
                                                OptionSetValue optValue = (OptionSetValue)de.Value;
                                                //string name = string.Empty;
                                                //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.String))
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }

                                            if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                            strContent.AppendLine(strLine + strAttributes);
                                        }
                                    }
                                    if (oldValues.Count > 0)
                                    {
                                        foreach (KeyValuePair<string, object> de in oldValues)
                                        {
                                            //strAttributes = de.Key.ToString() + ",";
                                            strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                            if (de.Value.GetType() == typeof(Decimal))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Double))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.DateTime))
                                            {
                                                DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                            {
                                                EntityReference entRef = (EntityReference)de.Value;
                                                if (entRef != null)
                                                {
                                                    strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                }
                                                else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                            }
                                            else if (de.Value.GetType() == typeof(System.Boolean))
                                            {
                                                string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.Int32))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                            {
                                                Money money = (Money)de.Value;
                                                strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                            {
                                                OptionSetValue optValue = (OptionSetValue)de.Value;
                                                //string name = string.Empty;
                                                //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.String))
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                            if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                            strContent.AppendLine(strLine + strAttributes);
                                        }
                                    }
                                }
                                else if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.UserAccessAuditDetail")
                                {
                                    var userAccessAuditDetail = ((UserAccessAuditDetail)attrAuditDetail);
                                    strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + userAccessAuditDetail.AccessTime.AddMinutes(330).ToString() + "," + string.Empty + "," + string.Empty;
                                    strContent.AppendLine(strLine + strAttributes);
                                }
                                else
                                {
                                    strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty;
                                    strContent.AppendLine(strLine + strAttributes);
                                }
                            }
                            //UploadAzureBlob(ent.Id.ToString(), strContent.ToString());
                            UploadAzureBlob(ent.Id.ToString(), strContent.ToString(), container);

                            var DeleteAudit = new DeleteRecordChangeHistoryRequest();
                            DeleteAudit.Target = (EntityReference)ent.ToEntityReference();
                            try
                            {
                                _service.Execute(DeleteAudit);
                                entTemp = new Entity("systemuser", ent.Id);
                                entTemp["hil_archivelog"] = true;
                                _service.Update(entTemp);
                            }
                            catch { }
                            Console.WriteLine("User # " + i.ToString() + " End... " + DateTime.Now.ToString());
                            i += 1;
                        }

                        #region Delete Audit Log

                        #endregion
                    }
                    Query.PageInfo.PageNumber += 1;
                    Query.PageInfo.PagingCookie = ec.PagingCookie;
                    ec = _service.RetrieveMultiple(Query);
                } while (ec.MoreRecords);
                Console.WriteLine("DONE !!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        static void D365AuditLogMigration_DeactiveSystemUser()
        {
            int i = 1;
            Console.WriteLine("Fetching Entity Attribute Metadata..." + DateTime.Now.ToString());
            List<EntityAttribute> lstEntityAttribute = GetEntityAttributes("systemuser");
            Console.WriteLine("Entity Attribute Metadata successfully fetched..." + DateTime.Now.ToString());
            Entity entTemp = null;
            string entityLogicalName = "systemuser";
            string[] primaryKeySchemaName = { "systemuserid", "internalemailaddress" };


            QueryExpression Query = new QueryExpression(entityLogicalName);
            Query.ColumnSet = new ColumnSet(primaryKeySchemaName);
            //Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, "2018-12-31");
            Query.Criteria.AddCondition("isdisabled", ConditionOperator.Equal, true);
            Query.Criteria.AddCondition("fullname", ConditionOperator.NotEqual, "SYSTEM");
            Query.Criteria.AddCondition("fullname", ConditionOperator.NotEqual, "INTEGRATION");
            Query.Criteria.AddCondition("hil_archivelog", ConditionOperator.NotEqual, true);

            Query.AddOrder("createdon", OrderType.Ascending);
            Query.PageInfo = new PagingInfo();
            Query.PageInfo.Count = 5000;
            Query.PageInfo.PageNumber = 1;
            Query.PageInfo.ReturnTotalRecordCount = true;

            try
            {
                EntityCollection ec = _service.RetrieveMultiple(Query);
                do
                {
                    if (ec.Entities.Count > 0)
                    {
                        string accName = "d365storagesa";
                        string accKey = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
                        //bool _retValue = false;
                        // Implement the accout, set true for https for SSL.  
                        StorageCredentials creds = new StorageCredentials(accName, accKey);
                        CloudStorageAccount strAcc = new CloudStorageAccount(creds, true);
                        // Create the blob client.  
                        CloudBlobClient blobClient = strAcc.CreateCloudBlobClient();
                        // Retrieve a reference to a container.   
                        CloudBlobContainer container = blobClient.GetContainerReference("d365auditlogarchive");
                        // Create the container if it doesn't already exist.  
                        container.CreateIfNotExistsAsync();

                        foreach (Entity ent in ec.Entities)
                        {
                            if (CheckIfAzureBlobExist(ent.Id.ToString(), container) || ent.GetAttributeValue<string>("internalemailaddress") == "havellconnect@havells.com")
                            {
                                Console.WriteLine("User # " + i.ToString() + " " + ent.GetAttributeValue<string>("internalemailaddress"));
                                try
                                {
                                    entTemp = new Entity("systemuser", ent.Id);
                                    entTemp["hil_archivelog"] = true;
                                    _service.Update(entTemp);
                                }
                                catch { }
                                i += 1;
                                continue;
                            }
                            Console.WriteLine("User # " + i.ToString() + " " + ent.GetAttributeValue<string>("internalemailaddress") + " Start..." + DateTime.Now.ToString());
                            RetrieveRecordChangeHistoryRequest changeRequest = new RetrieveRecordChangeHistoryRequest();
                            changeRequest.Target = new EntityReference(entityLogicalName, ent.Id);
                            RetrieveRecordChangeHistoryResponse changeResponse =
                            (RetrieveRecordChangeHistoryResponse)_service.Execute(changeRequest);

                            AuditDetailCollection auditDetailCollection = changeResponse.AuditDetailCollection;
                            StringBuilder strContent = new StringBuilder();
                            strContent.AppendLine("auditid,createdon,action,actionname,objectid,objecttypecode,objecttypecodename,operation,operationname,userid,useridname,fieldname,oldvalue,oldvaluename,oldvaluetype,newvalue,newvaluename,newvaluetype");
                            string strLine;
                            string strAttributes;
                            string dateString;
                            foreach (var attrAuditDetail in auditDetailCollection.AuditDetails)
                            {
                                var auditRecord = attrAuditDetail.AuditRecord;
                                strLine = string.Empty;
                                dateString = string.Empty;
                                if (auditRecord.Attributes.Contains("auditid"))
                                {
                                    strLine = auditRecord.Attributes["auditid"].ToString();
                                }
                                else
                                {
                                    strLine = null;
                                }
                                if (auditRecord.Attributes.Contains("createdon"))
                                {
                                    DateTime dt = ((DateTime)auditRecord.Attributes["createdon"]).AddMinutes(330);
                                    strLine += "," + dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0') + dt.Hour.ToString().PadLeft(2, '0') + dt.Minute.ToString().PadLeft(2, '0') + dt.Second.ToString().PadLeft(2, '0');
                                }
                                else
                                {
                                    strLine += "," + string.Empty;
                                }
                                if (auditRecord.Attributes.Contains("action"))
                                {
                                    strLine += "," + ((OptionSetValue)auditRecord.Attributes["action"]).Value.ToString();
                                    strLine += "," + auditRecord.FormattedValues["action"];
                                }
                                else
                                {

                                    strLine += "," + string.Empty + "," + string.Empty;
                                }
                                if (auditRecord.Attributes.Contains("objectid"))
                                {
                                    strLine += "," + ((EntityReference)auditRecord.Attributes["objectid"]).Id.ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("objecttypecode"))
                                {
                                    strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                    strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("operation"))
                                {
                                    strLine += "," + ((OptionSetValue)auditRecord.Attributes["operation"]).Value.ToString();
                                    strLine += "," + auditRecord.FormattedValues["operation"];
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("userid"))
                                {
                                    EntityReference er = auditRecord.GetAttributeValue<EntityReference>("userid");
                                    strLine += "," + er.Id.ToString();
                                    strLine += "," + er.Name.ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }
                                strLine += ",";
                                strAttributes = string.Empty;
                                if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.AttributeAuditDetail")
                                {
                                    Dictionary<string, Object> oldValues = new Dictionary<string, object>();
                                    Dictionary<string, Object> newValues = new Dictionary<string, object>();
                                    object _OldValue;
                                    //object _NewValue;
                                    var newValueEntity = ((AttributeAuditDetail)attrAuditDetail).NewValue;
                                    if (newValueEntity != null)
                                    {
                                        newValues = newValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                    }

                                    var oldValueEntity = ((AttributeAuditDetail)attrAuditDetail).OldValue;
                                    if (oldValueEntity != null)
                                    {
                                        oldValues = oldValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                    }
                                    if (newValues.Count > 0)
                                    {
                                        foreach (KeyValuePair<string, object> de in newValues)
                                        {
                                            strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                            //strAttributes = de.Key.ToString() + ",";
                                            //Console.WriteLine("Attribute Name :" + strAttributes);
                                            if (oldValues.ContainsKey(de.Key))
                                            {
                                                _OldValue = oldValues[de.Key.ToString()];
                                                if (_OldValue.GetType() == typeof(Decimal))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Double))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(System.DateTime))
                                                {
                                                    DateTime dtValue = ((DateTime)_OldValue).AddMinutes(330);
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                {
                                                    EntityReference entRef = (EntityReference)_OldValue;
                                                    if (entRef != null)
                                                    {
                                                        strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                    }
                                                    else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                }
                                                else if (_OldValue.GetType() == typeof(System.Boolean))
                                                {
                                                    string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                    strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(System.Int32))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                {
                                                    Money money = (Money)_OldValue;
                                                    strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                {
                                                    OptionSetValue optValue = (OptionSetValue)_OldValue;
                                                    EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                    EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                    strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(System.String))
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                }

                                                oldValues.Remove(de.Key.ToString());
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                            }

                                            if (de.Value.GetType() == typeof(Decimal))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Double))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.DateTime))
                                            {
                                                DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                            {
                                                EntityReference entRef = (EntityReference)de.Value;
                                                if (entRef != null)
                                                {
                                                    strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                }
                                                else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                            }
                                            else if (de.Value.GetType() == typeof(System.Boolean))
                                            {
                                                string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.Int32))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                            {
                                                Money money = (Money)de.Value;
                                                strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                            {
                                                OptionSetValue optValue = (OptionSetValue)de.Value;
                                                //string name = string.Empty;
                                                //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.String))
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }

                                            if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                            strContent.AppendLine(strLine + strAttributes);
                                        }
                                    }
                                    if (oldValues.Count > 0)
                                    {
                                        foreach (KeyValuePair<string, object> de in oldValues)
                                        {
                                            //strAttributes = de.Key.ToString() + ",";
                                            strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                            if (de.Value.GetType() == typeof(Decimal))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Double))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.DateTime))
                                            {
                                                DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                            {
                                                EntityReference entRef = (EntityReference)de.Value;
                                                if (entRef != null)
                                                {
                                                    strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                }
                                                else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                            }
                                            else if (de.Value.GetType() == typeof(System.Boolean))
                                            {
                                                string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.Int32))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                            {
                                                Money money = (Money)de.Value;
                                                strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                            {
                                                OptionSetValue optValue = (OptionSetValue)de.Value;
                                                //string name = string.Empty;
                                                //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.String))
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                            if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                            strContent.AppendLine(strLine + strAttributes);
                                        }
                                    }
                                }
                                else if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.UserAccessAuditDetail")
                                {
                                    var userAccessAuditDetail = ((UserAccessAuditDetail)attrAuditDetail);
                                    strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + userAccessAuditDetail.AccessTime.AddMinutes(330).ToString() + "," + string.Empty + "," + string.Empty;
                                    strContent.AppendLine(strLine + strAttributes);
                                }
                                else
                                {
                                    strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty;
                                    strContent.AppendLine(strLine + strAttributes);
                                }
                            }
                            //UploadAzureBlob(ent.Id.ToString(), strContent.ToString());
                            UploadAzureBlob(ent.Id.ToString(), strContent.ToString(), container);

                            var DeleteAudit = new DeleteRecordChangeHistoryRequest();
                            DeleteAudit.Target = (EntityReference)ent.ToEntityReference();
                            try
                            {
                                _service.Execute(DeleteAudit);
                                entTemp = new Entity("systemuser", ent.Id);
                                entTemp["hil_archivelog"] = true;
                                _service.Update(entTemp);
                            }
                            catch { }
                            Console.WriteLine("User # " + i.ToString() + " End... " + DateTime.Now.ToString());
                            i += 1;
                        }

                        #region Delete Audit Log

                        #endregion
                    }
                    Query.PageInfo.PageNumber += 1;
                    Query.PageInfo.PagingCookie = ec.PagingCookie;
                    ec = _service.RetrieveMultiple(Query);
                } while (ec.MoreRecords);
                Console.WriteLine("DONE !!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        static void D365AuditLogMigration_Customer()
        {
            int i = 1;
            Console.WriteLine("Fetching Entity Attribute Metadata..." + DateTime.Now.ToString());
            List<EntityAttribute> lstEntityAttribute = GetEntityAttributes("contact");
            Console.WriteLine("Entity Attribute Metadata successfully fetched..." + DateTime.Now.ToString());
            Entity entTemp = null;
            string entityLogicalName = "contact";
            string[] primaryKeySchemaName = { "contactid", "mobilephone", "createdon" };

            QueryExpression Query = new QueryExpression(entityLogicalName);
            Query.ColumnSet = new ColumnSet(primaryKeySchemaName);
            Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, "2019-05-01");
            Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrBefore, "2019-06-01");
            Query.Criteria.AddCondition("hil_archivelog", ConditionOperator.NotEqual, true);
            Query.AddOrder("createdon", OrderType.Ascending);
            Query.PageInfo = new PagingInfo();
            Query.PageInfo.Count = 5000;
            Query.PageInfo.PageNumber = 1;
            Query.PageInfo.ReturnTotalRecordCount = true;

            try
            {
                EntityCollection ec = _service.RetrieveMultiple(Query);
                do
                {
                    if (ec.Entities.Count > 0)
                    {
                        string accName = "d365storagesa";
                        string accKey = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
                        //bool _retValue = false;
                        // Implement the accout, set true for https for SSL.  
                        StorageCredentials creds = new StorageCredentials(accName, accKey);
                        CloudStorageAccount strAcc = new CloudStorageAccount(creds, true);
                        // Create the blob client.  
                        CloudBlobClient blobClient = strAcc.CreateCloudBlobClient();
                        // Retrieve a reference to a container.   
                        CloudBlobContainer container = blobClient.GetContainerReference("d365auditlogarchive");
                        // Create the container if it doesn't already exist.  
                        container.CreateIfNotExistsAsync();

                        foreach (Entity ent in ec.Entities)
                        {
                            if (CheckIfAzureBlobExist(ent.Id.ToString(), container))
                            {
                                Console.WriteLine("User # " + i.ToString() + " " + ent.GetAttributeValue<string>("mobilephone"));
                                try
                                {
                                    entTemp = new Entity("contact", ent.Id);
                                    entTemp["hil_archivelog"] = true;
                                    _service.Update(entTemp);
                                }
                                catch { }
                                i += 1;
                                continue;
                            }
                            Console.WriteLine("User # " + i.ToString() + " " + ent.GetAttributeValue<string>("mobilephone") + " Start..." + ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString());
                            RetrieveRecordChangeHistoryRequest changeRequest = new RetrieveRecordChangeHistoryRequest();
                            changeRequest.Target = new EntityReference(entityLogicalName, ent.Id);
                            RetrieveRecordChangeHistoryResponse changeResponse =
                            (RetrieveRecordChangeHistoryResponse)_service.Execute(changeRequest);

                            AuditDetailCollection auditDetailCollection = changeResponse.AuditDetailCollection;
                            StringBuilder strContent = new StringBuilder();
                            strContent.AppendLine("auditid,createdon,action,actionname,objectid,objecttypecode,objecttypecodename,operation,operationname,userid,useridname,fieldname,oldvalue,oldvaluename,oldvaluetype,newvalue,newvaluename,newvaluetype");
                            string strLine;
                            string strAttributes;
                            string dateString;
                            foreach (var attrAuditDetail in auditDetailCollection.AuditDetails)
                            {
                                var auditRecord = attrAuditDetail.AuditRecord;
                                strLine = string.Empty;
                                dateString = string.Empty;
                                if (auditRecord.Attributes.Contains("auditid"))
                                {
                                    strLine = auditRecord.Attributes["auditid"].ToString();
                                }
                                else
                                {
                                    strLine = null;
                                }
                                if (auditRecord.Attributes.Contains("createdon"))
                                {
                                    DateTime dt = ((DateTime)auditRecord.Attributes["createdon"]).AddMinutes(330);
                                    strLine += "," + dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0') + dt.Hour.ToString().PadLeft(2, '0') + dt.Minute.ToString().PadLeft(2, '0') + dt.Second.ToString().PadLeft(2, '0');
                                }
                                else
                                {
                                    strLine += "," + string.Empty;
                                }
                                if (auditRecord.Attributes.Contains("action"))
                                {
                                    strLine += "," + ((OptionSetValue)auditRecord.Attributes["action"]).Value.ToString();
                                    strLine += "," + auditRecord.FormattedValues["action"];
                                }
                                else
                                {

                                    strLine += "," + string.Empty + "," + string.Empty;
                                }
                                if (auditRecord.Attributes.Contains("objectid"))
                                {
                                    strLine += "," + ((EntityReference)auditRecord.Attributes["objectid"]).Id.ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("objecttypecode"))
                                {
                                    strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                    strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("operation"))
                                {
                                    strLine += "," + ((OptionSetValue)auditRecord.Attributes["operation"]).Value.ToString();
                                    strLine += "," + auditRecord.FormattedValues["operation"];
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("userid"))
                                {
                                    EntityReference er = auditRecord.GetAttributeValue<EntityReference>("userid");
                                    strLine += "," + er.Id.ToString();
                                    strLine += "," + er.Name.ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }
                                strLine += ",";
                                strAttributes = string.Empty;
                                if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.AttributeAuditDetail")
                                {
                                    Dictionary<string, Object> oldValues = new Dictionary<string, object>();
                                    Dictionary<string, Object> newValues = new Dictionary<string, object>();
                                    object _OldValue;
                                    //object _NewValue;
                                    var newValueEntity = ((AttributeAuditDetail)attrAuditDetail).NewValue;
                                    if (newValueEntity != null)
                                    {
                                        newValues = newValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                    }

                                    var oldValueEntity = ((AttributeAuditDetail)attrAuditDetail).OldValue;
                                    if (oldValueEntity != null)
                                    {
                                        oldValues = oldValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                    }
                                    if (newValues.Count > 0)
                                    {
                                        foreach (KeyValuePair<string, object> de in newValues)
                                        {
                                            strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                            //strAttributes = de.Key.ToString() + ",";
                                            //Console.WriteLine("Attribute Name :" + strAttributes);
                                            if (oldValues.ContainsKey(de.Key))
                                            {
                                                _OldValue = oldValues[de.Key.ToString()];
                                                if (_OldValue.GetType() == typeof(Decimal))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Double))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(System.DateTime))
                                                {
                                                    DateTime dtValue = ((DateTime)_OldValue).AddMinutes(330);
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                {
                                                    EntityReference entRef = (EntityReference)_OldValue;
                                                    if (entRef != null)
                                                    {
                                                        strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                    }
                                                    else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                }
                                                else if (_OldValue.GetType() == typeof(System.Boolean))
                                                {
                                                    string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                    strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(System.Int32))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                {
                                                    Money money = (Money)_OldValue;
                                                    strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                {
                                                    OptionSetValue optValue = (OptionSetValue)_OldValue;
                                                    EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                    EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                    if (optionSet != null)
                                                    {
                                                        strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                    }
                                                    else {
                                                        strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                    }
                                                }
                                                else if (_OldValue.GetType() == typeof(System.String))
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                }

                                                oldValues.Remove(de.Key.ToString());
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                            }

                                            if (de.Value.GetType() == typeof(Decimal))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Double))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.DateTime))
                                            {
                                                DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                            {
                                                EntityReference entRef = (EntityReference)de.Value;
                                                if (entRef != null)
                                                {
                                                    strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                }
                                                else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                            }
                                            else if (de.Value.GetType() == typeof(System.Boolean))
                                            {
                                                string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.Int32))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                            {
                                                Money money = (Money)de.Value;
                                                strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                            {
                                                OptionSetValue optValue = (OptionSetValue)de.Value;
                                                //string name = string.Empty;
                                                //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                if (optionSet != null)
                                                {
                                                    strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                }
                                                else {
                                                    strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                }
                                            }
                                            else if (de.Value.GetType() == typeof(System.String))
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }

                                            if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                            strContent.AppendLine(strLine + strAttributes);
                                        }
                                    }
                                    if (oldValues.Count > 0)
                                    {
                                        foreach (KeyValuePair<string, object> de in oldValues)
                                        {
                                            //strAttributes = de.Key.ToString() + ",";
                                            strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                            if (de.Value.GetType() == typeof(Decimal))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Double))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.DateTime))
                                            {
                                                DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                            {
                                                EntityReference entRef = (EntityReference)de.Value;
                                                if (entRef != null)
                                                {
                                                    strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                }
                                                else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                            }
                                            else if (de.Value.GetType() == typeof(System.Boolean))
                                            {
                                                string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.Int32))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                            {
                                                Money money = (Money)de.Value;
                                                strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                            {
                                                OptionSetValue optValue = (OptionSetValue)de.Value;
                                                //string name = string.Empty;
                                                //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                if (optionSet != null) {
                                                    strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                }
                                                else {
                                                    strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                }
                                            }
                                            else if (de.Value.GetType() == typeof(System.String))
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                            if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                            strContent.AppendLine(strLine + strAttributes);
                                        }
                                    }
                                }
                                else if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.UserAccessAuditDetail")
                                {
                                    var userAccessAuditDetail = ((UserAccessAuditDetail)attrAuditDetail);
                                    strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + userAccessAuditDetail.AccessTime.AddMinutes(330).ToString() + "," + string.Empty + "," + string.Empty;
                                    strContent.AppendLine(strLine + strAttributes);
                                }
                                else
                                {
                                    strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty;
                                    strContent.AppendLine(strLine + strAttributes);
                                }
                            }
                            //UploadAzureBlob(ent.Id.ToString(), strContent.ToString());
                            UploadAzureBlob(ent.Id.ToString(), strContent.ToString(), container);

                            var DeleteAudit = new DeleteRecordChangeHistoryRequest();
                            DeleteAudit.Target = (EntityReference)ent.ToEntityReference();
                            try
                            {
                                _service.Execute(DeleteAudit);
                                entTemp = new Entity("contact", ent.Id);
                                entTemp["hil_archivelog"] = true;
                                _service.Update(entTemp);
                            }
                            catch { }
                            Console.WriteLine("User # " + i.ToString() + " End... " + DateTime.Now.ToString());
                            i += 1;
                        }

                        #region Delete Audit Log

                        #endregion
                    }
                    Query.PageInfo.PageNumber += 1;
                    Query.PageInfo.PagingCookie = ec.PagingCookie;
                    ec = _service.RetrieveMultiple(Query);
                } while (ec.MoreRecords);
                Console.WriteLine("DONE !!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        static void D365AuditLogMigration_Delta(string _entityName, string _fromDate, string _toDate)
        {
            int i = 1;
            string _primaryfield = string.Empty;
            string _primaryKey = string.Empty;
            GetPrimaryIdFieldName(_entityName, out _primaryKey, out _primaryfield);
            List<EntityAttribute> lstEntityAttribute = GetEntityAttributes(_entityName);
            Console.WriteLine("Entity Attribute Metadata successfully fetched..." + DateTime.Now.ToString());
            Entity entTemp = null;
            string entityLogicalName = _entityName;
            string[] primaryKeySchemaName = { _primaryKey, _primaryfield, "createdon" };

            QueryExpression query = new QueryExpression(entityLogicalName);
            query.ColumnSet = new ColumnSet(primaryKeySchemaName);
            //query.Criteria.AddCondition("hil_isauditlogmigrated", ConditionOperator.NotEqual, true);
            query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, _fromDate);
            query.Criteria.AddCondition("createdon", ConditionOperator.OnOrBefore, _toDate);
            query.AddOrder("createdon", OrderType.Ascending);
            query.NoLock = true;
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;

            try
            {
                EntityCollection ec = _service.RetrieveMultiple(query);
                do
                {
                    if (ec.Entities.Count > 0)
                    {
                        string accName = "d365storagesa";
                        string accKey = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
                        //bool _retValue = false;
                        string _recID = string.Empty;

                        // Implement the accout, set true for https for SSL.  
                        StorageCredentials creds = new StorageCredentials(accName, accKey);
                        CloudStorageAccount strAcc = new CloudStorageAccount(creds, true);
                        // Create the blob client.  
                        CloudBlobClient blobClient = strAcc.CreateCloudBlobClient();
                        // Retrieve a reference to a container.   
                        CloudBlobContainer container = blobClient.GetContainerReference("d365auditlogarchive");
                        // Create the container if it doesn't already exist.  
                        container.CreateIfNotExistsAsync();

                        foreach (Entity ent in ec.Entities)
                        {
                            if (ent.Attributes.Contains(_primaryfield))
                            {
                                _recID = ent.GetAttributeValue<string>(_primaryfield);
                            }
                            else { _recID = ""; }

                            StringBuilder strContent = new StringBuilder();
                            if (!CheckIfAzureBlobExist(ent.Id.ToString(), container))
                            {
                                strContent.AppendLine("auditid,createdon,action,actionname,objectid,objecttypecode,objecttypecodename,operation,operationname,userid,useridname,fieldname,oldvalue,oldvaluename,oldvaluetype,newvalue,newvaluename,newvaluetype");
                            }
                            Console.WriteLine("Record # " + i.ToString() + " " + _recID + " Start..." + ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString());
                            RetrieveRecordChangeHistoryRequest changeRequest = new RetrieveRecordChangeHistoryRequest();
                            changeRequest.Target = new EntityReference(entityLogicalName, ent.Id);
                            RetrieveRecordChangeHistoryResponse changeResponse=null;

                            Repeat:
                            try
                            {
                                changeResponse =(RetrieveRecordChangeHistoryResponse)_service.Execute(changeRequest);
                            }
                            catch
                            {
                                goto Repeat;
                            }

                            AuditDetailCollection auditDetailCollection = changeResponse.AuditDetailCollection;
                            string strLine;
                            string strAttributes;
                            string dateString;
                            bool _logExists = false;
                            foreach (var attrAuditDetail in auditDetailCollection.AuditDetails)
                            {
                                var auditRecord = attrAuditDetail.AuditRecord;
                                strLine = string.Empty;
                                dateString = string.Empty;
                                string _action = auditRecord.FormattedValues["action"];

                                if (auditRecord.Attributes.Contains("objectid") && (_action == "Create" || _action == "Update" || _action == "Delete"))
                                {
                                    if (auditRecord.Attributes.Contains("auditid"))
                                    {
                                        strLine = auditRecord.Attributes["auditid"].ToString();
                                    }
                                    else
                                    {
                                        strLine = null;
                                    }
                                    if (auditRecord.Attributes.Contains("createdon"))
                                    {
                                        DateTime dt = ((DateTime)auditRecord.Attributes["createdon"]).AddMinutes(330);
                                        strLine += "," + dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0') + dt.Hour.ToString().PadLeft(2, '0') + dt.Minute.ToString().PadLeft(2, '0') + dt.Second.ToString().PadLeft(2, '0');
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty;
                                    }
                                    if (auditRecord.Attributes.Contains("action"))
                                    {
                                        strLine += "," + ((OptionSetValue)auditRecord.Attributes["action"]).Value.ToString();
                                        strLine += "," + auditRecord.FormattedValues["action"];
                                    }
                                    else
                                    {

                                        strLine += "," + string.Empty + "," + string.Empty;
                                    }
                                    if (auditRecord.Attributes.Contains("objectid"))
                                    {
                                        strLine += "," + ((EntityReference)auditRecord.Attributes["objectid"]).Id.ToString();
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty;
                                    }

                                    if (auditRecord.Attributes.Contains("objecttypecode"))
                                    {
                                        strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                        strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty + "," + string.Empty;
                                    }

                                    if (auditRecord.Attributes.Contains("operation"))
                                    {
                                        strLine += "," + ((OptionSetValue)auditRecord.Attributes["operation"]).Value.ToString();
                                        strLine += "," + auditRecord.FormattedValues["operation"];
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty + "," + string.Empty;
                                    }

                                    if (auditRecord.Attributes.Contains("userid"))
                                    {
                                        EntityReference er = auditRecord.GetAttributeValue<EntityReference>("userid");
                                        strLine += "," + er.Id.ToString();
                                        strLine += "," + er.Name.ToString();
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty + "," + string.Empty;
                                    }
                                    strLine += ",";
                                    strAttributes = string.Empty;
                                    if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.AttributeAuditDetail")
                                    {
                                        Dictionary<string, Object> oldValues = new Dictionary<string, object>();
                                        Dictionary<string, Object> newValues = new Dictionary<string, object>();
                                        object _OldValue;
                                        //object _NewValue;
                                        var newValueEntity = ((AttributeAuditDetail)attrAuditDetail).NewValue;
                                        if (newValueEntity != null)
                                        {
                                            newValues = newValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                        }

                                        var oldValueEntity = ((AttributeAuditDetail)attrAuditDetail).OldValue;
                                        if (oldValueEntity != null)
                                        {
                                            oldValues = oldValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                        }
                                        if (newValues.Count > 0)
                                        {
                                            foreach (KeyValuePair<string, object> de in newValues)
                                            {
                                                strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                                //strAttributes = de.Key.ToString() + ",";
                                                //Console.WriteLine("Attribute Name :" + strAttributes);
                                                if (oldValues.ContainsKey(de.Key))
                                                {
                                                    _OldValue = oldValues[de.Key.ToString()];
                                                    if (_OldValue.GetType() == typeof(Decimal))
                                                    {
                                                        strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(Double))
                                                    {
                                                        strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(System.DateTime))
                                                    {
                                                        DateTime dtValue = ((DateTime)_OldValue).AddMinutes(330);
                                                        strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                    {
                                                        EntityReference entRef = (EntityReference)_OldValue;
                                                        if (entRef != null)
                                                        {
                                                            strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                        }
                                                        else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                    }
                                                    else if (_OldValue.GetType() == typeof(System.Boolean))
                                                    {
                                                        string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(_OldValue));
                                                        //string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                        strAttributes += _OldValue.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(System.Int32))
                                                    {
                                                        strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                    {
                                                        Money money = (Money)_OldValue;
                                                        strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                    {
                                                        OptionSetValue optValue = (OptionSetValue)_OldValue;
                                                        EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                        EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                        if (optionSet != null)
                                                        {
                                                            strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                        }
                                                        else
                                                        {
                                                            strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                        }
                                                    }
                                                    else if (_OldValue.GetType() == typeof(System.String))
                                                    {
                                                        strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                    }
                                                    else
                                                    {
                                                        strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                    }

                                                    oldValues.Remove(de.Key.ToString());
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                                }

                                                if (de.Value.GetType() == typeof(Decimal))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Double))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(System.DateTime))
                                                {
                                                    DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                    strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                {
                                                    EntityReference entRef = (EntityReference)de.Value;
                                                    if (entRef != null)
                                                    {
                                                        strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                    }
                                                    else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                }
                                                else if (de.Value.GetType() == typeof(System.Boolean))
                                                {
                                                    string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                    strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(System.Int32))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                {
                                                    Money money = (Money)de.Value;
                                                    strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                {
                                                    OptionSetValue optValue = (OptionSetValue)de.Value;
                                                    //string name = string.Empty;
                                                    //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                    EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                    EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                    if (optionSet != null)
                                                    {
                                                        strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                    }
                                                    else
                                                    {
                                                        strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                    }
                                                }
                                                else if (de.Value.GetType() == typeof(System.String))
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                                }

                                                if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                                strContent.AppendLine(strLine + strAttributes);
                                            }
                                        }
                                        if (oldValues.Count > 0)
                                        {
                                            foreach (KeyValuePair<string, object> de in oldValues)
                                            {
                                                //strAttributes = de.Key.ToString() + ",";
                                                strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                                if (de.Value.GetType() == typeof(Decimal))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Double))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(System.DateTime))
                                                {
                                                    DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                    strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                {
                                                    EntityReference entRef = (EntityReference)de.Value;
                                                    if (entRef != null)
                                                    {
                                                        strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                    }
                                                    else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                }
                                                else if (de.Value.GetType() == typeof(System.Boolean))
                                                {
                                                    string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                    strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(System.Int32))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                {
                                                    Money money = (Money)de.Value;
                                                    strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                {
                                                    OptionSetValue optValue = (OptionSetValue)de.Value;
                                                    //string name = string.Empty;
                                                    //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                    EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                    EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                    if (optionSet != null)
                                                    {
                                                        strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                    }
                                                    else
                                                    {
                                                        strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                    }
                                                }
                                                else if (de.Value.GetType() == typeof(System.String))
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                                }
                                                strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                                if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                                strContent.AppendLine(strLine + strAttributes);
                                            }
                                        }
                                    }
                                    else if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.UserAccessAuditDetail")
                                    {
                                        var userAccessAuditDetail = ((UserAccessAuditDetail)attrAuditDetail);
                                        strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + userAccessAuditDetail.AccessTime.AddMinutes(330).ToString() + "," + string.Empty + "," + string.Empty;
                                        strContent.AppendLine(strLine + strAttributes);
                                    }
                                    else
                                    {
                                        strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty;
                                        strContent.AppendLine(strLine + strAttributes);
                                    }
                                    _logExists = true;
                                }
                                else {
                                    Console.WriteLine(_action);
                                }
                            }
                            //UploadAzureBlob(ent.Id.ToString(), strContent.ToString());
                            if (strContent.ToString().Length > 0 && _logExists)
                            {
                                UploadAzureBlob(ent.Id.ToString(), strContent.ToString(), container);

                                var DeleteAudit = new DeleteRecordChangeHistoryRequest();
                                DeleteAudit.Target = (EntityReference)ent.ToEntityReference();
                                try
                                {
                                    _service.Execute(DeleteAudit);
                                }
                                catch { }
                            }
                            //entTemp = new Entity(entityLogicalName, ent.Id);
                            //entTemp["hil_isauditlogmigrated"] = true;
                            //_service.Update(entTemp);
                            Console.WriteLine("Record # " + i.ToString() + " End... " + DateTime.Now.ToString());
                            i += 1;
                        }

                        #region Delete Audit Log

                        #endregion
                    }
                    query.PageInfo.PageNumber += 1;
                    query.PageInfo.PagingCookie = ec.PagingCookie;
                    ec = _service.RetrieveMultiple(query);
                } while (ec.MoreRecords);
                Console.WriteLine("DONE !!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }

        static void D365AuditLogMigrationDaily(string _entityName)
        {
            int i = 1;
            string _primaryfield = string.Empty;
            string _primaryKey = string.Empty;
            GetPrimaryIdFieldName(_entityName, out _primaryKey, out _primaryfield);
            List<EntityAttribute> lstEntityAttribute = GetEntityAttributes(_entityName);
            Console.WriteLine("Entity Attribute Metadata successfully fetched..." + DateTime.Now.ToString());
            Entity entTemp = null;
            string entityLogicalName = _entityName;
            string[] primaryKeySchemaName = { _primaryKey, _primaryfield, "createdon" };

            QueryExpression query = new QueryExpression(entityLogicalName);
            query.ColumnSet = new ColumnSet(primaryKeySchemaName);
            query.Criteria.AddCondition("hil_isauditlogmigrated", ConditionOperator.NotEqual, true);
            query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, "2022-05-01");
            //query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, "03082110503597");
            query.AddOrder("createdon", OrderType.Ascending);
            query.NoLock = true;
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;

            try
            {
                EntityCollection ec = _service.RetrieveMultiple(query);
                do
                {
                    if (ec.Entities.Count > 0)
                    {
                        string accName = "d365storagesa";
                        string accKey = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
                        //bool _retValue = false;
                        string _recID = string.Empty;

                        // Implement the accout, set true for https for SSL.  
                        StorageCredentials creds = new StorageCredentials(accName, accKey);
                        CloudStorageAccount strAcc = new CloudStorageAccount(creds, true);
                        // Create the blob client.  
                        CloudBlobClient blobClient = strAcc.CreateCloudBlobClient();
                        // Retrieve a reference to a container.   
                        CloudBlobContainer container = blobClient.GetContainerReference("d365auditlogarchive");
                        // Create the container if it doesn't already exist.  
                        container.CreateIfNotExistsAsync();

                        foreach (Entity ent in ec.Entities)
                        {
                            if (ent.Attributes.Contains(_primaryfield))
                            {
                                _recID = ent.GetAttributeValue<string>(_primaryfield);
                            }
                            else { _recID = ""; }

                            StringBuilder strContent = new StringBuilder();
                            if (!CheckIfAzureBlobExist(ent.Id.ToString(), container))
                            {
                                strContent.AppendLine("auditid,createdon,action,actionname,objectid,objecttypecode,objecttypecodename,operation,operationname,userid,useridname,fieldname,oldvalue,oldvaluename,oldvaluetype,newvalue,newvaluename,newvaluetype");
                            }
                            Console.WriteLine("Record # " + i.ToString() + " " + _recID + " Start..." + ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString());
                            RetrieveRecordChangeHistoryRequest changeRequest = new RetrieveRecordChangeHistoryRequest();
                            changeRequest.Target = new EntityReference(entityLogicalName, ent.Id);
                            RetrieveRecordChangeHistoryResponse changeResponse = null;

                            Repeat:
                            try
                            {
                                changeResponse = (RetrieveRecordChangeHistoryResponse)_service.Execute(changeRequest);
                            }
                            catch
                            {
                                goto Repeat;
                            }

                            AuditDetailCollection auditDetailCollection = changeResponse.AuditDetailCollection;
                            string strLine;
                            string strAttributes;
                            string dateString;
                            bool _logExists = false;
                            foreach (var attrAuditDetail in auditDetailCollection.AuditDetails)
                            {
                                var auditRecord = attrAuditDetail.AuditRecord;
                                strLine = string.Empty;
                                dateString = string.Empty;
                                string _action = auditRecord.FormattedValues["action"];

                                if (auditRecord.Attributes.Contains("objectid") && (_action == "Create" || _action == "Update" || _action == "Delete" || _action == "Assign" || _action == "Activate" || _action == "Deactivate"))
                                {
                                    if (auditRecord.Attributes.Contains("auditid"))
                                    {
                                        strLine = auditRecord.Attributes["auditid"].ToString();
                                    }
                                    else
                                    {
                                        strLine = null;
                                    }
                                    if (auditRecord.Attributes.Contains("createdon"))
                                    {
                                        DateTime dt = ((DateTime)auditRecord.Attributes["createdon"]).AddMinutes(330);
                                        strLine += "," + dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0') + dt.Hour.ToString().PadLeft(2, '0') + dt.Minute.ToString().PadLeft(2, '0') + dt.Second.ToString().PadLeft(2, '0');
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty;
                                    }
                                    if (auditRecord.Attributes.Contains("action"))
                                    {
                                        strLine += "," + ((OptionSetValue)auditRecord.Attributes["action"]).Value.ToString();
                                        strLine += "," + auditRecord.FormattedValues["action"];
                                    }
                                    else
                                    {

                                        strLine += "," + string.Empty + "," + string.Empty;
                                    }
                                    if (auditRecord.Attributes.Contains("objectid"))
                                    {
                                        strLine += "," + ((EntityReference)auditRecord.Attributes["objectid"]).Id.ToString();
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty;
                                    }

                                    if (auditRecord.Attributes.Contains("objecttypecode"))
                                    {
                                        strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                        strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty + "," + string.Empty;
                                    }

                                    if (auditRecord.Attributes.Contains("operation"))
                                    {
                                        strLine += "," + ((OptionSetValue)auditRecord.Attributes["operation"]).Value.ToString();
                                        strLine += "," + auditRecord.FormattedValues["operation"];
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty + "," + string.Empty;
                                    }

                                    if (auditRecord.Attributes.Contains("userid"))
                                    {
                                        EntityReference er = auditRecord.GetAttributeValue<EntityReference>("userid");
                                        strLine += "," + er.Id.ToString();
                                        strLine += "," + er.Name.ToString();
                                    }
                                    else
                                    {
                                        strLine += "," + string.Empty + "," + string.Empty;
                                    }
                                    strLine += ",";
                                    strAttributes = string.Empty;
                                    if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.AttributeAuditDetail")
                                    {
                                        Dictionary<string, Object> oldValues = new Dictionary<string, object>();
                                        Dictionary<string, Object> newValues = new Dictionary<string, object>();
                                        object _OldValue;
                                        //object _NewValue;
                                        var newValueEntity = ((AttributeAuditDetail)attrAuditDetail).NewValue;
                                        if (newValueEntity != null)
                                        {
                                            newValues = newValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                        }

                                        var oldValueEntity = ((AttributeAuditDetail)attrAuditDetail).OldValue;
                                        if (oldValueEntity != null)
                                        {
                                            oldValues = oldValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                        }
                                        if (newValues.Count > 0)
                                        {
                                            foreach (KeyValuePair<string, object> de in newValues)
                                            {
                                                strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                                //strAttributes = de.Key.ToString() + ",";
                                                //Console.WriteLine("Attribute Name :" + strAttributes);
                                                if (oldValues.ContainsKey(de.Key))
                                                {
                                                    _OldValue = oldValues[de.Key.ToString()];
                                                    if (_OldValue.GetType() == typeof(Decimal))
                                                    {
                                                        strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(Double))
                                                    {
                                                        strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(System.DateTime))
                                                    {
                                                        DateTime dtValue = ((DateTime)_OldValue).AddMinutes(330);
                                                        strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                    {
                                                        EntityReference entRef = (EntityReference)_OldValue;
                                                        if (entRef != null)
                                                        {
                                                            strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                        }
                                                        else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                    }
                                                    else if (_OldValue.GetType() == typeof(System.Boolean))
                                                    {
                                                        string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(_OldValue));
                                                        //string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                        strAttributes += _OldValue.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(System.Int32))
                                                    {
                                                        strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                    {
                                                        Money money = (Money)_OldValue;
                                                        strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                    }
                                                    else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                    {
                                                        OptionSetValue optValue = (OptionSetValue)_OldValue;
                                                        EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                        EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                        if (optionSet != null)
                                                        {
                                                            strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                        }
                                                        else
                                                        {
                                                            strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                        }
                                                    }
                                                    else if (_OldValue.GetType() == typeof(System.String))
                                                    {
                                                        strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                    }
                                                    else
                                                    {
                                                        strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                    }

                                                    oldValues.Remove(de.Key.ToString());
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                                }

                                                if (de.Value.GetType() == typeof(Decimal))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Double))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(System.DateTime))
                                                {
                                                    DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                    strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                {
                                                    EntityReference entRef = (EntityReference)de.Value;
                                                    if (entRef != null)
                                                    {
                                                        strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                    }
                                                    else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                }
                                                else if (de.Value.GetType() == typeof(System.Boolean))
                                                {
                                                    string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                    strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(System.Int32))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                {
                                                    Money money = (Money)de.Value;
                                                    strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                {
                                                    OptionSetValue optValue = (OptionSetValue)de.Value;
                                                    //string name = string.Empty;
                                                    //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                    EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                    EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                    if (optionSet != null)
                                                    {
                                                        strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                    }
                                                    else
                                                    {
                                                        strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                    }
                                                }
                                                else if (de.Value.GetType() == typeof(System.String))
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                                }

                                                if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                                strContent.AppendLine(strLine + strAttributes);
                                            }
                                        }
                                        if (oldValues.Count > 0)
                                        {
                                            foreach (KeyValuePair<string, object> de in oldValues)
                                            {
                                                //strAttributes = de.Key.ToString() + ",";
                                                strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                                if (de.Value.GetType() == typeof(Decimal))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Double))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(System.DateTime))
                                                {
                                                    DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                    strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                {
                                                    EntityReference entRef = (EntityReference)de.Value;
                                                    if (entRef != null)
                                                    {
                                                        strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                    }
                                                    else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                }
                                                else if (de.Value.GetType() == typeof(System.Boolean))
                                                {
                                                    string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                    strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(System.Int32))
                                                {
                                                    strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                {
                                                    Money money = (Money)de.Value;
                                                    strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                {
                                                    OptionSetValue optValue = (OptionSetValue)de.Value;
                                                    //string name = string.Empty;
                                                    //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                    EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                    EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                    if (optionSet != null)
                                                    {
                                                        strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                    }
                                                    else
                                                    {
                                                        strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                    }
                                                }
                                                else if (de.Value.GetType() == typeof(System.String))
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                                }
                                                strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                                if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                                strContent.AppendLine(strLine + strAttributes);
                                            }
                                        }
                                    }
                                    else if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.UserAccessAuditDetail")
                                    {
                                        var userAccessAuditDetail = ((UserAccessAuditDetail)attrAuditDetail);
                                        strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + userAccessAuditDetail.AccessTime.AddMinutes(330).ToString() + "," + string.Empty + "," + string.Empty;
                                        strContent.AppendLine(strLine + strAttributes);
                                    }
                                    else
                                    {
                                        strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty;
                                        strContent.AppendLine(strLine + strAttributes);
                                    }
                                    _logExists = true;
                                }
                                else
                                {
                                    Console.WriteLine(_action);
                                }
                            }
                            //UploadAzureBlob(ent.Id.ToString(), strContent.ToString());
                            if (strContent.ToString().Length > 0 && _logExists)
                            {
                                UploadAzureBlob(ent.Id.ToString(), strContent.ToString(), container);

                                var DeleteAudit = new DeleteRecordChangeHistoryRequest();
                                DeleteAudit.Target = (EntityReference)ent.ToEntityReference();
                                try
                                {
                                    _service.Execute(DeleteAudit);
                                }
                                catch { }
                            }
                            entTemp = new Entity(entityLogicalName, ent.Id);
                            entTemp["hil_isauditlogmigrated"] = true;
                            _service.Update(entTemp);
                            Console.WriteLine("Record # " + i.ToString() + " End... " + DateTime.Now.ToString());
                            i += 1;
                        }

                        #region Delete Audit Log

                        #endregion
                    }
                    query.PageInfo.PageNumber += 1;
                    query.PageInfo.PagingCookie = ec.PagingCookie;
                    ec = _service.RetrieveMultiple(query);
                } while (ec.MoreRecords);
                Console.WriteLine("DONE !!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }

        static void D365AuditLogMigration_CustAsset()
        {
            int i = 1;
            Console.WriteLine("Fetching Entity Attribute Metadata..." + DateTime.Now.ToString());
            List<EntityAttribute> lstEntityAttribute = GetEntityAttributes("msdyn_workorder");
            Console.WriteLine("Entity Attribute Metadata successfully fetched..." + DateTime.Now.ToString());
            Entity entTemp = null;
            string entityLogicalName = "msdyn_workorder";
            string[] primaryKeySchemaName = { "msdyn_workorderid", "msdyn_name", "createdon" };

            QueryExpression Query = new QueryExpression(entityLogicalName);
            Query.ColumnSet = new ColumnSet(primaryKeySchemaName);
            Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, "2020-08-01");
            Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrBefore, "2020-09-01");
            Query.Criteria.AddCondition("hil_archivelog", ConditionOperator.Equal, true);
            Query.AddOrder("createdon", OrderType.Ascending);
            Query.PageInfo = new PagingInfo();
            Query.PageInfo.Count = 5000;
            Query.PageInfo.PageNumber = 1;
            Query.PageInfo.ReturnTotalRecordCount = true;

            try
            {
                EntityCollection ec = _service.RetrieveMultiple(Query);
                do
                {
                    if (ec.Entities.Count > 0)
                    {
                        string accName = "d365storagesa";
                        string accKey = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
                        //bool _retValue = false;
                        string _JobNumber = string.Empty;

                        // Implement the accout, set true for https for SSL.  
                        StorageCredentials creds = new StorageCredentials(accName, accKey);
                        CloudStorageAccount strAcc = new CloudStorageAccount(creds, true);
                        // Create the blob client.  
                        CloudBlobClient blobClient = strAcc.CreateCloudBlobClient();
                        // Retrieve a reference to a container.   
                        CloudBlobContainer container = blobClient.GetContainerReference("d365auditlogarchive");
                        // Create the container if it doesn't already exist.  
                        container.CreateIfNotExistsAsync();

                        foreach (Entity ent in ec.Entities)
                        {
                            if (ent.Attributes.Contains("msdyn_name"))
                            {
                                _JobNumber = ent.GetAttributeValue<string>("msdyn_name");
                            }
                            else { _JobNumber = ""; }
                            if (CheckIfAzureBlobExist(ent.Id.ToString(), container))
                            {
                                Console.WriteLine("Job # " + i.ToString() + " " + _JobNumber);
                                try
                                {
                                    var DeleteAudit1 = new DeleteRecordChangeHistoryRequest();
                                    DeleteAudit1.Target = (EntityReference)ent.ToEntityReference();
                                    _service.Execute(DeleteAudit1);
                                    entTemp = new Entity(entityLogicalName, ent.Id);
                                    entTemp["hil_archivelog"] = true;
                                    _service.Update(entTemp);
                                }
                                catch { }
                                i += 1;
                                continue;
                            }
                            Console.WriteLine("Job # " + i.ToString() + " " + _JobNumber + " Start..." + ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString());
                            RetrieveRecordChangeHistoryRequest changeRequest = new RetrieveRecordChangeHistoryRequest();
                            changeRequest.Target = new EntityReference(entityLogicalName, ent.Id);
                            RetrieveRecordChangeHistoryResponse changeResponse =
                            (RetrieveRecordChangeHistoryResponse)_service.Execute(changeRequest);

                            AuditDetailCollection auditDetailCollection = changeResponse.AuditDetailCollection;
                            StringBuilder strContent = new StringBuilder();
                            strContent.AppendLine("auditid,createdon,action,actionname,objectid,objecttypecode,objecttypecodename,operation,operationname,userid,useridname,fieldname,oldvalue,oldvaluename,oldvaluetype,newvalue,newvaluename,newvaluetype");
                            string strLine;
                            string strAttributes;
                            string dateString;
                            foreach (var attrAuditDetail in auditDetailCollection.AuditDetails)
                            {
                                var auditRecord = attrAuditDetail.AuditRecord;
                                strLine = string.Empty;
                                dateString = string.Empty;
                                if (auditRecord.Attributes.Contains("auditid"))
                                {
                                    strLine = auditRecord.Attributes["auditid"].ToString();
                                }
                                else
                                {
                                    strLine = null;
                                }
                                if (auditRecord.Attributes.Contains("createdon"))
                                {
                                    DateTime dt = ((DateTime)auditRecord.Attributes["createdon"]).AddMinutes(330);
                                    strLine += "," + dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0') + dt.Hour.ToString().PadLeft(2, '0') + dt.Minute.ToString().PadLeft(2, '0') + dt.Second.ToString().PadLeft(2, '0');
                                }
                                else
                                {
                                    strLine += "," + string.Empty;
                                }
                                if (auditRecord.Attributes.Contains("action"))
                                {
                                    strLine += "," + ((OptionSetValue)auditRecord.Attributes["action"]).Value.ToString();
                                    strLine += "," + auditRecord.FormattedValues["action"];
                                }
                                else
                                {

                                    strLine += "," + string.Empty + "," + string.Empty;
                                }
                                if (auditRecord.Attributes.Contains("objectid"))
                                {
                                    strLine += "," + ((EntityReference)auditRecord.Attributes["objectid"]).Id.ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("objecttypecode"))
                                {
                                    strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                    strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("operation"))
                                {
                                    strLine += "," + ((OptionSetValue)auditRecord.Attributes["operation"]).Value.ToString();
                                    strLine += "," + auditRecord.FormattedValues["operation"];
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("userid"))
                                {
                                    EntityReference er = auditRecord.GetAttributeValue<EntityReference>("userid");
                                    strLine += "," + er.Id.ToString();
                                    strLine += "," + er.Name.ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }
                                strLine += ",";
                                strAttributes = string.Empty;
                                if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.AttributeAuditDetail")
                                {
                                    Dictionary<string, Object> oldValues = new Dictionary<string, object>();
                                    Dictionary<string, Object> newValues = new Dictionary<string, object>();
                                    object _OldValue;
                                    //object _NewValue;
                                    var newValueEntity = ((AttributeAuditDetail)attrAuditDetail).NewValue;
                                    if (newValueEntity != null)
                                    {
                                        newValues = newValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                    }

                                    var oldValueEntity = ((AttributeAuditDetail)attrAuditDetail).OldValue;
                                    if (oldValueEntity != null)
                                    {
                                        oldValues = oldValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                    }
                                    if (newValues.Count > 0)
                                    {
                                        foreach (KeyValuePair<string, object> de in newValues)
                                        {
                                            strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                            //strAttributes = de.Key.ToString() + ",";
                                            //Console.WriteLine("Attribute Name :" + strAttributes);
                                            if (oldValues.ContainsKey(de.Key))
                                            {
                                                _OldValue = oldValues[de.Key.ToString()];
                                                if (_OldValue.GetType() == typeof(Decimal))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Double))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(System.DateTime))
                                                {
                                                    DateTime dtValue = ((DateTime)_OldValue).AddMinutes(330);
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                {
                                                    EntityReference entRef = (EntityReference)_OldValue;
                                                    if (entRef != null)
                                                    {
                                                        strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                    }
                                                    else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                }
                                                else if (_OldValue.GetType() == typeof(System.Boolean))
                                                {
                                                    string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                    strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(System.Int32))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                {
                                                    Money money = (Money)_OldValue;
                                                    strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                {
                                                    OptionSetValue optValue = (OptionSetValue)_OldValue;
                                                    EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                    EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                    if (optionSet != null)
                                                    {
                                                        strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                    }
                                                    else
                                                    {
                                                        strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                    }
                                                }
                                                else if (_OldValue.GetType() == typeof(System.String))
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                }

                                                oldValues.Remove(de.Key.ToString());
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                            }

                                            if (de.Value.GetType() == typeof(Decimal))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Double))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.DateTime))
                                            {
                                                DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                            {
                                                EntityReference entRef = (EntityReference)de.Value;
                                                if (entRef != null)
                                                {
                                                    strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                }
                                                else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                            }
                                            else if (de.Value.GetType() == typeof(System.Boolean))
                                            {
                                                string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.Int32))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                            {
                                                Money money = (Money)de.Value;
                                                strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                            {
                                                OptionSetValue optValue = (OptionSetValue)de.Value;
                                                //string name = string.Empty;
                                                //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                if (optionSet != null)
                                                {
                                                    strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                }
                                                else
                                                {
                                                    strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                }
                                            }
                                            else if (de.Value.GetType() == typeof(System.String))
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }

                                            if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                            strContent.AppendLine(strLine + strAttributes);
                                        }
                                    }
                                    if (oldValues.Count > 0)
                                    {
                                        foreach (KeyValuePair<string, object> de in oldValues)
                                        {
                                            //strAttributes = de.Key.ToString() + ",";
                                            strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                            if (de.Value.GetType() == typeof(Decimal))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Double))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.DateTime))
                                            {
                                                DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                            {
                                                EntityReference entRef = (EntityReference)de.Value;
                                                if (entRef != null)
                                                {
                                                    strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                }
                                                else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                            }
                                            else if (de.Value.GetType() == typeof(System.Boolean))
                                            {
                                                string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.Int32))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                            {
                                                Money money = (Money)de.Value;
                                                strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                            {
                                                OptionSetValue optValue = (OptionSetValue)de.Value;
                                                //string name = string.Empty;
                                                //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                if (optionSet != null)
                                                {
                                                    strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                }
                                                else
                                                {
                                                    strAttributes += optValue.Value.ToString() + ",," + string.Empty + ",";
                                                }
                                            }
                                            else if (de.Value.GetType() == typeof(System.String))
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                            if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                            strContent.AppendLine(strLine + strAttributes);
                                        }
                                    }
                                }
                                else if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.UserAccessAuditDetail")
                                {
                                    var userAccessAuditDetail = ((UserAccessAuditDetail)attrAuditDetail);
                                    strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + userAccessAuditDetail.AccessTime.AddMinutes(330).ToString() + "," + string.Empty + "," + string.Empty;
                                    strContent.AppendLine(strLine + strAttributes);
                                }
                                else
                                {
                                    strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty;
                                    strContent.AppendLine(strLine + strAttributes);
                                }
                            }
                            //UploadAzureBlob(ent.Id.ToString(), strContent.ToString());
                            UploadAzureBlob(ent.Id.ToString(), strContent.ToString(), container);

                            var DeleteAudit = new DeleteRecordChangeHistoryRequest();
                            DeleteAudit.Target = (EntityReference)ent.ToEntityReference();
                            try
                            {
                                _service.Execute(DeleteAudit);
                                entTemp = new Entity(entityLogicalName, ent.Id);
                                entTemp["hil_archivelog"] = true;
                                _service.Update(entTemp);
                            }
                            catch { }
                            Console.WriteLine("Job # " + i.ToString() + " End... " + DateTime.Now.ToString());
                            i += 1;
                        }

                        #region Delete Audit Log

                        #endregion
                    }
                    Query.PageInfo.PageNumber += 1;
                    Query.PageInfo.PagingCookie = ec.PagingCookie;
                    ec = _service.RetrieveMultiple(Query);
                } while (ec.MoreRecords);
                Console.WriteLine("DONE !!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        static void D365AuditLogMigrationSystemUser()
        {
            int i = 1;
            Console.WriteLine("Fetching Entity Attribute Metadata..." + DateTime.Now.ToString());
            List<EntityAttribute> lstEntityAttribute = GetEntityAttributes("systemuser");
            Console.WriteLine("Entity Attribute Metadata successfully fetched..." + DateTime.Now.ToString());
            Entity entTemp = null;
            string entityLogicalName = "systemuser";
            string[] primaryKeySchemaName = { "systemuserid", "internalemailaddress" };
            DateTime _createdOn = new DateTime(1900, 1, 1);

            string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='audit'>
            <attribute name='objecttypecode' />
            <filter type='and'>
                <condition attribute='objecttypecode' operator='eq' value='8' />
                <condition attribute='createdon' operator='on-or-before' value='2020-09-01' />
            </filter>
            </entity>
            </fetch>";

            EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
            if (EntityList.Entities.Count > 0)
            {
                _createdOn = EntityList.Entities[0].GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                //_createdOn = EntityList.Entities[0].GetAttributeValue<DateTime>("ad.createdon").AddMinutes(330);
            }

            QueryExpression Query = new QueryExpression(entityLogicalName);
            Query.ColumnSet = new ColumnSet(primaryKeySchemaName);
            Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, _createdOn.Year.ToString() + "-" + _createdOn.Month.ToString().PadLeft(2, '0') + "-" + _createdOn.Day.ToString().PadLeft(2, '0'));
            Query.Criteria.AddCondition("isdisabled", ConditionOperator.Equal, false);
            Query.Criteria.AddCondition("accessmode", ConditionOperator.NotEqual, 3);
            Query.Criteria.AddCondition("accessmode", ConditionOperator.NotEqual, 5);
            Query.Criteria.AddCondition("fullname", ConditionOperator.NotEqual, "SYSTEM");
            Query.Criteria.AddCondition("fullname", ConditionOperator.NotEqual, "INTEGRATION");
            //Query.Criteria.AddCondition("hil_archivelog", ConditionOperator.NotEqual, true);
            Query.AddOrder("createdon", OrderType.Ascending);
            Query.PageInfo = new PagingInfo();
            Query.PageInfo.Count = 5000;
            Query.PageInfo.PageNumber = 1;
            Query.PageInfo.ReturnTotalRecordCount = true;

            try
            {
                EntityCollection ec = _service.RetrieveMultiple(Query);
                do
                {
                    if (ec.Entities.Count > 0)
                    {
                        string accName = "d365storagesa";
                        string accKey = "6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==";
                        //bool _retValue = false;
                        // Implement the accout, set true for https for SSL.  
                        StorageCredentials creds = new StorageCredentials(accName, accKey);
                        CloudStorageAccount strAcc = new CloudStorageAccount(creds, true);
                        // Create the blob client.  
                        CloudBlobClient blobClient = strAcc.CreateCloudBlobClient();
                        // Retrieve a reference to a container.   
                        CloudBlobContainer container = blobClient.GetContainerReference("d365auditlogarchive");
                        // Create the container if it doesn't already exist.  
                        container.CreateIfNotExistsAsync();

                        foreach (Entity ent in ec.Entities)
                        {
                            //CheckIfAzureBlobExist(ent.Id.ToString(), container) || 
                            if (ent.GetAttributeValue<string>("internalemailaddress") == "havellconnect@havells.com")
                            {
                                Console.WriteLine("User # " + i.ToString() + " " + ent.GetAttributeValue<string>("internalemailaddress"));
                                //try
                                //{
                                //    entTemp = new Entity("systemuser", ent.Id);
                                //    entTemp["hil_archivelog"] = true;
                                //    _service.Update(entTemp);
                                //}
                                //catch  { }
                                i += 1;
                                continue;
                            }
                            Console.WriteLine("User # " + i.ToString() + " " + ent.GetAttributeValue<string>("internalemailaddress") + " Start..." + DateTime.Now.ToString());
                            RetrieveRecordChangeHistoryRequest changeRequest = new RetrieveRecordChangeHistoryRequest();
                            changeRequest.Target = new EntityReference(entityLogicalName, ent.Id);
                            RetrieveRecordChangeHistoryResponse changeResponse =
                            (RetrieveRecordChangeHistoryResponse)_service.Execute(changeRequest);

                            AuditDetailCollection auditDetailCollection = changeResponse.AuditDetailCollection;
                            StringBuilder strContent = new StringBuilder();
                            strContent.AppendLine("auditid,createdon,action,actionname,objectid,objecttypecode,objecttypecodename,operation,operationname,userid,useridname,fieldname,oldvalue,oldvaluename,oldvaluetype,newvalue,newvaluename,newvaluetype");
                            string strLine;
                            string strAttributes;
                            string dateString;
                            foreach (var attrAuditDetail in auditDetailCollection.AuditDetails)
                            {
                                var auditRecord = attrAuditDetail.AuditRecord;
                                strLine = string.Empty;
                                dateString = string.Empty;
                                if (auditRecord.Attributes.Contains("auditid"))
                                {
                                    strLine = auditRecord.Attributes["auditid"].ToString();
                                }
                                else
                                {
                                    strLine = null;
                                }
                                if (auditRecord.Attributes.Contains("createdon"))
                                {
                                    DateTime dt = ((DateTime)auditRecord.Attributes["createdon"]).AddMinutes(330);
                                    strLine += "," + dt.Year.ToString() + dt.Month.ToString().PadLeft(2, '0') + dt.Day.ToString().PadLeft(2, '0') + dt.Hour.ToString().PadLeft(2, '0') + dt.Minute.ToString().PadLeft(2, '0') + dt.Second.ToString().PadLeft(2, '0');
                                }
                                else
                                {
                                    strLine += "," + string.Empty;
                                }
                                if (auditRecord.Attributes.Contains("action"))
                                {
                                    strLine += "," + ((OptionSetValue)auditRecord.Attributes["action"]).Value.ToString();
                                    strLine += "," + auditRecord.FormattedValues["action"];
                                }
                                else
                                {

                                    strLine += "," + string.Empty + "," + string.Empty;
                                }
                                if (auditRecord.Attributes.Contains("objectid"))
                                {
                                    strLine += "," + ((EntityReference)auditRecord.Attributes["objectid"]).Id.ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("objecttypecode"))
                                {
                                    strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                    strLine += "," + auditRecord.Attributes["objecttypecode"].ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("operation"))
                                {
                                    strLine += "," + ((OptionSetValue)auditRecord.Attributes["operation"]).Value.ToString();
                                    strLine += "," + auditRecord.FormattedValues["operation"];
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }

                                if (auditRecord.Attributes.Contains("userid"))
                                {
                                    EntityReference er = auditRecord.GetAttributeValue<EntityReference>("userid");
                                    strLine += "," + er.Id.ToString();
                                    strLine += "," + er.Name.ToString();
                                }
                                else
                                {
                                    strLine += "," + string.Empty + "," + string.Empty;
                                }
                                strLine += ",";
                                strAttributes = string.Empty;
                                if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.AttributeAuditDetail")
                                {
                                    Dictionary<string, Object> oldValues = new Dictionary<string, object>();
                                    Dictionary<string, Object> newValues = new Dictionary<string, object>();
                                    object _OldValue;
                                    //object _NewValue;
                                    var newValueEntity = ((AttributeAuditDetail)attrAuditDetail).NewValue;
                                    if (newValueEntity != null)
                                    {
                                        newValues = newValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                    }

                                    var oldValueEntity = ((AttributeAuditDetail)attrAuditDetail).OldValue;
                                    if (oldValueEntity != null)
                                    {
                                        oldValues = oldValueEntity.Attributes.ToDictionary(v => v.Key, v => v.Value);
                                    }
                                    if (newValues.Count > 0)
                                    {
                                        foreach (KeyValuePair<string, object> de in newValues)
                                        {
                                            strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                            //strAttributes = de.Key.ToString() + ",";
                                            //Console.WriteLine("Attribute Name :" + strAttributes);
                                            if (oldValues.ContainsKey(de.Key))
                                            {
                                                _OldValue = oldValues[de.Key.ToString()];
                                                if (_OldValue.GetType() == typeof(Decimal))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Double))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(System.DateTime))
                                                {
                                                    DateTime dtValue = ((DateTime)_OldValue).AddMinutes(330);
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                                {
                                                    EntityReference entRef = (EntityReference)_OldValue;
                                                    if (entRef != null)
                                                    {
                                                        strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                    }
                                                    else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                                }
                                                else if (_OldValue.GetType() == typeof(System.Boolean))
                                                {
                                                    string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                    strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(System.Int32))
                                                {
                                                    strAttributes += string.Empty + "," + _OldValue.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                                {
                                                    Money money = (Money)_OldValue;
                                                    strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                                {
                                                    OptionSetValue optValue = (OptionSetValue)_OldValue;
                                                    EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                    EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                    strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                                }
                                                else if (_OldValue.GetType() == typeof(System.String))
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                }
                                                else
                                                {
                                                    strAttributes += string.Empty + "," + RemoveSpecialCharacters(_OldValue.ToString()) + "," + string.Empty + ",";
                                                }

                                                oldValues.Remove(de.Key.ToString());
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                            }

                                            if (de.Value.GetType() == typeof(Decimal))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Double))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.DateTime))
                                            {
                                                DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                            {
                                                EntityReference entRef = (EntityReference)de.Value;
                                                if (entRef != null)
                                                {
                                                    strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                }
                                                else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                            }
                                            else if (de.Value.GetType() == typeof(System.Boolean))
                                            {
                                                string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.Int32))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                            {
                                                Money money = (Money)de.Value;
                                                strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                            {
                                                OptionSetValue optValue = (OptionSetValue)de.Value;
                                                //string name = string.Empty;
                                                //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.String))
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }

                                            if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                            strContent.AppendLine(strLine + strAttributes);
                                        }
                                    }
                                    if (oldValues.Count > 0)
                                    {
                                        foreach (KeyValuePair<string, object> de in oldValues)
                                        {
                                            //strAttributes = de.Key.ToString() + ",";
                                            strAttributes = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString()).DisplayName + ",";
                                            if (de.Value.GetType() == typeof(Decimal))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Double))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.DateTime))
                                            {
                                                DateTime dtValue = ((DateTime)de.Value).AddMinutes(330);
                                                strAttributes += string.Empty + "," + dtValue.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.EntityReference))
                                            {
                                                EntityReference entRef = (EntityReference)de.Value;
                                                if (entRef != null)
                                                {
                                                    strAttributes += entRef.Id.ToString() + "," + RemoveSpecialCharacters(entRef.Name == null ? "" : entRef.Name.ToString()) + "," + entRef.LogicalName + ",";
                                                }
                                                else { strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ","; }
                                            }
                                            else if (de.Value.GetType() == typeof(System.Boolean))
                                            {
                                                string boolLabel = GetBoolText(_service, entityLogicalName, de.Key.ToString(), Convert.ToBoolean(de.Value.ToString()));
                                                strAttributes += de.Value.ToString() + "," + boolLabel + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.Int32))
                                            {
                                                strAttributes += string.Empty + "," + de.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Money))
                                            {
                                                Money money = (Money)de.Value;
                                                strAttributes += string.Empty + "," + money.Value.ToString() + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(Microsoft.Xrm.Sdk.OptionSetValue))
                                            {
                                                OptionSetValue optValue = (OptionSetValue)de.Value;
                                                //string name = string.Empty;
                                                //name = GetOptionSetLabelByValue(_service, entityLogicalName, de.Key.ToString(), optValue.Value, 1036);
                                                EntityAttribute entAttr = lstEntityAttribute.FirstOrDefault(o => o.LogicalName == de.Key.ToString());
                                                EntityOptionSet optionSet = entAttr.Optionset.FirstOrDefault(o => o.OptionValue == optValue.Value);
                                                strAttributes += optValue.Value.ToString() + "," + RemoveSpecialCharacters(optionSet.OptionName) + "," + string.Empty + ",";
                                            }
                                            else if (de.Value.GetType() == typeof(System.String))
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            else
                                            {
                                                strAttributes += string.Empty + "," + RemoveSpecialCharacters(de.Value.ToString()) + "," + string.Empty + ",";
                                            }
                                            strAttributes += string.Empty + "," + string.Empty + "," + string.Empty + ",";
                                            if (strAttributes.Length > 0) { strAttributes = strAttributes.Substring(0, strAttributes.Length - 1); }
                                            strContent.AppendLine(strLine + strAttributes);
                                        }
                                    }
                                }
                                else if (attrAuditDetail.GetType().ToString() == "Microsoft.Crm.Sdk.Messages.UserAccessAuditDetail")
                                {
                                    var userAccessAuditDetail = ((UserAccessAuditDetail)attrAuditDetail);
                                    strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + userAccessAuditDetail.AccessTime.AddMinutes(330).ToString() + "," + string.Empty + "," + string.Empty;
                                    strContent.AppendLine(strLine + strAttributes);
                                }
                                else
                                {
                                    strAttributes = "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty + "," + string.Empty;
                                    strContent.AppendLine(strLine + strAttributes);
                                }
                            }
                            //UploadAzureBlob(ent.Id.ToString(), strContent.ToString());
                            UploadAzureBlob(ent.Id.ToString(), strContent.ToString(), container);

                            var DeleteAudit = new DeleteRecordChangeHistoryRequest();
                            DeleteAudit.Target = (EntityReference)ent.ToEntityReference();
                            try
                            {
                                _service.Execute(DeleteAudit);
                                entTemp = new Entity("systemuser", ent.Id);
                                entTemp["hil_archivelog"] = true;
                                _service.Update(entTemp);
                            }
                            catch { }
                            Console.WriteLine("User # " + i.ToString() + " End... " + DateTime.Now.ToString());
                            i += 1;
                        }

                        #region Delete Audit Log

                        #endregion
                    }
                    Query.PageInfo.PageNumber += 1;
                    Query.PageInfo.PagingCookie = ec.PagingCookie;
                    ec = _service.RetrieveMultiple(Query);
                } while (ec.MoreRecords);
                Console.WriteLine("DONE !!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        static void DeleteAuditLogPartitionWise()
        {

            // Get the list of audit partitions.
            RetrieveAuditPartitionListResponse partitionRequest =
            (RetrieveAuditPartitionListResponse)_service.Execute(new RetrieveAuditPartitionListRequest());
            AuditPartitionDetailCollection partitions = partitionRequest.AuditPartitionDetailCollection;

            // Create a delete request with an end date earlier than possible.
            DeleteAuditDataRequest deleteRequest = new DeleteAuditDataRequest();
            deleteRequest.EndDate = new DateTime(2019, 01, 01);

            // Check if partitions are not supported as is the case with SQL Server Standard edition.
            if (partitions.IsLogicalCollection)
            {
                deleteRequest.EndDate = DateTime.Now;
            }
            else
            {
                for (int n = partitions.Count - 1; n >= 0; --n)
                {
                    if (partitions[n].EndDate <= DateTime.Now && partitions[n].EndDate > deleteRequest.EndDate)
                    {
                        deleteRequest.EndDate = (DateTime)partitions[n].EndDate;
                        break;
                    }
                }
            }

            // Delete the audit records.
            if (deleteRequest.EndDate != new DateTime(2018, 12, 10))
            {
                _service.Execute(deleteRequest);
                Console.WriteLine("Audit records have been deleted.");
            }
            else
                Console.WriteLine("There were no audit records that could be deleted.");


        }
        static void D365AuditLogMigrationProductRequest()
        {
            int i = 1;

            string entityLogicalName = "hil_productrequest";
            string[] primaryKeySchemaName = { "hil_productrequestid", "hil_name", "createdon" };
            string _RecID = string.Empty;
            Entity _updateArchiveLog = null;

            QueryExpression Query = new QueryExpression(entityLogicalName);
            Query.ColumnSet = new ColumnSet(primaryKeySchemaName);
            Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, "2018-01-01");
            Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrBefore, "2019-02-01");
            Query.Criteria.AddCondition("hil_archivelog", ConditionOperator.NotEqual, true);
            Query.AddOrder("createdon", OrderType.Ascending);
            Query.PageInfo = new PagingInfo();
            Query.PageInfo.Count = 5000;
            Query.PageInfo.PageNumber = 1;
            Query.PageInfo.ReturnTotalRecordCount = true;

            try
            {
                EntityCollection ec = _service.RetrieveMultiple(Query);
                do
                {
                    if (ec.Entities.Count > 0)
                    {
                        foreach (Entity ent in ec.Entities)
                        {
                            try
                            {
                                _RecID = ent.GetAttributeValue<string>("hil_name");
                                Console.WriteLine("PR # " + i.ToString() + " " + _RecID);
                                var DeleteAudit1 = new DeleteRecordChangeHistoryRequest();
                                DeleteAudit1.Target = (EntityReference)ent.ToEntityReference();
                                _service.Execute(DeleteAudit1);
                                _updateArchiveLog = new Entity("hil_productrequest", ent.Id);
                                _updateArchiveLog["hil_archivelog"] = true;
                                _service.Update(_updateArchiveLog);
                            }
                            catch { }
                            i += 1;
                        }
                    }
                    Query.PageInfo.PageNumber += 1;
                    Query.PageInfo.PagingCookie = ec.PagingCookie;
                    ec = _service.RetrieveMultiple(Query);
                } while (ec.MoreRecords);
                Console.WriteLine("DONE !!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        static void D365AuditLogMigrationProductRequestHeader()
        {
            int i = 1;
            string entityLogicalName = "hil_productrequestheader";
            string primaryFieldName = "hil_name";
            string[] primaryKeySchemaName = { "hil_productrequestheaderid", "hil_name", "createdon" };
            string _RecID = string.Empty;
            Entity _updateArchiveLog = null;
            QueryExpression Query = new QueryExpression(entityLogicalName);
            Query.ColumnSet = new ColumnSet(primaryKeySchemaName);
            Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, "2020-08-01");
            Query.Criteria.AddCondition("hil_archivelog", ConditionOperator.NotEqual, true);
            Query.AddOrder("createdon", OrderType.Ascending);
            Query.PageInfo = new PagingInfo();
            Query.PageInfo.Count = 5000;
            Query.PageInfo.PageNumber = 1;
            Query.PageInfo.ReturnTotalRecordCount = true;

            try
            {
                EntityCollection ec = _service.RetrieveMultiple(Query);
                do
                {
                    if (ec.Entities.Count > 0)
                    {
                        foreach (Entity ent in ec.Entities)
                        {
                            try
                            {
                                _RecID = ent.GetAttributeValue<string>(primaryFieldName);
                                Console.WriteLine("PRH # " + i.ToString() + " " + _RecID);
                                var DeleteAudit1 = new DeleteRecordChangeHistoryRequest();
                                DeleteAudit1.Target = (EntityReference)ent.ToEntityReference();
                                _service.Execute(DeleteAudit1);
                                _updateArchiveLog = new Entity(entityLogicalName, ent.Id);
                                _updateArchiveLog["hil_archivelog"] = true;
                                _service.Update(_updateArchiveLog);
                            }
                            catch { }
                            i += 1;
                        }
                    }
                    Query.PageInfo.PageNumber += 1;
                    Query.PageInfo.PagingCookie = ec.PagingCookie;
                    ec = _service.RetrieveMultiple(Query);
                } while (ec.MoreRecords);
                Console.WriteLine("DONE !!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        static void D365AuditLogMigrationInventoryRequest()
        {
            int i = 1;
            string entityLogicalName = "hil_inventoryrequest";
            string primaryFieldName = "hil_name";
            string[] primaryKeySchemaName = { "hil_inventoryrequestid", "hil_name", "createdon" };
            string _RecID = string.Empty;
            Entity _updateArchiveLog = null;
            QueryExpression Query = new QueryExpression(entityLogicalName);
            Query.ColumnSet = new ColumnSet(primaryKeySchemaName);
            Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, "2019-01-01");
            Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrBefore, "2019-02-01");
            Query.Criteria.AddCondition("hil_archivelog", ConditionOperator.NotEqual, true);
            Query.AddOrder("createdon", OrderType.Ascending);
            Query.PageInfo = new PagingInfo();
            Query.PageInfo.Count = 5000;
            Query.PageInfo.PageNumber = 1;
            Query.PageInfo.ReturnTotalRecordCount = true;

            try
            {
                EntityCollection ec = _service.RetrieveMultiple(Query);
                do
                {
                    if (ec.Entities.Count > 0)
                    {
                        foreach (Entity ent in ec.Entities)
                        {
                            try
                            {
                                _RecID = ent.GetAttributeValue<string>(primaryFieldName);
                                Console.WriteLine("IR # " + i.ToString() + " " + _RecID);
                                var DeleteAudit1 = new DeleteRecordChangeHistoryRequest();
                                DeleteAudit1.Target = (EntityReference)ent.ToEntityReference();
                                _service.Execute(DeleteAudit1);
                                _updateArchiveLog = new Entity(entityLogicalName, ent.Id);
                                _updateArchiveLog["hil_archivelog"] = true;
                                _service.Update(_updateArchiveLog);
                            }
                            catch { }
                            i += 1;
                        }
                    }
                    Query.PageInfo.PageNumber += 1;
                    Query.PageInfo.PagingCookie = ec.PagingCookie;
                    ec = _service.RetrieveMultiple(Query);
                } while (ec.MoreRecords);
                Console.WriteLine("DONE !!!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR !!! " + ex.Message);
            }
        }
        static void GetPrimaryIdFieldName(string _entityName,out string _primaryKey,out string _primaryField) {
            //Create RetrieveEntityRequest
            _primaryKey = "";
            _primaryField = "";
            try
            {
                RetrieveEntityRequest retrievesEntityRequest = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.Entity,
                    LogicalName = _entityName
                };

                //Execute Request
                RetrieveEntityResponse retrieveEntityResponse = (RetrieveEntityResponse)_service.Execute(retrievesEntityRequest);
                _primaryKey = retrieveEntityResponse.EntityMetadata.PrimaryIdAttribute;

                _primaryField = retrieveEntityResponse.EntityMetadata.PrimaryNameAttribute;
            }
            catch (Exception ex)
            {
                _primaryKey = "ERROR";
                _primaryKey = ex.Message;
            }
        }
        #region Common Libraries
        static string RemoveSpecialCharacters(string _attributeValue)
        {
            return _attributeValue.Replace(", ", " ").Replace(",", " ").Replace("\n", " ").Replace("\r", " ").Replace("\t", " ").Replace("@", "at");
        }
        static string GetBoolText(IOrganizationService service, string entitySchemaName, string attributeSchemaName, bool value)
        {
            RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entitySchemaName,
                LogicalName = attributeSchemaName,
                RetrieveAsIfPublished = true
            };
            RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)service.Execute(retrieveAttributeRequest);
            BooleanAttributeMetadata retrievedBooleanAttributeMetadata = (BooleanAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
            string boolText = string.Empty;
            if (value)
            {
                boolText = retrievedBooleanAttributeMetadata.OptionSet.TrueOption.Label.UserLocalizedLabel.Label;
            }
            else
            {
                boolText = retrievedBooleanAttributeMetadata.OptionSet.FalseOption.Label.UserLocalizedLabel.Label;
            }
            return boolText;
        }
        static void UploadAzureBlob(string _fileName, string _strContent, CloudBlobContainer container)
        {
            // This creates a reference to the append blob.  
            CloudAppendBlob appBlob = container.GetAppendBlobReference(_fileName + ".csv");

            // Now we are going to check if todays file exists and if it doesn't we create it.  
            if (!appBlob.Exists())
            {
                appBlob.CreateOrReplace();
            }

            // Add the entry to file.  
            _strContent = _strContent.Replace("@", System.Environment.NewLine);
            appBlob.AppendText
            (
            string.Format(
                    "{0}\r\n",
                    _strContent)
             );
        }

        static bool CheckIfAzureBlobExist(string _fileName, CloudBlobContainer container)
        {
            bool _retValue = false;
            CloudAppendBlob appBlob = container.GetAppendBlobReference(_fileName + ".csv");

            if (appBlob.Exists())
            {
                _retValue = true;
            }
            return _retValue;
        }
        #endregion

        #region App Setting Load/CRM Connection
        static IOrganizationService ConnectToCRM(string connectionString)
        {
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;

        }
        #endregion
    }

    public class EntityAttribute
    {
        public string LogicalName { get; set; }
        public string DisplayName { get; set; }
        public string AttributeType { get; set; }
        public List<EntityOptionSet> Optionset { get; set; }
    }
    public class EntityOptionSet
    {
        public string OptionName { get; set; }
        public int? OptionValue { get; set; }
    }
}
