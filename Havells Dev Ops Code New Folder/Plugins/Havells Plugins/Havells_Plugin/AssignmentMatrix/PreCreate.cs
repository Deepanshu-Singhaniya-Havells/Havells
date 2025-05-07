using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.AssignmentMatrix
{
    public class PreCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            Entity entity = (Entity)context.InputParameters["Target"];
            Guid _currentRecId = Guid.Empty;
            Guid _FranchiseeGuid = Guid.Empty;
            QueryExpression qsAssignMatrix = null;
            ConditionExpression condExp = null;
            EntityCollection collct_assMatrix = null;
            EntityReference erDivision = null;
            EntityReference erFranchise = null;
            EntityReference erBSH = null;
            try
            {
                EntityReference callSubType = entity.GetAttributeValue<EntityReference>("hil_callsubtype");
                //EntityReference division = entity.GetAttributeValue<EntityReference>("hil_division");
                string _divisionCode = entity.GetAttributeValue<string>("hil_divisioncode");
                EntityReference pinCode = entity.GetAttributeValue<EntityReference>("hil_pincode");
                //EntityReference franchiseDSE = entity.GetAttributeValue<EntityReference>("hil_franchisedirectengineer");
                string _franchiseDSECode = entity.GetAttributeValue<string>("hil_franchisedirectengineercode");
                //EntityReference owner = entity.GetAttributeValue<EntityReference>("hil_branchservicehead");
                string _ownerEmailId = entity.GetAttributeValue<string>("hil_branchserviceheademailid");
                OptionSetValue actionTaken = entity.GetAttributeValue<OptionSetValue>("hil_actionrequired");
                bool? _upCountry = entity.GetAttributeValue<bool>("hil_upcountry");

                if (callSubType == null)
                {
                    entity["hil_uploadstatus"] = false;
                    entity["hil_errordescription"] = "Call Sub Type is required.";
                }
                else if (_divisionCode == null || _divisionCode.Trim().Length == 0)
                {
                    entity["hil_uploadstatus"] = false;
                    entity["hil_errordescription"] = "Division Code is required.";
                }
                else if (pinCode == null)
                {
                    entity["hil_uploadstatus"] = false;
                    entity["hil_errordescription"] = "PIN Code is required.";
                }
                else if (_franchiseDSECode == null || _franchiseDSECode.Trim().Length == 0)
                {
                    entity["hil_uploadstatus"] = false;
                    entity["hil_errordescription"] = "Franchise/DSE Code is required.";
                }
                else if (_ownerEmailId == null || _ownerEmailId.Trim().Length == 0)
                {
                    entity["hil_uploadstatus"] = false;
                    entity["hil_errordescription"] = "Owner/BSH Email Id is required.";
                }
                else if (_upCountry == null && actionTaken.Value == 4)
                {
                    entity["hil_uploadstatus"] = false;
                    entity["hil_errordescription"] = "Upcountry is required.";
                }
                else
                {
                    qsAssignMatrix = new QueryExpression("product");
                    qsAssignMatrix.ColumnSet = new ColumnSet("productid");
                    condExp = new ConditionExpression("hil_hierarchylevel", ConditionOperator.Equal, 2);
                    qsAssignMatrix.Criteria.AddCondition(condExp);
                    condExp = new ConditionExpression("hil_sapcode", ConditionOperator.Equal, _divisionCode);
                    qsAssignMatrix.Criteria.AddCondition(condExp);
                    condExp = new ConditionExpression("name", ConditionOperator.NotLike, "Defective Spare Return");
                    qsAssignMatrix.Criteria.AddCondition(condExp);
                    collct_assMatrix = service.RetrieveMultiple(qsAssignMatrix);
                    if (collct_assMatrix.Entities.Count >= 1)
                    {
                        erDivision = collct_assMatrix.Entities[0].ToEntityReference();
                    }
                    else
                    {
                        entity["hil_uploadstatus"] = false;
                        entity["hil_errordescription"] = "Division Code does not exist.";
                    }

                    qsAssignMatrix = new QueryExpression("account");
                    qsAssignMatrix.ColumnSet = new ColumnSet("accountid");
                    condExp = new ConditionExpression("accountnumber", ConditionOperator.Equal, _franchiseDSECode);
                    qsAssignMatrix.Criteria.AddCondition(condExp);
                    collct_assMatrix = service.RetrieveMultiple(qsAssignMatrix);
                    if (collct_assMatrix.Entities.Count >= 1)
                    {
                        erFranchise = collct_assMatrix.Entities[0].ToEntityReference();
                    }
                    else
                    {
                        entity["hil_uploadstatus"] = false;
                        entity["hil_errordescription"] = "Channel Partner Code does not exist.";
                    }

                    qsAssignMatrix = new QueryExpression("systemuser");
                    qsAssignMatrix.ColumnSet = new ColumnSet("systemuserid");
                    condExp = new ConditionExpression("internalemailaddress", ConditionOperator.Equal, _ownerEmailId);
                    qsAssignMatrix.Criteria.AddCondition(condExp);
                    collct_assMatrix = service.RetrieveMultiple(qsAssignMatrix);
                    if (collct_assMatrix.Entities.Count >= 1)
                    {
                        erBSH = collct_assMatrix.Entities[0].ToEntityReference();
                    }
                    else
                    {
                        entity["hil_uploadstatus"] = false;
                        entity["hil_errordescription"] = "BSH EmailId does not exist.";
                    }

                    if (erDivision != null && erFranchise != null && erBSH != null)
                    {
                        if (actionTaken.Value == 3 || actionTaken.Value == 4) //Owner Update & Update Upcountry
                        {
                            qsAssignMatrix = new QueryExpression("hil_assignmentmatrix");
                            qsAssignMatrix.ColumnSet = new ColumnSet("hil_franchiseedirectengineer");
                            condExp = new ConditionExpression("hil_callsubtype", ConditionOperator.Equal, callSubType.Id);
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            condExp = new ConditionExpression("hil_division", ConditionOperator.Equal, erDivision.Id);
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            condExp = new ConditionExpression("hil_pincode", ConditionOperator.Equal, pinCode.Id);
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            condExp = new ConditionExpression("hil_franchiseedirectengineer", ConditionOperator.Equal, erFranchise.Id);
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            condExp = new ConditionExpression("statecode", ConditionOperator.Equal, 0); //Active Assignment Matrix Record
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            qsAssignMatrix.AddOrder("createdon", OrderType.Descending);
                            collct_assMatrix = service.RetrieveMultiple(qsAssignMatrix);
                            if (collct_assMatrix.Entities.Count >= 1)
                            {
                                if (actionTaken.Value == 4)
                                {
                                    // Update upcountry Flag (02/Aug/2020)
                                    Entity entUpdateUpcountry = new Entity("hil_assignmentmatrix", collct_assMatrix.Entities[0].Id);
                                    entUpdateUpcountry.Attributes["hil_upcountry"] = _upCountry;
                                    service.Update(entUpdateUpcountry);
                                }
                                else
                                {
                                    // Change BSH/ASH
                                    AssignRequest assign = new AssignRequest();
                                    assign.Assignee = new EntityReference("systemuser", erBSH.Id); //User or team
                                    assign.Target = new EntityReference("hil_assignmentmatrix", collct_assMatrix.Entities[0].Id); //Record to be assigned
                                    service.Execute(assign);
                                    entity["hil_uploadstatus"] = true;
                                    entity["hil_errordescription"] = "New Record created successfully.";
                                    entity["hil_assignmentmatrix"] = new EntityReference("hil_assignmentmatrix", collct_assMatrix.Entities[0].Id);
                                }
                            }
                            else
                            {
                                entity["hil_uploadstatus"] = true;
                                entity["hil_errordescription"] = "No Record found.";
                            }
                        }
                        else if (actionTaken.Value == 1)
                        {
                            qsAssignMatrix = new QueryExpression("hil_assignmentmatrix");
                            qsAssignMatrix.ColumnSet = new ColumnSet("hil_franchiseedirectengineer");
                            condExp = new ConditionExpression("hil_callsubtype", ConditionOperator.Equal, callSubType.Id);
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            condExp = new ConditionExpression("hil_division", ConditionOperator.Equal, erDivision.Id);
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            condExp = new ConditionExpression("hil_pincode", ConditionOperator.Equal, pinCode.Id);
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            condExp = new ConditionExpression("hil_franchiseedirectengineer", ConditionOperator.Equal, erFranchise.Id);
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            condExp = new ConditionExpression("statecode", ConditionOperator.Equal, 0); //Active Assignment Matrix Record
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            collct_assMatrix = service.RetrieveMultiple(qsAssignMatrix);
                            if (collct_assMatrix.Entities.Count >= 1)
                            {
                                foreach (Entity assignMatrix_record in collct_assMatrix.Entities)
                                {
                                    //Set Record status : Inactive
                                    SetStateRequest setStateRequest = new SetStateRequest()
                                    {
                                        EntityMoniker = new EntityReference
                                        {
                                            Id = assignMatrix_record.Id,
                                            LogicalName = "hil_assignmentmatrix",
                                        },
                                        State = new OptionSetValue(1), //Inactive
                                        Status = new OptionSetValue(2) //Inactive
                                    };
                                    service.Execute(setStateRequest);
                                }
                                entity["hil_uploadstatus"] = true;
                                entity["hil_errordescription"] = "Record Inactivated Successfully.";
                            }
                            else
                            {
                                entity["hil_uploadstatus"] = true;
                                entity["hil_errordescription"] = "No Action Required.";
                            }
                        }
                        else
                        {
                            qsAssignMatrix = new QueryExpression("hil_assignmentmatrix");
                            qsAssignMatrix.ColumnSet = new ColumnSet("hil_franchiseedirectengineer");
                            condExp = new ConditionExpression("hil_callsubtype", ConditionOperator.Equal, callSubType.Id);
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            condExp = new ConditionExpression("hil_division", ConditionOperator.Equal, erDivision.Id);
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            condExp = new ConditionExpression("hil_pincode", ConditionOperator.Equal, pinCode.Id);
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            condExp = new ConditionExpression("hil_franchiseedirectengineer", ConditionOperator.NotEqual, erFranchise.Id);
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            condExp = new ConditionExpression("statecode", ConditionOperator.Equal, 0); //Active Assignment Matrix Record
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            collct_assMatrix = service.RetrieveMultiple(qsAssignMatrix);
                            if (collct_assMatrix.Entities.Count >= 1)
                            {
                                foreach (Entity assignMatrix_record in collct_assMatrix.Entities)
                                {
                                    // Set Record status : Inactive
                                    SetStateRequest setStateRequest = new SetStateRequest()
                                    {
                                        EntityMoniker = new EntityReference
                                        {
                                            Id = assignMatrix_record.Id,
                                            LogicalName = "hil_assignmentmatrix",
                                        },
                                        State = new OptionSetValue(1), //Inactive
                                        Status = new OptionSetValue(2) //Inactive
                                    };
                                    service.Execute(setStateRequest);
                                }
                            }

                            qsAssignMatrix = new QueryExpression("hil_assignmentmatrix");
                            qsAssignMatrix.ColumnSet = new ColumnSet("hil_franchiseedirectengineer");
                            condExp = new ConditionExpression("hil_callsubtype", ConditionOperator.Equal, callSubType.Id);
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            condExp = new ConditionExpression("hil_division", ConditionOperator.Equal, erDivision.Id);
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            condExp = new ConditionExpression("hil_pincode", ConditionOperator.Equal, pinCode.Id);
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            condExp = new ConditionExpression("hil_franchiseedirectengineer", ConditionOperator.Equal, erFranchise.Id);
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            condExp = new ConditionExpression("statecode", ConditionOperator.Equal, 0); //Active Assignment Matrix Record
                            qsAssignMatrix.Criteria.AddCondition(condExp);
                            qsAssignMatrix.AddOrder("createdon", OrderType.Descending);

                            collct_assMatrix = service.RetrieveMultiple(qsAssignMatrix);
                            if (collct_assMatrix.Entities.Count > 0)
                            {
                                int count = 0;
                                foreach (Entity assignMatrix_record in collct_assMatrix.Entities)
                                {
                                    if (count == 0) { count += 1; continue; }
                                    SetStateRequest setStateRequest = new SetStateRequest()
                                    {
                                        EntityMoniker = new EntityReference
                                        {
                                            Id = assignMatrix_record.Id,
                                            LogicalName = "hil_assignmentmatrix",
                                        },
                                        State = new OptionSetValue(1), //Inactive
                                        Status = new OptionSetValue(2) //Inactive
                                    };
                                    service.Execute(setStateRequest);
                                    count += 1;
                                }
                                entity["hil_uploadstatus"] = true;
                                entity["hil_errordescription"] = "Allready Exists";
                            }
                            else
                            {
                                Entity obj_CreateAssMatrix = new Entity("hil_assignmentmatrix");
                                obj_CreateAssMatrix.Attributes["hil_pincode"] = new EntityReference("hil_pincode", pinCode.Id);
                                obj_CreateAssMatrix.Attributes["hil_division"] = new EntityReference("product", erDivision.Id);
                                obj_CreateAssMatrix.Attributes["hil_callsubtype"] = new EntityReference("hil_callsubtype", callSubType.Id);
                                obj_CreateAssMatrix.Attributes["hil_franchiseedirectengineer"] = new EntityReference("account", erFranchise.Id);
                                obj_CreateAssMatrix.Attributes["hil_upcountry"] = _upCountry;
                                _currentRecId = service.Create(obj_CreateAssMatrix);
                                AssignRequest assign = new AssignRequest();
                                assign.Assignee = new EntityReference("systemuser", erBSH.Id); //User or team
                                assign.Target = new EntityReference("hil_assignmentmatrix", _currentRecId); //Record to be assigned
                                service.Execute(assign);
                                entity["hil_uploadstatus"] = true;
                                entity["hil_errordescription"] = "New Record created successfully.";
                                entity["hil_assignmentmatrix"] = new EntityReference("hil_assignmentmatrix", _currentRecId);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                entity["hil_uploadstatus"] = false;
                entity["hil_errordescription"] = ex.Message;
            }
        }
    }
}
