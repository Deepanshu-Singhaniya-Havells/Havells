using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;

namespace HavellsNewPlugin.WorkOrder
{
    public class PostClaimsRejection : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {
                #region PluginConfig
                ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
                IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
                #endregion

                if (context.InputParameters.Contains("jobid") && context.InputParameters["jobid"] is string && context.InputParameters.Contains("remarks")
                    && context.InputParameters["remarks"] is string && context.InputParameters.Contains("operation") &&
                    context.InputParameters["operation"] is string && context.Depth == 1)
                {

                    Guid initiatingUserId = context.InitiatingUserId;
                    Guid jobid = new Guid(context.InputParameters["jobid"].ToString());
                    var remarks = context.InputParameters["remarks"].ToString();
                    var operation = context.InputParameters["operation"].ToString();
                    tracingService.Trace("Jobid " + jobid.ToString() + " remarks " + remarks + "  operation " + operation);
                    if (operation == "1")
                    {
                        string returnMsg = CliamRejection(jobid, service, remarks, initiatingUserId);
                        if (returnMsg == "1")
                        {
                            context.OutputParameters["issuccess"] = true;
                            context.OutputParameters["outremarks"] = "Claim Rejected Successfully";
                        }
                        else
                        {
                            context.OutputParameters["issuccess"] = false;
                            context.OutputParameters["outremarks"] = returnMsg;
                        }
                    }
                    else if (operation == "2")
                    {
                        //try
                        //{
                        //    tracingService.Trace("Jobid " + jobid.ToString() + " remarks " + remarks + "  operation " + operation);
                        //    CommonLib lib = new CommonLib();
                        //    CommonLib sawresponse = lib.CreateSAWActivity(jobid, 0, SAWCategoryConst._OneTimeLaborException, service, remarks, null);

                        //    if (sawresponse.statusRemarks == "OK")
                        //    {
                        //        string returnMsg = CliamRejectionSAW(jobid, service, remarks, initiatingUserId, SAWCategoryConst._OneTimeLaborException, 0);
                        //        context.OutputParameters["issuccess"] = true;
                        //        context.OutputParameters["outremarks"] = "OW to IW converted Successfully";
                        //    }
                        //    else
                        //    {
                        //        context.OutputParameters["issuccess"] = false;
                        //        context.OutputParameters["outremarks"] = sawresponse.statusRemarks;
                        //    }
                        //}
                        //catch (Exception ex)
                        //{
                        //    throw new InvalidPluginExecutionException
                        //}
                    }
                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error :- " + ex.Message);
            }
        }
        public string CliamRejectionSAW(Guid jobid, IOrganizationService service, string remarks, Guid initiatingUserId, string _sawCategory, int _amount)
        {
            string retmsg;


            try
            {
                QueryExpression qryExp = new QueryExpression("msdyn_workorder");
                qryExp.ColumnSet = new ColumnSet("hil_owneraccount", "ownerid");
                qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                qryExp.Criteria.AddCondition("msdyn_workorderid", ConditionOperator.Equal, jobid);
                EntityCollection entCol = service.RetrieveMultiple(qryExp);
                if (entCol.Entities.Count > 0)
                {
                    if (!CheckIfSAWActivityExist(jobid, service, _sawCategory))
                    {
                        Entity entSAWActivity = new Entity("hil_sawactivity");
                        entSAWActivity["hil_sawcategory"] = new EntityReference("hil_serviceactionwork", new Guid(_sawCategory));
                        entSAWActivity["hil_relatedchannelpartner"] = entCol.Entities[0].GetAttributeValue<EntityReference>("hil_owneraccount");
                        entSAWActivity["hil_jobid"] = new EntityReference("msdyn_workorder", jobid);
                        entSAWActivity["hil_description"] = remarks;

                        entSAWActivity["hil_amount"] = new Money(_amount);
                        entSAWActivity["hil_approvalstatus"] = new OptionSetValue(1); //requested
                        Guid sawActivityId = service.Create(entSAWActivity);
                        CreateSAWActivityApprovals(sawActivityId, service);
                    }
                }
                else
                {
                    retmsg = "job Not Found";
                }
                Entity wo = new Entity("msdyn_workorder");
                wo["hil_claimstatus"] = new OptionSetValue(3);
                wo.Id = jobid;
                service.Update(wo);

                QueryExpression _query = new QueryExpression("hil_jobsextension");
                _query.ColumnSet = new ColumnSet(false);
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                _query.Criteria.AddCondition(new ConditionExpression("hil_jobs", ConditionOperator.Equal, jobid));
                EntityCollection JobExtColl = service.RetrieveMultiple(_query);
                if (JobExtColl.Entities.Count > 0)
                {
                    Entity ent = new Entity(JobExtColl.EntityName);
                    ent.Id = JobExtColl[0].Id;

                    ent["hil_wocreatedby"] = new EntityReference("systemuser", initiatingUserId);
                    ent["hil_wocreatedon"] = DateTime.Now.AddMinutes(330);
                    service.Update(ent);
                    retmsg = "1";
                }
                else
                {
                    Entity ent = new Entity(JobExtColl.EntityName);

                    ent["hil_wocreatedby"] = new EntityReference("systemuser", initiatingUserId);
                    ent["hil_wocreatedon"] = DateTime.Now.AddMinutes(330);
                    ent["hil_jobs"] = new EntityReference("msdyn_workorder", jobid);
                    service.Create(ent);
                    retmsg = "1";
                }

                _query = new QueryExpression("hil_claimline");
                _query.ColumnSet = new ColumnSet(false);
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                _query.Criteria.AddCondition(new ConditionExpression("hil_jobid", ConditionOperator.Equal, jobid));
                EntityCollection claimsColl = service.RetrieveMultiple(_query);
                foreach (Entity calim in claimsColl.Entities)
                {
                    SetStateRequest setStateRequest = new SetStateRequest()
                    {
                        EntityMoniker = new EntityReference
                        {
                            Id = calim.Id,
                            LogicalName = calim.LogicalName,
                        },
                        State = new OptionSetValue(1),
                        Status = new OptionSetValue(2)
                    };
                    service.Execute(setStateRequest);
                }
            }
            catch (Exception ex)
            {
                retmsg = ex.Message.ToString();
            }
            return retmsg;
        }
        public void CreateSAWActivityApprovals(Guid _sawActivityId, IOrganizationService _service)
        {

            try
            {
                Entity ent = _service.Retrieve("hil_sawactivity", _sawActivityId, new ColumnSet("hil_sawcategory", "hil_jobid"));
                if (ent != null)
                {
                    EntityReference _salesOffice = null;
                    EntityReference _productCatg = null;
                    EntityReference _job = null;
                    EntityReference _picUser = null;
                    EntityReference _picPosition = null;
                    Entity entJob = _service.Retrieve("msdyn_workorder", ent.GetAttributeValue<EntityReference>("hil_jobid").Id, new ColumnSet("hil_salesoffice", "hil_productcategory"));
                    if (entJob != null)
                    {
                        if (entJob.Attributes.Contains("hil_salesoffice"))
                        {
                            _salesOffice = entJob.GetAttributeValue<EntityReference>("hil_salesoffice");
                            _productCatg = entJob.GetAttributeValue<EntityReference>("hil_productcategory");
                            _job = ent.GetAttributeValue<EntityReference>("hil_jobid");
                        }
                    }
                    if (_salesOffice != null && _productCatg != null)
                    {
                        QueryExpression qrySBU = new QueryExpression("hil_sbubranchmapping");
                        qrySBU.ColumnSet = new ColumnSet("hil_nsh", "hil_nph", "hil_branchheaduser");
                        qrySBU.Criteria = new FilterExpression(LogicalOperator.And);
                        qrySBU.Criteria.AddCondition("hil_salesoffice", ConditionOperator.Equal, _salesOffice.Id);
                        qrySBU.Criteria.AddCondition("hil_productdivision", ConditionOperator.Equal, _productCatg.Id);
                        EntityCollection entColSBU = _service.RetrieveMultiple(qrySBU);

                        QueryExpression qryExp = new QueryExpression("hil_sawcategoryapprovals");
                        qryExp.ColumnSet = new ColumnSet("hil_picuser", "hil_picposition", "hil_level");
                        qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qryExp.Criteria.AddCondition("hil_sawcategoryid", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_sawcategory").Id);
                        EntityCollection entCol = _service.RetrieveMultiple(qryExp);
                        if (entCol.Entities.Count > 0)
                        {
                            foreach (Entity entTemp in entCol.Entities)
                            {
                                if (entTemp.Attributes.Contains("hil_picuser"))
                                {
                                    _picUser = entTemp.GetAttributeValue<EntityReference>("hil_picuser");
                                }
                                else
                                {
                                    _picPosition = entTemp.GetAttributeValue<EntityReference>("hil_picposition");
                                    if (_picPosition.Name.ToUpper() == "BSH")
                                    {
                                        if (entColSBU.Entities.Count > 0)
                                        {
                                            _picUser = entColSBU.Entities[0].GetAttributeValue<EntityReference>("hil_branchheaduser");
                                        }
                                    }
                                    else if (_picPosition.Name.ToUpper() == "NSH")
                                    {
                                        _picUser = entColSBU.Entities[0].GetAttributeValue<EntityReference>("hil_nsh");

                                    }
                                    else if (_picPosition.Name.ToUpper() == "NPH")
                                    {
                                        _picUser = entColSBU.Entities[0].GetAttributeValue<EntityReference>("hil_nph");
                                    }
                                }
                                Entity entSAWActivity = new Entity("hil_sawactivityapproval");
                                entSAWActivity["hil_sawactivity"] = new EntityReference("hil_sawactivity", _sawActivityId);
                                entSAWActivity["hil_jobid"] = _job;
                                entSAWActivity["hil_level"] = entTemp.GetAttributeValue<OptionSetValue>("hil_level");
                                if (entTemp.GetAttributeValue<OptionSetValue>("hil_level").Value == 1)
                                {
                                    entSAWActivity["hil_isenabled"] = true;
                                }
                                entSAWActivity["hil_approver"] = _picUser;
                                entSAWActivity["hil_approvalstatus"] = new OptionSetValue(1); //requested
                                Guid sawActivityId = _service.Create(entSAWActivity);
                            }

                        }

                    }

                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error :- " + ex.Message);
            }
        }

        public bool CheckIfSAWActivityExist(Guid _jobId, IOrganizationService _service, string _sawCategory)
        {
            try
            {
                QueryExpression qryExp = new QueryExpression("hil_sawactivity");
                qryExp.ColumnSet = new ColumnSet("hil_jobid");
                qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                qryExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _jobId);
                qryExp.Criteria.AddCondition("hil_sawcategory", ConditionOperator.Equal, new Guid(_sawCategory));
                EntityCollection entCol = _service.RetrieveMultiple(qryExp);
                if (entCol.Entities.Count > 0) { return true; } else { return false; }
            }
            catch
            {
                return false;
            }
        }
        public string CliamRejection(Guid jobid, IOrganizationService service, string remarks, Guid initiatingUserId)
        {
            string retmsg;
            try
            {
                Entity wo = new Entity("msdyn_workorder");
                wo["hil_claimstatus"] = new OptionSetValue(3);
                wo.Id = jobid;
                service.Update(wo);

                QueryExpression _query = new QueryExpression("hil_jobsextension");
                _query.ColumnSet = new ColumnSet(false);
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                _query.Criteria.AddCondition(new ConditionExpression("hil_jobs", ConditionOperator.Equal, jobid));
                EntityCollection JobExtColl = service.RetrieveMultiple(_query);
                if (JobExtColl.Entities.Count > 0)
                {
                    Entity ent = new Entity(JobExtColl.EntityName);
                    ent.Id = JobExtColl[0].Id;
                    ent["hil_claimrejectionremarks"] = remarks;
                    ent["hil_claimejectedby"] = new EntityReference("systemuser", initiatingUserId);
                    ent["hil_claimrejectedon"] = DateTime.Now.AddMinutes(330);
                    service.Update(ent);
                    retmsg = "1";
                }
                else
                {
                    Entity ent = new Entity(JobExtColl.EntityName);
                    ent["hil_claimrejectionremarks"] = remarks;
                    ent["hil_claimejectedby"] = new EntityReference("systemuser", initiatingUserId);
                    ent["hil_claimrejectedon"] = DateTime.Now.AddMinutes(330);
                    ent["hil_jobs"] = new EntityReference("msdyn_workorder", jobid);
                    service.Create(ent);
                    retmsg = "1";
                }

                _query = new QueryExpression("hil_claimline");
                _query.ColumnSet = new ColumnSet(false);
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                _query.Criteria.AddCondition(new ConditionExpression("hil_jobid", ConditionOperator.Equal, jobid));
                EntityCollection claimsColl = service.RetrieveMultiple(_query);
                foreach (Entity calim in claimsColl.Entities)
                {
                    SetStateRequest setStateRequest = new SetStateRequest()
                    {
                        EntityMoniker = new EntityReference
                        {
                            Id = calim.Id,
                            LogicalName = calim.LogicalName,
                        },
                        State = new OptionSetValue(1),
                        Status = new OptionSetValue(2)
                    };
                    service.Execute(setStateRequest);
                }
            }
            catch (Exception ex)
            {
                retmsg = ex.Message.ToString();
            }
            return retmsg;
        }

    }
}
