using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System.Net;
using System.Runtime.Serialization.Json;
using System.IO;
using Havells_Plugin.HelperIntegration;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk.Query;
using Havells_Plugin.SAWActivity;

namespace Havells_Plugin.SAWActivityApproval
{
    public class PostUpdate : IPlugin
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

            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "hil_sawactivityapproval"
                    && context.MessageName.ToUpper() == "UPDATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];

                    Entity ent = service.Retrieve("hil_sawactivityapproval", entity.Id, new ColumnSet("hil_jobid", "hil_sawactivity", "hil_approvalstatus", "hil_level"));
                    string _fetchXML = string.Empty;
                    QueryExpression qrExp;
                    Entity sawActivity;
                    bool _KKGAuditFailedSAW = false;
                    bool _oneTimeLaborExceptionSAW = false;
                    if (ent != null)
                    {
                        EntityReference erJobId = ent.GetAttributeValue<EntityReference>("hil_jobid");
                        EntityReference erSAWActivityId = ent.GetAttributeValue<EntityReference>("hil_sawactivity");
                        OptionSetValue approvalStatus = ent.GetAttributeValue<OptionSetValue>("hil_approvalstatus");
                        OptionSetValue approvalLevel = ent.GetAttributeValue<OptionSetValue>("hil_level");
                        EntityCollection entCol;
                        bool _relatedToClaim = false;
                        Entity entSAWActivity = service.Retrieve("hil_sawactivity", erSAWActivityId.Id, new ColumnSet("hil_sawcategory"));
                        if (entSAWActivity != null)
                        {
                            Entity entSAWCatg = service.Retrieve("hil_serviceactionwork", entSAWActivity.GetAttributeValue<EntityReference>("hil_sawcategory").Id, new ColumnSet("hil_relatedtoclaim"));
                            if (entSAWCatg.Attributes.Contains("hil_relatedtoclaim"))
                            {
                                _relatedToClaim = entSAWCatg.GetAttributeValue<bool>("hil_relatedtoclaim");
                            }
                            if (entSAWActivity.GetAttributeValue<EntityReference>("hil_sawcategory").Id == new Guid(SAWCategoryConst._KKGFailureReview))
                            {
                                _KKGAuditFailedSAW = true;
                            }
                            else if (entSAWActivity.GetAttributeValue<EntityReference>("hil_sawcategory").Id == new Guid(SAWCategoryConst._OneTimeLaborException))
                            {
                                _oneTimeLaborExceptionSAW = true;
                            }
                        }
                        if (_KKGAuditFailedSAW)
                        {
                            if (approvalStatus.Value == 4)
                            {
                                sawActivity = new Entity("hil_sawactivity");
                                sawActivity.Id = erSAWActivityId.Id;
                                sawActivity["hil_approvalstatus"] = new OptionSetValue(4); //Rejected
                                service.Update(sawActivity);

                                Entity Ticket = new Entity("msdyn_workorder");
                                Ticket.Id = erJobId.Id;
                                Ticket["hil_claimstatus"] = new OptionSetValue(7); //Abnormal Penalty Imposed
                                service.Update(Ticket);
                            }
                            else if (approvalStatus.Value == 3)
                            {
                                sawActivity = new Entity("hil_sawactivity");
                                sawActivity.Id = erSAWActivityId.Id;
                                sawActivity["hil_approvalstatus"] = new OptionSetValue(3); //Approved
                                service.Update(sawActivity);

                                Entity Ticket = new Entity("msdyn_workorder");
                                Ticket.Id = erJobId.Id;
                                Ticket["hil_claimstatus"] = new OptionSetValue(8); //Abnormal Penalty waived
                                service.Update(Ticket);

                            }
                        }
                        else
                        {
                            if (approvalStatus.Value == 4)
                            { //Rejected
                                if (approvalLevel.Value != 3)
                                {
                                    qrExp = new QueryExpression("hil_sawactivityapproval");
                                    qrExp.ColumnSet = new ColumnSet("hil_sawactivityapprovalid", "hil_approvalstatus");
                                    qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                                    qrExp.Criteria.AddCondition("hil_sawactivity", ConditionOperator.Equal, erSAWActivityId.Id);
                                    qrExp.Criteria.AddCondition("hil_sawactivityapprovalid", ConditionOperator.NotEqual, ent.Id);
                                    if (approvalLevel.Value == 1)
                                    {
                                        qrExp.Criteria.AddCondition("hil_level", ConditionOperator.In, new object[] { 2, 3 });
                                    }
                                    else if (approvalLevel.Value == 2)
                                    {
                                        qrExp.Criteria.AddCondition("hil_level", ConditionOperator.In, new object[] { 3 });
                                    }
                                    entCol = service.RetrieveMultiple(qrExp);
                                    if (entCol.Entities.Count > 0)
                                    {
                                        foreach (Entity entObj in entCol.Entities)
                                        {
                                            entObj["hil_approvalstatus"] = new OptionSetValue(6); //Rejected at Previous Level
                                            entObj["hil_approvedate"] = DateTime.Now;
                                            service.Update(entObj);
                                        }
                                    }
                                }
                                sawActivity = new Entity("hil_sawactivity");
                                sawActivity.Id = erSAWActivityId.Id;
                                sawActivity["hil_approvalstatus"] = new OptionSetValue(4); //Rejected
                                service.Update(sawActivity);
                            }
                            else if (approvalStatus.Value == 3)
                            { //Approved
                                qrExp = new QueryExpression("hil_sawactivityapproval");
                                qrExp.ColumnSet = new ColumnSet("hil_sawactivityapprovalid", "hil_approvalstatus");
                                qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                                qrExp.Criteria.AddCondition("hil_sawactivity", ConditionOperator.Equal, erSAWActivityId.Id);
                                qrExp.Criteria.AddCondition("hil_approvalstatus", ConditionOperator.NotEqual, 3);
                                entCol = service.RetrieveMultiple(qrExp);
                                if (entCol.Entities.Count == 0)
                                {
                                    sawActivity = new Entity("hil_sawactivity");
                                    sawActivity.Id = erSAWActivityId.Id;
                                    sawActivity["hil_approvalstatus"] = new OptionSetValue(3); //Approved
                                    service.Update(sawActivity);
                                }
                                if (approvalLevel.Value == 1)
                                {
                                    qrExp = new QueryExpression("hil_sawactivityapproval");
                                    qrExp.ColumnSet = new ColumnSet("hil_sawactivityapprovalid", "hil_approvalstatus");
                                    qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                                    qrExp.Criteria.AddCondition("hil_sawactivity", ConditionOperator.Equal, erSAWActivityId.Id);
                                    qrExp.Criteria.AddCondition("hil_level", ConditionOperator.Equal, 2);
                                    entCol = service.RetrieveMultiple(qrExp);
                                    if (entCol.Entities.Count > 0)
                                    {
                                        sawActivity = new Entity("hil_sawactivityapproval");
                                        sawActivity.Id = entCol.Entities[0].Id;
                                        sawActivity["hil_isenabled"] = true;
                                        service.Update(sawActivity);
                                    }
                                }
                                else if (approvalLevel.Value == 2)
                                {
                                    qrExp = new QueryExpression("hil_sawactivityapproval");
                                    qrExp.ColumnSet = new ColumnSet("hil_sawactivityapprovalid", "hil_approvalstatus");
                                    qrExp.Criteria = new FilterExpression(LogicalOperator.And);
                                    qrExp.Criteria.AddCondition("hil_sawactivity", ConditionOperator.Equal, erSAWActivityId.Id);
                                    qrExp.Criteria.AddCondition("hil_level", ConditionOperator.Equal, 3);
                                    entCol = service.RetrieveMultiple(qrExp);
                                    if (entCol.Entities.Count > 0)
                                    {
                                        sawActivity = new Entity("hil_sawactivityapproval");
                                        sawActivity.Id = entCol.Entities[0].Id;
                                        sawActivity["hil_isenabled"] = true;
                                        service.Update(sawActivity);
                                    }
                                }
                            }

                            if (_relatedToClaim)
                            {
                                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_sawactivity'>
                                <attribute name='hil_sawactivityid' />
                                <filter type='and'>
                                    <condition attribute='hil_jobid' operator='eq' value='" + erJobId.Id + @"' />
                                    <condition attribute='hil_approvalstatus' operator='eq' value='4' />
                                </filter>
                                <link-entity name='hil_serviceactionwork' from='hil_serviceactionworkid' to='hil_sawcategory' link-type='inner' alias='ad'>
                                <filter type='and'>
                                    <condition attribute='hil_relatedtoclaim' operator='eq' value='1' />
                                </filter>
                                </link-entity>
                                </entity>
                                </fetch>";
                                entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                if (entCol.Entities.Count > 0)
                                {
                                    Entity Ticket = new Entity("msdyn_workorder");
                                    Ticket.Id = erJobId.Id;
                                    Ticket["hil_claimstatus"] = new OptionSetValue(3); //Claim Rejected
                                    service.Update(Ticket);
                                }
                                else
                                {
                                    _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_sawactivity'>
                                <attribute name='hil_sawactivityid' />
                                <filter type='and'>
                                    <condition attribute='hil_jobid' operator='eq' value='" + erJobId.Id + @"' />
                                    <condition attribute='hil_approvalstatus' operator='ne' value='3' />
                                </filter>
                                <link-entity name='hil_serviceactionwork' from='hil_serviceactionworkid' to='hil_sawcategory' link-type='inner' alias='ad'>
                                <filter type='and'>
                                    <condition attribute='hil_relatedtoclaim' operator='eq' value='1' />
                                </filter>
                                </link-entity>
                                </entity>
                                </fetch>";
                                    entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                    if (entCol.Entities.Count == 0)
                                    {
                                        Entity Ticket = new Entity("msdyn_workorder");
                                        Ticket.Id = erJobId.Id;
                                        Ticket["hil_claimstatus"] = new OptionSetValue(4); //Claim Approved
                                        service.Update(Ticket);
                                    }
                                }
                            }
                            else
                            {
                                if (_oneTimeLaborExceptionSAW)
                                {
                                    Entity Ticket = new Entity("msdyn_workorder");
                                    Ticket.Id = erJobId.Id;
                                    Ticket["hil_laborinwarranty"] = true; //Labor In Warranty 
                                    service.Update(Ticket);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.SAWActivityApproval.PreCreate.Execute" + ex.Message);
            }
            #endregion
        }
    }
}
