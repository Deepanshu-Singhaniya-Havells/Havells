using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
// Microsoft Dynamics CRM namespace(s)
using System.Net;
using System.Net.Http;
using System.Runtime.Serialization.Json;
using System.IO;
using Havells_Plugin.HelperIntegration;
using System.Linq;

using System.Text;
using Havells_Plugin;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;

namespace Havells_Plugin.ClaimLine
{
  public class PreCreate : IPlugin
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
                    && context.PrimaryEntityName.ToLower() == "hil_claimline"
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("1");
                    //PopulateData(entity, service);
                    RestictDuplicateClaimLine(service, entity);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ClaimLine.PreCreate.Execute" + ex.Message);
            }
            #endregion
        }

        public static Int32 GetMobileAppClosurePenalt(IOrganizationService service)
        {
            try
            {
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    Int32 Value = 0;
                    QueryExpression qe = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
                    qe.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "ClaimJobMobileAppClosureIncentive");
                    qe.Criteria.AddCondition("hil_mobileappclosure", ConditionOperator.Equal, 1);
                    EntityCollection enColl = service.RetrieveMultiple(qe);
                    foreach (Entity en in enColl.Entities)
                    {
                        if (en.Contains("hil_priceforincentive"))
                        {
                            Value = en.GetAttributeValue<Int32>("hil_priceforincentive");
                        }
                    }
                    return Value;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ClaimLine.PreCreate.PopulateData" + ex.Message);
            }
        }
        protected static void RestictDuplicateClaimLine(IOrganizationService service, Entity claimline)
        {
            if (claimline.Contains("hil_jobid"))
            {
                EntityReference job = claimline.GetAttributeValue<EntityReference>("hil_jobid");
                EntityReference hil_claimperiod = claimline.GetAttributeValue<EntityReference>("hil_claimperiod");
                QueryExpression qryExp = new QueryExpression("hil_claimline");
                qryExp.ColumnSet = new ColumnSet(false);
                qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                qryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                qryExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, job.Id);
                qryExp.Criteria.AddCondition("hil_claimperiod", ConditionOperator.NotEqual, hil_claimperiod.Id);
                EntityCollection entColClaim = service.RetrieveMultiple(qryExp);
                if (entColClaim.Entities.Count > 0)
                {
                    throw new InvalidPluginExecutionException("Job already Closed!!! Duplicate Claim Lines are not allowed.");
                }
               
            }
            if (claimline.Contains("hil_jobid") && claimline.Contains("hil_claimcategory"))
            {
                EntityReference job = claimline.GetAttributeValue<EntityReference>("hil_jobid");
                EntityReference hil_claimcategory = claimline.GetAttributeValue<EntityReference>("hil_claimcategory");

                QueryExpression qryExp = new QueryExpression("hil_claimline");
                qryExp.ColumnSet = new ColumnSet(false);
                qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                qryExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, job.Id);
                qryExp.Criteria.AddCondition("hil_claimcategory", ConditionOperator.Equal, hil_claimcategory.Id);
                qryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                EntityCollection entColClaim = service.RetrieveMultiple(qryExp);
                if (entColClaim.Entities.Count > 0)
                {
                    throw new InvalidPluginExecutionException("Duplicate Claim Lines are not allowed.");
                }

            }
            if (claimline.Contains("hil_jobid") && claimline.Contains("hil_franchisee") && claimline.Contains("hil_claimcategory"))
            {
                EntityReference job = claimline.GetAttributeValue<EntityReference>("hil_jobid");
                EntityReference hil_franchisee = claimline.GetAttributeValue<EntityReference>("hil_franchisee");
                EntityReference hil_claimcategory = claimline.GetAttributeValue<EntityReference>("hil_claimcategory");

                if (hil_claimcategory.Id == new Guid("884A1BCC-6EE5-EA11-A817-000D3AF0501C")) // Upcountry Travellines
                {
                    Entity _entJob = service.Retrieve("msdyn_workorder", job.Id, new ColumnSet("hil_pincode"));
                    Entity _entChannelPartner = service.Retrieve("account", hil_franchisee.Id, new ColumnSet("hil_pincode"));
                    if (_entJob.GetAttributeValue<EntityReference>("hil_pincode").Id == _entChannelPartner.GetAttributeValue<EntityReference>("hil_pincode").Id)
                    {
                        throw new InvalidPluginExecutionException("Invalid Upcountry Claim Lines are not allowed.");
                    }
                }
            }
            if (claimline.Contains("hil_jobid")) {
                UpdateActivityCodeOnClaimLine(service, claimline.GetAttributeValue<EntityReference>("hil_jobid").Id, claimline);
            }
        }

        public static void UpdateActivityCodeOnClaimLine(IOrganizationService _service,Guid  _workOrderId, Entity entity)
        {
            #region Variable declaration
            QueryExpression queryExp;
            EntityCollection entcoll;
            Guid _performaInvoiceId = Guid.Empty;
            #endregion
            
            LinkEntity lnkEntSO = new LinkEntity
            {
                LinkFromEntityName = "account",
                LinkToEntityName = "hil_salesoffice",
                LinkFromAttributeName = "hil_salesoffice",
                LinkToAttributeName = "hil_salesofficeid",
                Columns = new ColumnSet("hil_state"),
                EntityAlias = "so",
                JoinOperator = JoinOperator.Inner
            };

            LinkEntity lnkEntCP = new LinkEntity
            {
                LinkFromEntityName = "msdyn_workorder",
                LinkToEntityName = "account",
                LinkFromAttributeName = "hil_owneraccount",
                LinkToAttributeName = "accountid",
                Columns = new ColumnSet("hil_state", "hil_salesoffice", "ownerid"),
                EntityAlias = "cp",
                JoinOperator = JoinOperator.Inner
            };

            lnkEntCP.LinkEntities.Add(lnkEntSO);

            queryExp = new QueryExpression("msdyn_workorder");
            queryExp.ColumnSet = new ColumnSet("hil_productcategory", "hil_callsubtype", "hil_owneraccount", "hil_warrantysubstatus");
            queryExp.Criteria = new FilterExpression(LogicalOperator.And);
            queryExp.Criteria.AddCondition("msdyn_workorderid", ConditionOperator.Equal, _workOrderId);
            queryExp.LinkEntities.Add(lnkEntCP);
            entcoll = _service.RetrieveMultiple(queryExp);
            if (entcoll.Entities.Count > 0)
            {
                int _gstCatg;
                QueryExpression queryExpTemp;
                EntityCollection entcollTemp;
                string _activityCode = string.Empty;
                foreach (Entity ent in entcoll.Entities)
                {
                    _activityCode = string.Empty;
                    if (ent.Contains("cp.hil_state") && ent.Contains("so.hil_state") && ent.Contains("hil_productcategory") && ent.Contains("hil_callsubtype") && ent.Contains("hil_warrantysubstatus"))
                    {
                        EntityReference _erCPState = ((EntityReference)ent.GetAttributeValue<AliasedValue>("cp.hil_state").Value);
                        EntityReference _erSOState = ((EntityReference)ent.GetAttributeValue<AliasedValue>("so.hil_state").Value);
                        EntityReference _erProdCategory = ent.GetAttributeValue<EntityReference>("hil_productcategory");
                        EntityReference _erCallSubtype = ent.GetAttributeValue<EntityReference>("hil_callsubtype");
                        OptionSetValue _osWarrantySubstatus = ent.GetAttributeValue<OptionSetValue>("hil_warrantysubstatus");

                        _gstCatg = _erCPState.Id == _erSOState.Id ? 1 : 2;

                        queryExpTemp = new QueryExpression("hil_claimpostingsetup");
                        queryExpTemp.ColumnSet = new ColumnSet("hil_activitycode");
                        queryExpTemp.Criteria = new FilterExpression(LogicalOperator.And);
                        queryExpTemp.Criteria.AddCondition("hil_callsubtype", ConditionOperator.Equal, _erCallSubtype.Id); //Breakdown
                        queryExpTemp.Criteria.AddCondition("hil_productcategory", ConditionOperator.Equal, _erProdCategory.Id);
                        queryExpTemp.Criteria.AddCondition("hil_activitygstslab", ConditionOperator.Equal, _gstCatg);
                        queryExpTemp.Criteria.AddCondition("hil_warrantysubstatus", ConditionOperator.Equal, _osWarrantySubstatus.Value); //Standard | Under AMC
                        queryExpTemp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        entcollTemp = _service.RetrieveMultiple(queryExpTemp);
                        if (entcollTemp.Entities.Count > 0)
                        {
                            _activityCode = entcollTemp.Entities[0].GetAttributeValue<string>("hil_activitycode");
                            entity["hil_callsubtype"] = _erCallSubtype;
                            entity["hil_productcategory"] = _erProdCategory;
                            entity["hil_activitycode"] = _activityCode;
                        }
                        else
                        {
                            entity["hil_callsubtype"] = _erCallSubtype;
                            entity["hil_productcategory"] = _erProdCategory;
                        }
                    }
                }
            }
        }
    }
}
