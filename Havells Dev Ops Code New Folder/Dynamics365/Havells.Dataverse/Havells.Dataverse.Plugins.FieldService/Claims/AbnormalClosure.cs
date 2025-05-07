using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.Dataverse.Plugins.FieldService.Claims
{
    public class AbnormalClosure : IPlugin
    {
        private ITracingService tracingService;
        private IOrganizationService service;
        private Guid ClosedJob = new Guid("1727fa6c-fa0f-e911-a94e-000d3af060a1");

        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("JobId") && context.InputParameters["JobId"] is string
                        && context.Depth == 1)
                {
                    string jobId = (string)context.InputParameters["JobId"];

                    if (string.IsNullOrEmpty(jobId))
                    {
                        context.OutputParameters["Status"] = false;
                        context.OutputParameters["Message"] = "Job Id Cannot be empty";
                        return;
                    }

                    context.OutputParameters["Message"] = Action(new Guid(jobId));
                    context.OutputParameters["Status"] = true;

                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["Status"] = false;
                context.OutputParameters["Message"] = "D365 Internal Server Error : " + ex.Message;
            }
        }

        private string Action(Guid JobId)
        {

            Data data = new Data();
            data.Job = service.Retrieve("msdyn_workorder", JobId, new ColumnSet("hil_productcategory", "hil_callsubtype", "hil_owneraccount", "hil_salesoffice", "hil_brand", "hil_fiscalmonth", "hil_fiscalmonth", "msdyn_substatus", "createdon"));
            data.WorkOrderSubStatus = data.Job.GetAttributeValue<EntityReference>("msdyn_substatus");
            data.ClaimPeriod = data.Job.GetAttributeValue<EntityReference>("hil_fiscalmonth");
            data.ProductCategory = data.Job.GetAttributeValue<EntityReference>("hil_productcategory");
            data.Franchise = data.Job.GetAttributeValue<EntityReference>("hil_owneraccount");
            data.SalesOffice = data.Job.GetAttributeValue<EntityReference>("hil_salesoffice");
            data.Brand = data.Job.GetAttributeValue<OptionSetValue>("hil_brand").Value;
            data.CallSubType = data.Job.GetAttributeValue<EntityReference>("hil_callsubtype");
            data.ClaimPeriod = data.Job.GetAttributeValue<EntityReference>("hil_fiscalmonth");
            data.JobCreationDate = data.Job.GetAttributeValue<DateTime>("createdon").AddHours(5.5);


            QueryExpression query = new QueryExpression("hil_claimheader");
            query.ColumnSet = new ColumnSet("hil_performastatus");
            query.Criteria.AddCondition("hil_franchisee", ConditionOperator.Equal, data.Franchise.Id);
            query.Criteria.AddCondition("hil_fiscalmonth", ConditionOperator.Equal, data.ClaimPeriod.Id);
            EntityCollection claimHeaderCollection = service.RetrieveMultiple(query);

            if (claimHeaderCollection.Entities.Count > 0)
            {
                data.ClaimHeader = claimHeaderCollection.Entities[0].ToEntityReference();
                int status = claimHeaderCollection.Entities[0].Contains("hil_performastatus") ? claimHeaderCollection.Entities[0].GetAttributeValue<OptionSetValue>("hil_performastatus").Value : -1;

                // from Job : brand, sales office, product category, call subtype 

                if (data.WorkOrderSubStatus.Id == ClosedJob)
                {
                    if (status == 1 || status == 2) // 
                    {
                        query = new QueryExpression("hil_wrongclosurepenalty");
                        query.ColumnSet = new ColumnSet("hil_amount");
                        query.Criteria.AddCondition("hil_brand", ConditionOperator.Equal, data.Brand);
                        query.Criteria.AddCondition("hil_wrongclosurecategory", ConditionOperator.Equal, 2); //Abnormal Closure
                        query.Criteria.AddCondition("hil_fromdate", ConditionOperator.LessThan, data.JobCreationDate);
                        query.Criteria.AddCondition("hil_todate", ConditionOperator.GreaterThan, data.JobCreationDate);
                        query.Criteria.AddCondition("hil_amount", ConditionOperator.NotNull);
                        query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //Active

                        EntityCollection penaltyCollection = service.RetrieveMultiple(query);

                        if (penaltyCollection.Entities.Count > 0)
                        {
                            data.ClaimAmount = penaltyCollection.Entities[0].Contains("hil_amount") ? penaltyCollection.Entities[0].GetAttributeValue<Money>("hil_amount").Value : -8759437;

                            if (data.ClaimAmount != -8759437)
                            {
                                return CreateClaimLine(data);
                            }
                            else
                            {
                                return "Claim Amount is not present";
                            }
                        }
                        else
                        {
                            return "No corresponding warranty close rrecod exist in the system";
                        }
                    }
                    else
                    {
                        return "Claim Line is alrady approved";
                    }

                }
                else
                {
                    return "Job Should be in closed state";
                }




            }
            else
            {
                return "Corresponding Claim Header is not present";
            }

        }

        private string CreateClaimLine(Data data)
        {
            QueryExpression query = new QueryExpression("hil_claimline");
            query.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, data.Job.Id);
            query.Criteria.AddCondition("hil_claimheader", ConditionOperator.Equal, data.ClaimHeader.Id);
            query.Criteria.AddCondition("hil_claimperiod", ConditionOperator.Equal, data.ClaimPeriod.Id);
            query.Criteria.AddCondition("hil_claimcategory", ConditionOperator.Equal, new Guid("8dbfb0c4-33f3-ea11-a815-000d3af05a4b"));
            query.Criteria.AddCondition("hil_claimamount", ConditionOperator.Equal, -data.ClaimAmount);
            query.Criteria.AddCondition("hil_productcategory", ConditionOperator.Equal, data.ProductCategory.Id);
            query.Criteria.AddCondition("hil_callsubtype", ConditionOperator.Equal, data.CallSubType.Id);
            query.Criteria.AddCondition("hil_franchisee", ConditionOperator.Equal, data.Franchise.Id);
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); //active

            EntityCollection existingClaimLines = service.RetrieveMultiple(query);
            if (existingClaimLines.Entities.Count == 0)
            {
                Entity newClaimLine = new Entity("hil_claimline");
                newClaimLine["hil_jobid"] = data.Job.ToEntityReference();
                newClaimLine["hil_claimheader"] = data.ClaimHeader;
                newClaimLine["hil_claimperiod"] = data.ClaimPeriod;
                newClaimLine["hil_claimcategory"] = new EntityReference("hil_claimcategory", new Guid("8dbfb0c4-33f3-ea11-a815-000d3af05a4b"));
                newClaimLine["hil_claimamount"] = new Money(-data.ClaimAmount);
                newClaimLine["hil_productcategory"] = data.ProductCategory;
                newClaimLine["hil_callsubtype"] = data.CallSubType;
                newClaimLine["hil_franchisee"] = data.Franchise;
                service.Create(newClaimLine);

                return "Claim Line Created Successfully";
            }
            else
            {
                return "Claim line Already Created";
            }

        }

        class Data
        {
            public Entity Job { get; set; }
            public EntityReference SalesOffice { get; set; }
            public int Brand { get; set; }
            public EntityReference ProductCategory { get; set; }
            public EntityReference CallSubType { get; set; }
            public EntityReference Franchise { get; set; }
            public EntityReference ClaimPeriod { get; set; }
            public EntityReference ClaimHeader { get; set; }
            public EntityReference ClaimCategory { get; set; }
            public EntityReference WorkOrderSubStatus { get; set; }
            public decimal ClaimAmount { get; set; }
            public DateTime JobCreationDate { get; set; }
        }
    }
}
