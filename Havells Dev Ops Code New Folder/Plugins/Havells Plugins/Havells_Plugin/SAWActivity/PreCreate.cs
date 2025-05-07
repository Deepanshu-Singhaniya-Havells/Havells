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

namespace Havells_Plugin.SAWActivity
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
                    && context.PrimaryEntityName.ToLower() == "hil_sawactivity"
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("1");

                    if (entity.Attributes.Contains("hil_jobid"))
                    {
                        tracingService.Trace("2");
                        var job = ((EntityReference)(entity.Attributes["hil_jobid"]));
                        var erSAWCatg = ((EntityReference)(entity.Attributes["hil_sawcategory"]));

                        if (erSAWCatg.Id == new Guid("AD96A922-0AEE-EA11-A815-000D3AF057DD")) //One Time Labor Exception
                        {
                            QueryExpression qryExp = new QueryExpression(msdyn_workorder.EntityLogicalName);
                            qryExp.ColumnSet = new ColumnSet("msdyn_substatus", "hil_laborinwarranty");
                            qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                            qryExp.Criteria.AddCondition(new ConditionExpression("msdyn_workorderid", ConditionOperator.Equal, job.Id));
                            qryExp.NoLock = true;
                            EntityCollection entCol = service.RetrieveMultiple(qryExp);
                            if (entCol != null)
                            {
                                EntityReference entSubStatus = entCol.Entities[0].GetAttributeValue<EntityReference>("msdyn_substatus");
                                bool laborInWarranty = entCol.Entities[0].GetAttributeValue<bool>("hil_laborinwarranty");
                                if (entSubStatus.Id != new Guid("2927FA6C-FA0F-E911-A94E-000D3AF060A1") || laborInWarranty)
                                {
                                    throw new InvalidPluginExecutionException(" ***************“One time Labor Exception” can only be raised in case of Work Done & Out-Warranty Job.*************** ");
                                }
                            }
                        }

                        var entjob = service.Retrieve(job.LogicalName, job.Id, new ColumnSet("msdyn_name"));
                        string jobId = entjob["msdyn_name"].ToString();
                        int rowCount = 0;
                        tracingService.Trace(jobId);
                        QueryExpression Query = new QueryExpression("hil_sawactivity");
                        Query.ColumnSet = new ColumnSet("hil_name");
                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                        Query.Criteria.AddCondition(new ConditionExpression("hil_jobid", ConditionOperator.Equal, job.Id));
                        Query.NoLock = true;
                        EntityCollection Found = service.RetrieveMultiple(Query);
                        rowCount = Found.Entities.Count + 1;
                        tracingService.Trace(rowCount.ToString());
                        entity["hil_name"] = jobId + "-" + rowCount.ToString().PadLeft(3, '0');
                        tracingService.Trace("3");

                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.SAWActivity.PreCreate.Execute" + ex.Message);
            }
            #endregion
        }
    }
}
