using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.ServiceModel;

namespace Havells_Plugin.WorkOrderIncident
{
    public class PostCreateAsync : IPlugin
    {
        static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == msdyn_workorderincident.EntityLogicalName
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("-1");
                    msdyn_workorderincident _entWrkInc = entity.ToEntity<msdyn_workorderincident>();
                    tracingService.Trace("0");
                    Guid Model = Common.SetModeldetail(entity, service);
                    tracingService.Trace("1 " + Model.ToString());
                    if (Model != Guid.Empty)
                    {
                        tracingService.Trace("2 " + Model.ToString());
                        PopulateCauseActions(service, _entWrkInc, context);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrderIncident.PostCreateAsync" + ex.Message);
            }
        }
        #region Populate Cause Actions
        public static void PopulateCauseActions(IOrganizationService service, msdyn_workorderincident _inc, IPluginExecutionContext context)
        {
            msdyn_workorderservice iWServ = new msdyn_workorderservice();
            QueryExpression Query = new QueryExpression(msdyn_incidenttypeservice.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("msdyn_service");
            Query.Criteria.AddFilter(LogicalOperator.And);
            Query.Criteria.AddCondition("msdyn_incidenttype", ConditionOperator.Equal, _inc.msdyn_IncidentType.Id);
            Query.Criteria.AddCondition("msdyn_service", ConditionOperator.NotNull);
            EntityCollection Found = service.RetrieveMultiple(Query);
            tracingService.Trace("3 " + _inc.Id.ToString());
            if (Found.Entities.Count > 0)
            {
                tracingService.Trace("4 Cause Service Count: " + Found.Entities.Count.ToString());
                foreach (msdyn_incidenttypeservice iType in Found.Entities)
                {
                    iWServ = new msdyn_workorderservice();
                    tracingService.Trace("5 " + iType.msdyn_Service.Id.ToString());
                    iWServ.msdyn_Service = iType.msdyn_Service;
                    tracingService.Trace("6 " + _inc.msdyn_WorkOrder.Id.ToString());
                    iWServ.msdyn_WorkOrder = _inc.msdyn_WorkOrder;
                    tracingService.Trace("7 " + context.Depth.ToString());
                    iWServ.msdyn_Description = context.Depth.ToString();
                    tracingService.Trace("8 " + _inc.Id.ToString());
                    iWServ.msdyn_name = iType.msdyn_Service.Name;
                    tracingService.Trace("8.1 " + iWServ.msdyn_name);
                    iWServ.msdyn_WorkOrderIncident = new EntityReference(msdyn_workorderincident.EntityLogicalName, _inc.Id);
                    tracingService.Trace("9 " + _inc.msdyn_CustomerAsset.Id.ToString());
                    iWServ.msdyn_CustomerAsset = _inc.msdyn_CustomerAsset;
                    tracingService.Trace("10 " + _inc.Id.ToString());
                    try
                    {
                        service.Create(iWServ);
                    }
                    catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
                    {
                        tracingService.Trace(ex.Detail.Message + ":" + ex.Detail.InnerFault.Message);
                    }
                    tracingService.Trace("11 - Creation Done");
                }
            }
        }
        #endregion
    }
}
