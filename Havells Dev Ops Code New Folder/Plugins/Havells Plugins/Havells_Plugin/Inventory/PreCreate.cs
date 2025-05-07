using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.Inventory
{
    public class PreCreate : IPlugin
    {
        private static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == hil_inventory.EntityLogicalName && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("1 - Create Inventory");
                    hil_inventory _enInv = entity.ToEntity<hil_inventory>();
                    string UqKey = string.Empty;
                    if(_enInv.hil_Part != null && _enInv.hil_OwnerAccount != null)
                    {
                        tracingService.Trace("1.1 - Create Inventory");
                        QueryExpression Query = new QueryExpression(hil_inventory.EntityLogicalName);
                        Query.ColumnSet = new ColumnSet(false);
                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                        tracingService.Trace("1.2 - Create Inventory - "+ _enInv.hil_Part.Id.ToString());
                        Query.Criteria.AddCondition("hil_part", ConditionOperator.Equal, _enInv.hil_Part.Id);
                        tracingService.Trace("1.3 - Create Inventory - " + _enInv.hil_OwnerAccount.Id.ToString());
                        Query.Criteria.AddCondition("hil_owneraccount", ConditionOperator.Equal, _enInv.hil_OwnerAccount.Id);
                        tracingService.Trace("1.4 - Create Inventory - " + _enInv.OwnerId.Id.ToString());
                        Query.Criteria.AddCondition("ownerid", ConditionOperator.Equal, _enInv.OwnerId.Id);
                        EntityCollection Found = service.RetrieveMultiple(Query);
                        if (Found.Entities.Count > 0)
                        {
                            throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXX - SIMILAR INVENTORY ROW ALREADY EXISTS - XXXXXXXXXXXXXXXXXXXXXXXXXX");
                        }
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXX - PART AND ACCOUNT CAN'T BE EMPTY - XXXXXXXXXXXXXXXXXXXXXXXXXX");
                    }
                }
            }
            catch(Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Inventory.PreCreate.Execute : " +ex.Message.ToUpper());
            }
        }
    }
}
