using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.AMC_OmniChannel
{
    public class AssignToBSH : IPlugin
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
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.Depth==1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    AssignRecord(service, entity);
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private static void AssignRecord(IOrganizationService service, Entity entity)
        {
            try
            {
                Entity invoice = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_address", "hil_modelcode"));
                if (invoice.Contains("hil_address") && invoice.Contains("hil_modelcode"))
                {
                    EntityReference addressRef = invoice.GetAttributeValue<EntityReference>("hil_address");
                    Entity address = service.Retrieve(addressRef.LogicalName, addressRef.Id, new ColumnSet("hil_salesoffice"));

                    EntityReference modelRef = invoice.GetAttributeValue<EntityReference>("hil_modelcode");
                    Entity model = service.Retrieve(modelRef.LogicalName, modelRef.Id, new ColumnSet("hil_division"));

                    EntityReference _erDivisionRef = model.GetAttributeValue<EntityReference>("hil_division");
                    EntityReference _erSalesOfficeRef = address.GetAttributeValue<EntityReference>("hil_salesoffice");
                    if (_erDivisionRef != null && _erSalesOfficeRef != null)
                    {
                        QueryExpression query = new QueryExpression("hil_sbubranchmapping");
                        query.ColumnSet = new ColumnSet("hil_branchheaduser");
                        query.Criteria.AddCondition("hil_salesoffice", ConditionOperator.Equal, _erSalesOfficeRef.Id);
                        query.Criteria.AddCondition("hil_productdivision", ConditionOperator.Equal, _erDivisionRef.Id);
                        query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); // Active
                        EntityCollection tempColl = service.RetrieveMultiple(query);

                        if (tempColl.Entities.Count > 0)
                        {
                            EntityReference _erBSH = tempColl.Entities[0].GetAttributeValue<EntityReference>("hil_branchheaduser");
                            invoice["ownerid"] = _erBSH;
                            service.Update(invoice);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
