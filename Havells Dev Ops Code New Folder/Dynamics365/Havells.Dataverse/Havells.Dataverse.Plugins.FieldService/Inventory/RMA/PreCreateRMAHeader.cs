using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Runtime.Remoting.Services;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreCreateRMAHeader : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                && context.PrimaryEntityName.ToLower() == "hil_inventoryrma" && (context.MessageName.ToUpper() == "CREATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    #region Checking for Open RMA for selected Return Type
                    if(!entity.Contains("hil_franchise"))
                        throw new InvalidPluginExecutionException("Franchise/DSE is required.");
                    if (!entity.Contains("hil_warehouse"))
                        throw new InvalidPluginExecutionException("Franchise/DSE Warehouse is required.");
                    if (!entity.Contains("hil_returntype"))
                        throw new InvalidPluginExecutionException("RMA Return Tyoe is required.");

                    EntityReference _franchise = entity.GetAttributeValue<EntityReference>("hil_franchise");
                    EntityReference _warehouse = entity.GetAttributeValue<EntityReference>("hil_warehouse");
                    EntityReference _returntype = entity.GetAttributeValue<EntityReference>("hil_returntype");

                    string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
                       <entity name='hil_inventoryrma'>
                         <attribute name='hil_inventoryrmaid' />
                         <attribute name='hil_name' />
                         <order attribute='hil_name' descending='false' />
                         <filter type='and'>
                           <condition attribute='statecode' operator='eq' value='0' />
                           <condition attribute='hil_franchise' operator='eq' value='{_franchise.Id}' />
                           <condition attribute='hil_warehouse' operator='eq' value='{_warehouse.Id}' />
                           <condition attribute='hil_returntype' operator='eq' value='{_returntype.Id}' />
                           <condition attribute='hil_inspectionnumber' operator='null' />
                         </filter>
                       </entity>
                     </fetch>";

                    EntityCollection result = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (result.Entities.Count > 0)
                    {
                        string _rmaNumber = result.Entities[0].GetAttributeValue<string>("hil_name");
                        throw new InvalidPluginExecutionException($"Open RMA is found with RMA Number: {_rmaNumber}. You are not allowed to create another RMA.");

                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
    }
}
