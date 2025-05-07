using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IdentityModel.Metadata;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.Dataverse.Plugins.FieldService.Inventory.RMA
{
    public class PreValidateRMA : IPlugin
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
                    string franchise = entity.Contains("hil_franchise") ? entity.GetAttributeValue<EntityReference>("hil_franchise").Name : string.Empty;
                    string warehouse = entity.Contains("hil_warehouse") ? entity.GetAttributeValue<EntityReference>("hil_warehouse").Name : string.Empty;
                    string returntype = entity.Contains("hil_returntype") ? entity.GetAttributeValue<EntityReference>("hil_returntype").Name : string.Empty;
                    OptionSetValue brandname = entity.Contains("hil_brandname") ? entity.GetAttributeValue<OptionSetValue>("hil_brandname") : null;

                    string _rmaFetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='hil_inventoryrma'>
                                        <attribute name='hil_inventoryrmaid'/>
                                        <attribute name='hil_name'/>
                                        <order attribute='hil_name' descending='false'/>
                                        <filter type='and'>
                                        <condition attribute='hil_franchise' operator='eq' value='{franchise}'/>
                                        <condition attribute='hil_warehouse' operator='eq'  value='{warehouse}'/>
                                        <condition attribute='hil_returntype' operator='eq' value='{returntype}'/>
                                        <condition attribute='hil_rmastatus' operator='not-in'>
                                        <value>4</value>
                                        <value>5</value>
                                        </condition>
                                        <condition attribute='hil_brandname' operator='eq' value='{brandname.Value}'/>
                                        <condition attribute='statecode' operator='eq' value='0'/>
                                        </filter>
                                        </entity>
                                        </fetch>";

                    EntityCollection entCollRMA = service.RetrieveMultiple(new FetchExpression(_rmaFetch));
                    if (entCollRMA.Entities.Count > 0)
                    {
                        {
                            throw new InvalidPluginExecutionException($"Access denied! Action RMA {entity.GetAttributeValue<string>("hil_name")} is found.");
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException("Havells.Dataverse.Plugins.FieldService.Inventory.RMA.PreValidateRMA" + ex.Message);
                }
            }
        }
    }
}
