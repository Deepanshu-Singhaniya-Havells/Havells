using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreUpdateInventoryAdjustment_SubmitForApprovalValidation : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "hil_inventoryspareadjustment" && (context.MessageName.ToUpper() == "UPDATE"))
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                if (entity.Contains("hil_adjustmentstatus"))
                {
                    OptionSetValue _adjStatus = entity.GetAttributeValue<OptionSetValue>("hil_adjustmentstatus");

                    if (_adjStatus.Value == 2 || _adjStatus.Value == 3)
                    { //Setting Approver Name on Submit for Approval 
                        Entity _entDivision = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_productdivision"));
                        string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='hil_inventoryspareadjustmentline'>
                            <attribute name='hil_partcode' />
                            <order attribute='hil_partcode' descending='false' />
                            <filter type='and'>
                              <condition attribute='hil_adjustmentnumber' operator='eq' value='{entity.Id}' />
                              <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            <link-entity name='product' from='productid' to='hil_partcode' link-type='inner' alias='ac'>
                              <filter type='and'>
                                <condition attribute='hil_division' operator='ne' value='{_entDivision.Id}' />
                              </filter>
                            </link-entity>
                          </entity>
                        </fetch>";
                        EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (_entCol.Entities.Count > 0)
                        {
                            throw new Exception($"ERROR! {_entCol.Entities.Count} Spare Part(s) do not belong to {_entDivision.GetAttributeValue<EntityReference>("hil_productdivision").Name}");
                        }
                    }
                }
            }
        }
    }
}
