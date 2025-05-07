using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreUpdateInventoryAdjustmentSetApproverDetails : IPlugin
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

                    if (_adjStatus.Value == 3)
                    {  // Pre Approval validations
                        
                    }
                    if (_adjStatus.Value == 3 || _adjStatus.Value == 6 || _adjStatus.Value == 5)
                    {  // Adjustment is approved/Rejected/Cancelled
                        entity["hil_approvedby"] = new EntityReference("systemuser", context.InitiatingUserId);
                        entity["hil_approvedon"] = DateTime.Now.AddMinutes(330);
                    }
                    else if (_adjStatus.Value == 2)
                    { //Setting Approver Name on Submit for Approval 

                        Entity _entityPO = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_salesoffice", "hil_productdivision"));

                        string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_inventoryspareadjustmentline'>
                        <attribute name='hil_inventoryspareadjustmentlineid' />
                        <filter type='and'>
                            <condition attribute='hil_adjustmentnumber' operator='eq' value='{entity.Id}' />
                            <condition attribute='hil_ajustmentquantity' operator='ne' value='0' />
                            <condition attribute='statecode' operator='eq' value='0' />
                        </filter>
                        </entity>
                        </fetch>";
                        EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(fetchXML));
                        if (entCol.Entities.Count == 0)
                        {
                            throw new Exception("Action denied! There is no Product selected in Adjustment.");
                        }
                        else
                        {
                            fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_sbubranchmapping'>
                            <attribute name='hil_sbubranchmappingid' />
                            <attribute name='hil_branchheaduser' />
                            <order attribute='hil_branchheaduser' descending='false' />
                            <filter type='and'>
                                <condition attribute='hil_productdivision' operator='eq' value='{_entityPO.GetAttributeValue<EntityReference>("hil_productdivision").Id}' />
                                <condition attribute='hil_salesoffice' operator='eq' value='{_entityPO.GetAttributeValue<EntityReference>("hil_salesoffice").Id}' />
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_branchheaduser' operator='not-null' />
                            </filter>
                            </entity>
                            </fetch>";
                            entCol = service.RetrieveMultiple(new FetchExpression(fetchXML));
                            if (entCol.Entities.Count > 0)
                            {
                                entity["hil_approver"] = entCol.Entities[0].GetAttributeValue<EntityReference>("hil_branchheaduser");
                            }
                            else
                            {
                                throw new Exception("Approver Mapping is missing in CRM. Please connect with HO Team.");
                            }
                        }
                    }
                }
            }
        }
    }
}
