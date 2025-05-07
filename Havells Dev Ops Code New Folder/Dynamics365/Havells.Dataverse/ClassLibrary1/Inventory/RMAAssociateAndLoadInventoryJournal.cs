using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

namespace Havells.Dataverse.CustomConnector.Inventory
{
    public class RMAAssociateAndLoadInventoryJournal : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            bool _statusCode = false;
            try
            {
                string _RMAId = Convert.ToString(context.InputParameters["RMA_Guid"]);
                if (string.IsNullOrWhiteSpace(_RMAId))
                {
                    context.OutputParameters["StatusMessage"] = "First Save RMA Record to generate RMA_ID";
                    context.OutputParameters["StatusCode"] = _statusCode;
                    return;
                }
                else
                {
                    Entity _InventoryRMA = service.Retrieve("hil_inventoryrma", new Guid(_RMAId), new ColumnSet("hil_franchise", "hil_warehouse", "hil_returntype"));
                    if (_InventoryRMA != null)
                    {
                        EntityReference _franchise = _InventoryRMA.GetAttributeValue<EntityReference>("hil_franchise");
                        EntityReference _warehouse = _InventoryRMA.GetAttributeValue<EntityReference>("hil_warehouse");
                        EntityReference _returntype = _InventoryRMA.GetAttributeValue<EntityReference>("hil_returntype");

                        string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_inventoryproductjournal'>
                            <attribute name='hil_inventoryproductjournalid' />
                            <attribute name='hil_name' />
                            <attribute name='createdon' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_transactiontype' operator='eq' value='3' />
                                <condition attribute='hil_franchise' operator='eq' value='{_franchise.Id}' />
                                <condition attribute='hil_warehouse' operator='eq' value='{_warehouse.Id}' />
                                <condition attribute='hil_rmatype' operator='eq' value='{_returntype.Id}' />
                                <condition attribute='hil_rma' operator='null' />
                            </filter>
                            </entity>
                        </fetch>";
                        EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entCol.Entities.Count > 0)
                        {
                            int batchSize = 1000;
                            for (int i = 0; i < entCol.Entities.Count; i += batchSize)
                            {
                                var batch = entCol.Entities.Skip(i).Take(batchSize).ToList();
                                ExecuteMultipleRequest requestWithResults = new ExecuteMultipleRequest()
                                {
                                    // Assign settings that define execution behavior: continue on error, return responses. 
                                    Settings = new ExecuteMultipleSettings()
                                    {
                                        ContinueOnError = false,
                                        ReturnResponses = true
                                    },
                                    // Create an empty organization request collection.
                                    Requests = new OrganizationRequestCollection()
                                };
                                Entity InventoryJournal = null;
                                foreach (var entity in batch)
                                {
                                    InventoryJournal = new Entity("hil_inventoryproductjournal", entity.Id);
                                    InventoryJournal["hil_rma"] = new EntityReference("hil_inventoryrma", new Guid(_RMAId));
                                    UpdateRequest updateRequest = new UpdateRequest() { Target = InventoryJournal };
                                    requestWithResults.Requests.Add(updateRequest);
                                }
                                try
                                {
                                    ExecuteMultipleResponse responseWithResults = (ExecuteMultipleResponse)service.Execute(requestWithResults);
                                    _statusCode = true;
                                    context.OutputParameters["StatusMessage"] = "Completed! All Spare Consumptions have been picked in selected RMA.";
                                }
                                catch (Exception ex)
                                {
                                    context.OutputParameters["StatusMessage"] = $"ERROR! - {ex.Message}";
                                }
                            }
                        }
                        else
                        {
                            context.OutputParameters["StatusMessage"] = "NO additional Spare Consumption found.";
                        }
                    }
                }
                context.OutputParameters["StatusCode"] = _statusCode;
            }
            catch (Exception ex)
            {
                context.OutputParameters["StatusMessage"] = ex.Message;
            }
        }
    }
}
