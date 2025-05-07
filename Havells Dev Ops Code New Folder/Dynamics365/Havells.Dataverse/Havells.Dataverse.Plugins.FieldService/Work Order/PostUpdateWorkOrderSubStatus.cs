using Havells.Dataverse.Plugins.FieldService.Inventory;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;

namespace Havells.Dataverse.Plugins.FieldService.Work_Order
{
    internal class PostUpdateWorkOrderSubStatus : IPlugin
    {
        private static readonly Guid Job_Workdone = new Guid("2927fa6c-fa0f-e911-a94e-000d3af060a1");

        public void Execute(IServiceProvider serviceProvider)
        {

            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                        && context.PrimaryEntityName.ToLower() == "msdyn_workorder" && (context.MessageName.ToUpper() == "UPDATE"))
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                UpdateInventoryAgainstConsumption(service, entity);
            }
        }
        public void UpdateInventoryAgainstConsumption(IOrganizationService service, Entity entity)
        {
            try
            {
                Guid _substatus = entity.GetAttributeValue<EntityReference>("msdyn_substatus").Id;

                if (_substatus == Job_Workdone)
                {
                    string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='msdyn_workorder'>
                                    <attribute name='msdyn_name' />
                                    <attribute name='hil_customerref' />
                                    <attribute name='hil_owneraccount' />
                                    <attribute name='msdyn_workorderid' />
                                    <order attribute='msdyn_name' descending='false' />
                                    <filter type='and'>
                                     <condition attribute='msdyn_workorderid' operator='eq' value='{entity.Id}' />
                                    </filter>
                                    <link-entity name='account' from='accountid' to='hil_owneraccount' link-type='inner' alias='aa'>
                                      <filter type='and'>
                                        <condition attribute='hil_spareinventoryenabled' operator='eq' value='1' />
                                      </filter>
                                    </link-entity>
                                  </entity>
                                </fetch>";

                    EntityCollection entColljob = service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (entColljob.Entities.Count > 0)
                    {
                        EntityReference franchise = entity.Contains("hil_owneraccount") ? entity.GetAttributeValue<EntityReference>("hil_owneraccount") : null;
                        EntityReference _entRefDefectiveWH = null;
                        EntityReference _entRefFreshWH = null;
                        EntityReference _entOwner = null;
                        string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='hil_inventorywarehouse'>
                                        <attribute name='hil_inventorywarehouseid' />
                                        <attribute name='ownerid' />
                                        <attribute name='hil_type' />
                                        <order attribute='hil_name' descending='false' />
                                        <filter type='and'>
                                          <condition attribute='hil_franchise' operator='eq' value='{franchise.Id}' />
                                          <condition attribute='statecode' operator='eq' value='0' />
                                        </filter>
                                      </entity>
                                    </fetch>";
                        EntityCollection entColWH = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entColWH.Entities.Count > 0)
                        {
                            _entOwner = entColWH.Entities[0].GetAttributeValue<EntityReference>("ownerid");

                            foreach (Entity _entWarehouse in entColWH.Entities)
                            {
                                if (_entWarehouse.GetAttributeValue<OptionSetValue>("hil_type").Value == 1)
                                {
                                    _entRefFreshWH = _entWarehouse.ToEntityReference();
                                }
                                if (_entWarehouse.GetAttributeValue<OptionSetValue>("hil_type").Value == 2)
                                {
                                    _entRefDefectiveWH = _entWarehouse.ToEntityReference();
                                }
                            }
                        }
                        fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='msdyn_workorderproduct'>
                                    <attribute name='createdon' />
                                    <attribute name='msdyn_product' />
                                    <attribute name='msdyn_workorderproductid' />
                                    <attribute name='msdyn_workorder' />
                                    <attribute name='hil_replacedpart' />
                                    <attribute name='hil_warrantystatus' />
                                    <attribute name='msdyn_quantity' />
                                    <order attribute='msdyn_product' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='hil_markused' operator='eq' value='1' />
                                      <condition attribute='statecode' operator='eq' value='0' />
                                      <condition attribute='hil_replacedpart' operator='not-null' />
                                      <condition attribute='msdyn_quantity' operator='gt' value='0' />
                                      <condition attribute='msdyn_workorder' operator='eq' value='{entity.Id}' />
                                    </filter>
                                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' visible='false' link-type='outer' alias='ac'>
                                      <attribute name='hil_owneraccount' />
                                    </link-entity>
                                  </entity>
                                </fetch>";
                        EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(fetchXML));
                        if (entCol.Entities.Count > 0)
                        {
                            double _qty = 0;
                            foreach (Entity _entLine in entCol.Entities)
                            {
                                _qty = 0;
                                if (_entLine.Contains("msdyn_quantity"))
                                {
                                    _qty = _entLine.GetAttributeValue<double>("msdyn_quantity");
                                    int warrantystatus = _entLine.GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value;
                                    if (_qty > 0)
                                    {
                                        InventoryJournalDTO _IJData = new InventoryJournalDTO()
                                        {
                                            channelPartner = franchise,
                                            warehouse = _entRefDefectiveWH,
                                            partCode = _entLine.GetAttributeValue<EntityReference>("hil_replacedpart"),
                                            quantity = Convert.ToInt32(_qty),
                                            isRevert = false,
                                            jobProduct = _entLine.ToEntityReference(),
                                            owner = _entOwner,
                                        };
                                        InventoryServices _invServices = new InventoryServices();
                                        if (warrantystatus == 1)// In Warranty
                                        {
                                            _invServices.CreateInventoryJournal(service, _IJData);
                                        }

                                        _IJData.quantity = -Convert.ToInt32(_qty);
                                        _IJData.warehouse = _entRefFreshWH;
                                        _invServices.CreateInventoryJournal(service, _IJData);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }

}
