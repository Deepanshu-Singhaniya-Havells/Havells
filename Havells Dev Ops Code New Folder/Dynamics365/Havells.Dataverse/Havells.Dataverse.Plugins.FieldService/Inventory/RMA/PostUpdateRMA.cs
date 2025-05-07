using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PostUpdateRMA : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "hil_inventoryrma" && (context.MessageName.ToUpper() == "UPDATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (entity.Contains("hil_rmastatus"))
                    {
                        int _rmastatus = entity.GetAttributeValue<OptionSetValue>("hil_rmastatus").Value;
                        if (_rmastatus == 4)//Posted
                        {
                            Entity _entHeader = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("ownerid", "hil_franchise", "hil_warehouse"));
                            if (_entHeader != null)
                            {
                                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='hil_inventoryproductjournal'>
                                    <attribute name='hil_inventoryproductjournalid' />
                                    <attribute name='hil_name' />
                                    <attribute name='createdon' />
                                    <attribute name='hil_partcode' />
                                    <attribute name='hil_quantity' />
                                    <order attribute='hil_name' descending='false' />
                                    <filter type='and'>
                                        <condition attribute='hil_rma' operator='eq' value='{entity.Id}' />
                                        <condition attribute='statecode' operator='eq' value='0' />
                                        <condition attribute='hil_isused' operator='eq' value='1' />
                                    </filter>
                                    </entity>
                                </fetch>";
                                EntityCollection entCollRMA = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                if (entCollRMA.Entities.Count > 0)
                                {
                                    foreach (Entity _entLine in entCollRMA.Entities)
                                    {
                                        InventoryJournalDTO _IJData = new InventoryJournalDTO()
                                        {
                                            channelPartner = _entHeader.GetAttributeValue<EntityReference>("hil_franchise"),
                                            warehouse = _entHeader.GetAttributeValue<EntityReference>("hil_warehouse"),
                                            partCode = _entLine.GetAttributeValue<EntityReference>("hil_partcode"),
                                            quantity = -_entLine.GetAttributeValue<int>("hil_quantity"),
                                            isRevert = false,
                                            rma = new EntityReference(entity.LogicalName, entity.Id),
                                            owner = _entHeader.GetAttributeValue<EntityReference>("ownerid"),
                                        };
                                        if (!CheckForDuplicateRMALine(service, _IJData.channelPartner, _IJData.warehouse, _IJData.partCode, _IJData.rma))
                                        {
                                            InventoryServices _invServices = new InventoryServices();
                                            _invServices.CreateInventoryJournal(service, _IJData);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    #region Old Function
                    //if (entity.Contains("hil_rmastatus"))
                    //{
                    //    int _rmastatus = entity.GetAttributeValue<OptionSetValue>("hil_rmastatus").Value;
                    //    if (_rmastatus == 4)//Posted
                    //    {
                    //        Entity _entHeader = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_syncedwithsap", "hil_inspectionnumber","ownerid", "hil_franchise", "hil_warehouse"));

                    //        bool IsSyncWithSAP = _entHeader.GetAttributeValue<bool>("hil_syncedwithsap");
                    //        string inspectionNum = _entHeader.GetAttributeValue<string>("hil_inspectionnumber");
                    //        if (IsSyncWithSAP && !string.IsNullOrWhiteSpace(inspectionNum))
                    //        {
                    //            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    //                <entity name='hil_inventoryrmaline'>
                    //                <attribute name='hil_inventoryrmalineid'/>
                    //                <attribute name='hil_name'/>
                    //                <attribute name='createdon'/>
                    //                <attribute name='hil_rma'/>
                    //                <attribute name='hil_quantity'/>
                    //                <attribute name='hil_product'/>
                    //                <attribute name='hil_jobproduct'/>
                    //                <attribute name='hil_job'/>
                    //                <order attribute='hil_name' descending='false'/>
                    //                <filter type='and'>
                    //                    <condition attribute='statecode' operator='eq' value='0'/>
                    //                    <condition attribute='hil_quantity' operator='gt' value='0'/>
                    //                    <condition attribute='hil_rma' operator='eq' value='{entity.Id}'/>
                    //                    <condition attribute='hil_product' operator='not-null'/>
                    //                </filter>
                    //                </entity>
                    //                </fetch>";
                    //            EntityCollection entCollRMA = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    //            foreach (Entity _entLine in entCollRMA.Entities)
                    //            {
                    //                InventoryJournalDTO _IJData = new InventoryJournalDTO()
                    //                {
                    //                    channelPartner = _entHeader.GetAttributeValue<EntityReference>("hil_franchise"),
                    //                    warehouse = _entHeader.GetAttributeValue<EntityReference>("hil_warehouse"),
                    //                    partCode = _entLine.GetAttributeValue<EntityReference>("hil_product"),
                    //                    quantity = -_entLine.GetAttributeValue<int>("hil_quantity"),
                    //                    isRevert = false,
                    //                    rmaLine = _entLine.ToEntityReference(),
                    //                    owner = _entHeader.GetAttributeValue<EntityReference>("ownerid"),
                    //                };
                    //                if (!CheckForDuplicateRMALine(service, _IJData.channelPartner, _IJData.warehouse, _IJData.partCode, _IJData.rmaLine))
                    //                {
                    //                    InventoryServices _invServices = new InventoryServices();
                    //                    _invServices.CreateInventoryJournal(service, _IJData);
                    //                }
                    //            }
                    //        }
                    //    }
                    //}
                    #endregion
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }

        public bool CheckForDuplicateRMALine(IOrganizationService service, EntityReference entRefAccount, EntityReference _entRefDefectiveWH, EntityReference _entRefProd, EntityReference rma)
        {
            bool _retFlag = false;
            try
            {
                string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_inventoryproductjournal'>
                    <attribute name='hil_inventoryproductjournalid'/>
                    <filter type='and'>
                        <condition attribute='hil_warehouse' operator='eq' value='{_entRefDefectiveWH.Id}'/>
                        <condition attribute='hil_partcode' operator='eq' value='{_entRefProd.Id}'/>
                        <condition attribute='hil_franchise' operator='eq' value='{entRefAccount.Id}'/>
                        <condition attribute='hil_rmaline' operator='eq' value='{rma.Id}'/>
                        <condition attribute='hil_transactiontype' operator='eq' value='4'/>
                        <condition attribute='statecode' operator='eq' value='0'/>
                    </filter>
                    </entity>
                    </fetch>";

                EntityCollection entColl = service.RetrieveMultiple(new FetchExpression(fetchXML));
                if (entColl.Entities.Count > 0)
                {
                    _retFlag = true;
                }
            }
            catch
            {

            }
            return _retFlag;
        }
    }
}
