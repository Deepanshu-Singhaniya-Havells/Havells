using Havells.Dataverse.Plugins.FieldService.Inventory;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

namespace Havells.Dataverse.Plugins.FieldService.Work_Order
{
    public class PostUpdateWorkOrderCreateSpareConsumption : IPlugin
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
                    && context.PrimaryEntityName.ToLower() == "msdyn_workorder" && context.MessageName.ToUpper() == "UPDATE")
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                if (entity.Attributes.Contains("hil_calculatecharges"))
                {
                    Boolean isWorkDone = entity.GetAttributeValue<Boolean>("hil_calculatecharges");

                    if (isWorkDone)
                        ProcessRequest(entity, _tracingService, service);
                }
            }
        }
        public void ProcessRequest(Entity entity, ITracingService _tracingService, IOrganizationService service)
        {
            try
            {
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='msdyn_workorderproduct'>
                    <attribute name='hil_replacedpart' />
                    <attribute name='msdyn_workorderproductid' />
                    <attribute name='msdyn_workorder' />
                    <attribute name='hil_warrantystatus' />
                    <attribute name='msdyn_quantity' />
                    <attribute name='msdyn_customerasset' />
                    <order attribute='msdyn_product' descending='false' />
                    <filter type='and'>
                        <condition attribute='hil_availabilitystatus' operator='eq' value='1' />
                        <condition attribute='hil_markused' operator='eq' value='1' />
                        <condition attribute='statecode' operator='eq' value='0' />
                        <condition attribute='hil_replacedpart' operator='not-null' />
                        <condition attribute='msdyn_quantity' operator='gt' value='0' />
                        <condition attribute='msdyn_workorder' operator='eq' value='{entity.Id}' />
                    </filter>
                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' visible='false' link-type='outer' alias='wo'>
                        <attribute name='hil_owneraccount' />
                        <attribute name='hil_warrantysubstatus' />
                    </link-entity>
                    <link-entity name='product' from='productid' to='hil_replacedpart' link-type='inner' alias='pd'>
                        <attribute name='hil_hierarchylevel' />
                    </link-entity>
                    </entity>
                    </fetch>";
                EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entCol.Entities.Count > 0)
                {
                    EntityReference entRefAccount = (EntityReference)entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.hil_owneraccount").Value;
                    Entity entAccount = service.Retrieve(entRefAccount.LogicalName, entRefAccount.Id, new ColumnSet("ownerid", "hil_spareinventoryenabled"));
                    bool isInventoryEnabled = entAccount.GetAttributeValue<Boolean>("hil_spareinventoryenabled");
                    EntityReference entRefCustomerAsset = entCol.Entities[0].GetAttributeValue<EntityReference>("msdyn_customerasset");

                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
                    <entity name='hil_inventorysettings'>
                    <attribute name='hil_inventorysettingsid' />
                    <attribute name='hil_amcdefectivermatype' />
                    <attribute name='hil_warrantydefectivermatype' />
                    <attribute name='modifiedon' />
                    <filter type='and'>
                        <condition attribute='statecode' operator='eq' value='0' />
                    </filter>
                    </entity>
                    </fetch>";
                    EntityCollection _entColSettings = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    EntityReference _warrantyDefectiveChallan = null;
                    EntityReference _amcDefectiveChallan = null;

                    if (_entColSettings.Entities.Count > 0)
                    {
                        _warrantyDefectiveChallan = _entColSettings.Entities.First().GetAttributeValue<EntityReference>("hil_warrantydefectivermatype");
                        _amcDefectiveChallan = _entColSettings.Entities.First().GetAttributeValue<EntityReference>("hil_amcdefectivermatype");
                    }
                    else {
                        throw new InvalidPluginExecutionException($"Inventory RMA Return Type Setting is not defined.");
                    }
                    if (isInventoryEnabled)
                    {
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_inventorywarehouse'>
                            <attribute name='hil_inventorywarehouseid' />
                            <attribute name='ownerid' />
                            <attribute name='hil_type' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                                <condition attribute='hil_franchise' operator='eq' value='{entRefAccount.Id}' />
                                <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            </entity>
                        </fetch>";
                        EntityCollection entColWH = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        EntityReference _entRefFreshWH = null;
                        EntityReference _entRefDefectiveWH = null;

                        foreach (Entity entWH in entColWH.Entities)
                        {
                            if (entWH.GetAttributeValue<OptionSetValue>("hil_type").Value == 1)
                            {
                                _entRefFreshWH = entWH.ToEntityReference();
                            }
                            if (entWH.GetAttributeValue<OptionSetValue>("hil_type").Value == 2)
                            {
                                _entRefDefectiveWH = entWH.ToEntityReference();
                            }
                        }
                        if (_entRefFreshWH == null || _entRefFreshWH == null)
                        {
                            throw new InvalidPluginExecutionException($"Franchise/DSE's Inventory Warehouse is not defined.");
                        }
                        int _qty = 0;
                        OptionSetValue _warrantyStatus = null;
                        OptionSetValue _warrantySubStatus = null;
                        EntityReference _entRefProd = null;
                        InventoryServices _invServices = new InventoryServices();
                        InventoryJournalDTO _IJData = null;
                        foreach (Entity ent in entCol.Entities)
                        {
                            _warrantySubStatus = null;
                            _warrantyStatus = null;
                            _qty = (Int32)ent.GetAttributeValue<double>("msdyn_quantity");
                            if (ent.Contains("hil_warrantystatus"))
                                _warrantyStatus = ent.GetAttributeValue<OptionSetValue>("hil_warrantystatus");
                            if (ent.Contains("wo.hil_warrantysubstatus"))
                                _warrantySubStatus = (OptionSetValue)ent.GetAttributeValue<AliasedValue>("wo.hil_warrantysubstatus").Value;
                            else { 
                             Entity _entCAWarrantySubstatus = service.Retrieve(entRefCustomerAsset.LogicalName, entRefCustomerAsset.Id, new ColumnSet("hil_warrantysubstatus"));
                                if (_entCAWarrantySubstatus.Contains("hil_warrantysubstatus"))
                                    _warrantySubStatus = _entCAWarrantySubstatus.GetAttributeValue<OptionSetValue>("hil_warrantysubstatus");
                            }
                            _entRefProd = ent.GetAttributeValue<EntityReference>("hil_replacedpart");
                            OptionSetValue partHierarchyLevel = (OptionSetValue)ent.GetAttributeValue<AliasedValue>("pd.hil_hierarchylevel").Value;
                            if (partHierarchyLevel.Value != 910590001) //If Product Hierarchy Type ! = "AMC"
                            {
                                _IJData = new InventoryJournalDTO()
                                {
                                    channelPartner = entRefAccount,
                                    warehouse = _entRefFreshWH,
                                    partCode = _entRefProd,
                                    quantity = _qty * -1,
                                    isRevert = false,
                                    jobProduct = ent.ToEntityReference(),
                                    owner = entAccount.GetAttributeValue<EntityReference>("ownerid"),
                                };
                                if (!CheckForDuplicateJobProduct(service, entRefAccount, _entRefFreshWH, _entRefProd, ent.ToEntityReference()))
                                    _invServices.CreateInventoryJournal(service, _IJData);

                                if (_warrantyStatus.Value == 1) //IN Warranty Spare 
                                {
                                    _IJData = new InventoryJournalDTO()
                                    {
                                        channelPartner = entRefAccount,
                                        warehouse = _entRefDefectiveWH,
                                        partCode = _entRefProd,
                                        quantity = _qty,
                                        isRevert = false,
                                        jobProduct = ent.ToEntityReference(),
                                        owner = entAccount.GetAttributeValue<EntityReference>("ownerid"),
                                        returnType = _warrantySubStatus != null ? _warrantySubStatus.Value == 4 ? _amcDefectiveChallan  : _warrantyDefectiveChallan : _warrantyDefectiveChallan
                                    };
                                    if (!CheckForDuplicateJobProduct(service, entRefAccount, _entRefDefectiveWH, _entRefProd, ent.ToEntityReference()))
                                        _invServices.CreateInventoryJournal(service, _IJData);
                                }
                            }
                        }
                    }
                    else
                    {
                        throw new InvalidPluginExecutionException($"Franchise/DSE Inventory Warehouse is not defined.");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException($"ERROR.{ex.Message}");
            }
        }


        public bool CheckForDuplicateJobProduct(IOrganizationService service, EntityReference entRefAccount, EntityReference _entRefDefectiveWH, EntityReference _entRefProd, EntityReference jobProduct)
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
                        <condition attribute='hil_jobproduct' operator='eq' value='{jobProduct.Id}'/>
                        <condition attribute='hil_transactiontype' operator='eq' value='3'/>
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
