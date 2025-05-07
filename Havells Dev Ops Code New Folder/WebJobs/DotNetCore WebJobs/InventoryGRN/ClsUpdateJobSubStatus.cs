using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace InventoryGRN
{
    public class ClsUpdateJobSubStatus
    {
        private readonly ServiceClient _service;
        public ClsUpdateJobSubStatus(ServiceClient service)
        {
            _service = service;
        }
        public void UpdatePOJobSubStatus()
        {
            try
            {
                int total = 0;
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='msdyn_workorder'>
                                        <attribute name='msdyn_name' />
                                        <attribute name='createdon' />
                                        <attribute name='hil_productsubcategory' />
                                        <attribute name='hil_customerref' />
                                        <attribute name='hil_callsubtype' />
                                        <attribute name='msdyn_workorderid' />
                                        <order attribute='msdyn_name' descending='false' />
                                        <filter type='and'>
                                            <condition attribute='msdyn_substatus' operator='eq' value='{{1B27FA6C-FA0F-E911-A94E-000D3AF060A1}}' /> ///Part PO Created
                                        </filter>
                                            <link-entity name='hil_inventorypurchaseorder' from='hil_jobid' to='msdyn_workorderid' link-type='inner' alias='ad' />
                                      </entity>
                                    </fetch>";

                EntityCollection entJobColl = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entJobColl.Entities.Count > 0)
                {
                    int i = 0;
                    total += entJobColl.Entities.Count;
                    foreach (var Job in entJobColl.Entities)
                    {
                        i++;
                        Console.WriteLine($"Processing# {i}/{entJobColl.Entities.Count} Job# {Job.GetAttributeValue<string>("msdyn_name")}");
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='hil_inventorypurchaseorderline'>
                                            <attribute name='hil_inventorypurchaseorderlineid' />
                                            <attribute name='hil_name' />
                                            <attribute name='createdon' />
                                            <order attribute='hil_name' descending='false' />
                                            <filter type='and'>
                                              <condition attribute='hil_pendingquantity' operator='gt' value='0' />
                                            </filter>
                                            <link-entity name='hil_inventorypurchaseorder' from='hil_inventorypurchaseorderid' to='hil_ponumber' link-type='inner' alias='aa'>
                                              <filter type='and'>
                                                <condition attribute='hil_jobid' operator='eq' value='{Job.Id}' />
                                              </filter>
                                            </link-entity>
                                          </entity>
                                        </fetch>";
                        EntityCollection entOrderlineColl = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entOrderlineColl.Entities.Count > 0)
                        {
                            foreach (var orderline in entOrderlineColl.Entities)
                            {
                                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' aggregate='true'>
                                      <entity name='hil_inventorypurchaseorderreceiptline'>
                                        <attribute name='hil_billedquantity' aggregate='sum' alias='totalbilledquantity' />
                                        <filter type='and'>
                                            <condition attribute='hil_purchaseorderline' operator='eq' value='{orderline.Id}' />
                                        </filter>
                                      </entity>
                                    </fetch>";
                                EntityCollection entColl = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                if (entColl.Entities.Count > 0)
                                {
                                    Int32 totalbilledquantity = Convert.ToInt32(((AliasedValue)entColl.Entities[0]["totalbilledquantity"]).Value);
                                    if (totalbilledquantity > 0)
                                    {
                                        Entity entorderline = new Entity(orderline.LogicalName, orderline.Id);
                                        entorderline["hil_suppliedquantity"] = totalbilledquantity;
                                        _service.Update(entorderline);

                                        Entity _entPOLine = _service.Retrieve(orderline.LogicalName, orderline.Id, new ColumnSet("hil_pendingquantity"));
                                        if (_entPOLine != null)
                                        {
                                            Entity _etUpdateLine = new Entity(orderline.LogicalName, orderline.Id);
                                            int _pendingQty = _entPOLine.GetAttributeValue<int>("hil_pendingquantity");
                                            if (_pendingQty > 0)
                                                _etUpdateLine["hil_partstatus"] = new OptionSetValue(2);//Partial Dispatched	
                                            else
                                                _etUpdateLine["hil_partstatus"] = new OptionSetValue(3);//Dispatched
                                            _service.Update(_etUpdateLine);
                                        }
                                    }
                                }
                            }
                        }
                        //Update Job Sub-Status
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_inventorypurchaseorderline'>
                                <attribute name='hil_inventorypurchaseorderlineid' />
                                <attribute name='hil_name' />
                                <attribute name='createdon' />
                                <order attribute='hil_name' descending='false' />
                                <filter type='and'>
                                    <condition attribute='hil_pendingquantity' operator='gt' value='0' />
                                </filter>
                                <link-entity name='hil_inventorypurchaseorder' from='hil_inventorypurchaseorderid' to='hil_ponumber' link-type='inner' alias='aa'>
                                    <filter type='and'>
                                    <condition attribute='hil_jobid' operator='eq' value='{Job.Id}' />
                                    </filter>
                                </link-entity>
                                </entity>
                            </fetch>";
                        EntityCollection OrderlineColl = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (OrderlineColl.Entities.Count == 0)
                        {
                            Entity wo = new Entity(Job.LogicalName, Job.Id);
                            wo["msdyn_substatus"] = new EntityReference("msdyn_workordersubstatus", new Guid("2b27fa6c-fa0f-e911-a94e-000d3af060a1"));//Work Initiated
                            _service.Update(wo);
                            Console.WriteLine($"Updated Sub-Status for Job# {Job.GetAttributeValue<string>("msdyn_name")}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
