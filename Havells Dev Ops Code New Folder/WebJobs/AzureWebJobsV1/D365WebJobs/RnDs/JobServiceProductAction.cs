using System;
using System.Collections.Generic;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;


namespace D365WebJobs
{
    class JobServiceProductAction
    {
        public static ServiceBOMResponse PopulateSpareParts(Guid _jobIncidentId, IOrganizationService _service)
        {
            ServiceBOMResponse _retObject = new ServiceBOMResponse() { _partList = new List<ServiceBOM>(), response = true, responseMessage = "OK" };
            EntityCollection _entCol = null;
            try
            {
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='msdyn_workorderincident'>
                    <attribute name='msdyn_workorderincidentid' />
                    <attribute name='msdyn_incidenttype' />
                    <attribute name='msdyn_incidenttype' />
                    <order attribute='msdyn_incidenttype' descending='false' />
                    <filter type='and'>
                      <condition attribute='msdyn_workorderincidentid' operator='eq' value='{_jobIncidentId}' />
                    </filter>
                    <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='msdyn_customerasset' visible='false' link-type='outer' alias='ca'>
                      <attribute name='msdyn_product' />
                    </link-entity>
                  </entity>
                </fetch>";
                EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entCol.Entities.Count > 0)
                {
                    EntityReference _incidentType = entCol.Entities[0].GetAttributeValue<EntityReference>("msdyn_incidenttype");
                    EntityReference _modelCode = ((EntityReference)(entCol.Entities[0].GetAttributeValue<AliasedValue>("ca.msdyn_product").Value));

                    ExecuteMultipleRequest multiReqExec = new ExecuteMultipleRequest()
                    {
                        Settings = new ExecuteMultipleSettings()
                        {
                            ContinueOnError = false,
                            ReturnResponses = true
                        },
                        Requests = new OrganizationRequestCollection()
                    };
                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_incidenttypeproduct'>
                        <attribute name='msdyn_product' />
                        <attribute name='msdyn_quantity' />
                        <attribute name='msdyn_incidenttypeproductid' />
                        <order attribute='msdyn_lineorder' descending='false' />
                        <filter type='and'>
                            <condition attribute='statecode' operator='eq' value='0' />
                            <condition attribute='msdyn_incidenttype' operator='eq' value='{_incidentType.Id}' />
                            <condition attribute='hil_model' operator='eq' value='{_modelCode.Id}' />
                        </filter>
                        <link-entity name='product' from='productid' to='msdyn_product' visible='false' link-type='outer' alias='prd'>
                            <attribute name='description' />
                            <attribute name='hil_amount' />
                        </link-entity>
                        </entity>
                        </fetch>";

                    _entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (_entCol.Entities.Count > 0)
                    {
                        foreach (Entity incProd in _entCol.Entities)
                        {
                            _retObject._partList.Add(new ServiceBOM()
                            {
                                PartCode = incProd.Contains("msdyn_product") ? incProd.GetAttributeValue<EntityReference>("msdyn_product").Name : null,
                                PartDescription = incProd.Contains("prd.description") ? incProd.GetAttributeValue<AliasedValue>("prd.description").Value.ToString() : null,
                                UnitAmount = incProd.Contains("prd.hil_amount") ? ((Money)(incProd.GetAttributeValue<AliasedValue>("prd.hil_amount").Value)).Value : 0,
                                Quantity = incProd.Contains("msdyn_quantity") ? Convert.ToInt32(incProd.GetAttributeValue<double>("msdyn_quantity")) : 0
                            });
                        }
                    }
                    else
                    {
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='hil_servicebom'>
                            <attribute name='hil_quantity' />
                            <attribute name='hil_priority' />
                            <attribute name='hil_product' />
                            <attribute name='hil_servicebomid' />
                            <order attribute='hil_priority' descending='false' />
                            <filter type='and'>
                              <condition attribute='hil_productcategory' operator='eq' value='{_modelCode.Id}' />
                            </filter>
                            <link-entity name='product' from='productid' to='hil_product' visible='false' link-type='outer' alias='prd'>
                              <attribute name='description' />
                              <attribute name='hil_amount' />
                            </link-entity>
                          </entity>
                        </fetch>";

                        _entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML)); ;
                        if (_entCol.Entities.Count > 0)
                        {
                            foreach (Entity incProd in _entCol.Entities)
                            {
                                _retObject._partList.Add(new ServiceBOM()
                                {
                                    PartCode = incProd.Contains("hil_product") ? incProd.GetAttributeValue<EntityReference>("hil_product").Name : null,
                                    PartDescription = incProd.Contains("prd.description") ? incProd.GetAttributeValue<AliasedValue>("prd.description").Value.ToString() : null,
                                    UnitAmount = incProd.Contains("prd.hil_amount") ? ((Money)(incProd.GetAttributeValue<AliasedValue>("prd.hil_amount").Value)).Value : 0,
                                    Quantity = incProd.Contains("hil_quantity") ? incProd.GetAttributeValue<int>("hil_quantity") : 0,
                                });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _retObject.response = false;
                _retObject.responseMessage = ex.Message;
                _retObject._partList = null;
            }
            return _retObject;
        }
    }
    public class ServiceBOMResponse
    {
        public bool response { get; set; }
        public string responseMessage { get; set; }
        public List<ServiceBOM> _partList { get; set; }
    }
    public class ServiceBOM
    {
        public int Quantity { get; set; }
        public string PartCode { get; set; }
        public string PartDescription { get; set; }
        public decimal UnitAmount { get; set; }
        public string WarrantyStatus { get; set; }
    }
}
