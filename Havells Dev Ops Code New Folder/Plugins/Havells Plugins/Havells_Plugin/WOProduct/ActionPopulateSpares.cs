using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace Havells_Plugin.WOProduct
{
    public class ActionPopulateSpares : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                //string _jobIncidentId = context.InputParameters.Contains("JobIncidentId") ? context.InputParameters["JobIncidentId"].ToString() : null;
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is EntityReference)
                {
                    EntityReference ent = (EntityReference)context.InputParameters["Target"];
                    if (ent==null)
                    {
                        context.OutputParameters["partList"] = null;
                        context.OutputParameters["status_code"] = "204";
                        context.OutputParameters["status_description"] = "Job Incident Id is required.";
                    }
                    else
                    {
                        string _jobIncidentId = ent.Id.ToString();
                        List<PartList> _retObj = new List<PartList>();
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

                            EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (entCol.Entities.Count > 0)
                            {
                                EntityReference _incidentType = entCol.Entities[0].GetAttributeValue<EntityReference>("msdyn_incidenttype");
                                EntityReference _modelCode = ((EntityReference)(entCol.Entities[0].GetAttributeValue<AliasedValue>("ca.msdyn_product").Value));

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

                                _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML)); ;
                                if (_entCol.Entities.Count > 0)
                                {
                                    foreach (Entity incProd in _entCol.Entities)
                                    {
                                        _retObj.Add(new PartList()
                                        {
                                            RecordId =incProd.Id.ToString(),
                                            PartCode = incProd.Contains("hil_product") ? incProd.GetAttributeValue<EntityReference>("hil_product").Name : null,
                                            PartDescription = incProd.Contains("prd.description") ? incProd.GetAttributeValue<AliasedValue>("prd.description").Value.ToString() : null,
                                            UnitAmount = incProd.Contains("prd.hil_amount") ? ((Money)(incProd.GetAttributeValue<AliasedValue>("prd.hil_amount").Value)).Value : 0,
                                            Quantity = incProd.Contains("hil_quantity") ? incProd.GetAttributeValue<int>("hil_quantity") : 0,
                                        });
                                    }
                                }
                            }
                            context.OutputParameters["partList"] = JsonConvert.SerializeObject(_retObj);
                            context.OutputParameters["status_code"] = "200";
                            context.OutputParameters["status_description"] = "OK.";
                        }
                        catch (Exception ex)
                        {
                            context.OutputParameters["partList"] = null;
                            context.OutputParameters["status_code"] = "204";
                            context.OutputParameters["status_description"] = ex.Message;
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
    public class PartList
    {
        public string RecordId { get; set; }
        public int Quantity { get; set; }
        public string PartCode { get; set; }
        public string PartDescription { get; set; }
        public decimal UnitAmount { get; set; }
        public string WarrantyStatus { get; set; }
    }
}
