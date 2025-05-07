using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace Havells.Dataverse.CustomConnector.Service_Call
{
    public class ActionAddSparePartToJobProduct : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                string jobIncidentId = context.InputParameters.Contains("JobIncidentId") ? context.InputParameters["JobIncidentId"].ToString() : null;
                string serviceBOMIds = context.InputParameters.Contains("ServiceBOMId") ? context.InputParameters["ServiceBOMId"].ToString() : null;
                if (string.IsNullOrEmpty(jobIncidentId) || string.IsNullOrEmpty(serviceBOMIds))
                {
                    tracingService.Trace("step-1");
                    context.OutputParameters["status_code"] = "204";
                    context.OutputParameters["status_description"] = "Job Incident Id and Service BOM Id are required.";
                    return;
                }
                List<string> listSBOMID = JsonConvert.DeserializeObject<List<string>>(serviceBOMIds);
                string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='msdyn_workorderincident'>
                                    <attribute name='msdyn_workorderincidentid' />
                                    <attribute name='msdyn_workorder' />
                                    <attribute name='msdyn_customerasset' />
                                    <filter type='and'>
                                    <condition attribute='msdyn_workorderincidentid' operator='eq' value='{jobIncidentId}' />
                                    <condition attribute='statecode' operator='eq' value='0' />
                                    </filter>
                                    </entity>
                                    </fetch>";
                EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(fetchXML));
                if (entCol.Entities.Count == 0)
                {
                    tracingService.Trace("step-3");
                    context.OutputParameters["status_code"] = "204";
                    context.OutputParameters["status_description"] = "Job Incident Id does not exist.";
                    return;
                }
                foreach (string serviceBOMId in listSBOMID)
                {
                    try
                    {
                        tracingService.Trace("step-4");
                        Entity serviceBOM = service.Retrieve("hil_servicebom", new Guid(serviceBOMId), new ColumnSet("hil_isserialized", "hil_quantity", "hil_product", "hil_priority", "hil_chargeableornot"));
                        if (serviceBOM == null) continue;
                        Entity jobProduct = new Entity("msdyn_workorderproduct");
                        jobProduct["msdyn_customerasset"] = entCol.Entities[0].GetAttributeValue<EntityReference>("msdyn_customerasset");
                        if (serviceBOM.Contains("hil_product"))
                        {
                            EntityReference sparePart = serviceBOM.GetAttributeValue<EntityReference>("hil_product");

                            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='msdyn_workorderproduct'>
                                <attribute name='msdyn_workorderproductid' />
                                <filter type='and'>
                                  <condition attribute='msdyn_workorderincident' operator='eq' value='{entCol.Entities[0].ToEntityReference().Id}' />
                                  <condition attribute='msdyn_product' operator='eq' value='{sparePart.Id}' />
                                </filter>
                              </entity>
                            </fetch>";
                            EntityCollection _entColDuplicate = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (_entColDuplicate.Entities.Count == 0)
                            {
                                jobProduct["msdyn_product"] = sparePart;
                                Entity product = service.Retrieve("product", sparePart.Id, new ColumnSet("name", "description", "hil_amount"));
                                string partDescription = product.Contains("description") ? $"{sparePart.Name}-{product.GetAttributeValue<string>("description")}" : $"{sparePart.Name}-";
                                jobProduct["hil_part"] = partDescription;
                                jobProduct["hil_priority"] = serviceBOM.Contains("hil_priority") ? serviceBOM["hil_priority"].ToString() : string.Empty;
                                jobProduct["msdyn_workorderincident"] = entCol.Entities[0].ToEntityReference();
                                jobProduct["msdyn_workorder"] = entCol.Entities[0].GetAttributeValue<EntityReference>("msdyn_workorder");
                                jobProduct["hil_maxquantity"] = Convert.ToDecimal(serviceBOM.Contains("hil_quantity") ? serviceBOM.GetAttributeValue<int>("hil_quantity") : 1);
                                jobProduct["msdyn_quantity"] = 1.0;
                                if (serviceBOM.Contains("hil_chargeableornot"))
                                {
                                    OptionSetValue chargeableOption = serviceBOM.GetAttributeValue<OptionSetValue>("hil_chargeableornot");
                                    jobProduct["hil_chargeableornot"] = chargeableOption;
                                    if (chargeableOption.Value == 1)
                                    {
                                        jobProduct["hil_warrantystatus"] = new OptionSetValue(2);
                                    }
                                }
                                if (product.GetAttributeValue<Money>("hil_amount") != null)
                                {
                                    jobProduct["msdyn_totalamount"] = product.GetAttributeValue<Money>("hil_amount");
                                    jobProduct["hil_partamount"] = product.GetAttributeValue<Money>("hil_amount").Value;
                                }
                                if (serviceBOM.Contains("hil_isserialized"))
                                {
                                    jobProduct["hil_isserialized"] = serviceBOM.GetAttributeValue<OptionSetValue>("hil_isserialized");
                                }
                            }
                        }
                        tracingService.Trace("End Step");
                        service.Create(jobProduct);
                    }
                    catch (Exception ex)
                    {
                        context.OutputParameters["status_code"] = "204";
                        context.OutputParameters["status_description"] = ex.Message;
                    }
                }
                context.OutputParameters["status_code"] = "200";
                context.OutputParameters["status_description"] = "OK.";
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}

