using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace Havells_Plugin.WOProduct
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
                string _jobIncidentId = context.InputParameters.Contains("JobIncidentId") ? context.InputParameters["JobIncidentId"].ToString() : null;
                string _serviceBOMId = context.InputParameters.Contains("ServiceBOMId") ? context.InputParameters["ServiceBOMId"].ToString() : null;

                if (string.IsNullOrEmpty(_jobIncidentId) || string.IsNullOrEmpty(_serviceBOMId))
                {
                    context.OutputParameters["status_code"] = "204";
                    context.OutputParameters["status_description"] = "Job Incident Id and Service BOM Id is required.";
                }
                else
                {
                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorderincident'>
                        <attribute name='msdyn_workorderincidentid' />
                        <attribute name='msdyn_workorder' />
                        <attribute name='msdyn_customerasset' />
                        <filter type='and'>
                            <condition attribute='msdyn_workorderincidentid' operator='eq' value='{_jobIncidentId}' />
                            <condition attribute='statecode' operator='eq' value='0' />
                        </filter>
                        </entity>
                    </fetch>";

                    EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entCol.Entities.Count > 0)
                    {
                        try
                        {
                            Entity _serviceBOM = service.Retrieve("hil_servicebom", new Guid(_serviceBOMId), new ColumnSet("hil_isserialized", "hil_quantity", "hil_product", "hil_priority", "hil_chargeableornot"));
                            if (_serviceBOM != null)
                            {

                                Entity _jobProduct = new Entity("msdyn_workorderproduct");
                                _jobProduct["msdyn_customerasset"] = entCol.Entities[0].GetAttributeValue<EntityReference>("msdyn_customerasset");

                                if (_serviceBOM.Contains("hil_product"))
                                {
                                    EntityReference _sparePart = _serviceBOM.GetAttributeValue<EntityReference>("hil_product");
                                    _jobProduct["msdyn_product"] = _sparePart;
                                    Product Pdt1 = (Product)service.Retrieve(Product.EntityLogicalName, _sparePart.Id, new ColumnSet("name", "description", "hil_amount"));
                                    string Uq = Pdt1.Description != null ? _sparePart.Name + "-" + Pdt1.Description : _sparePart.Name + "-";
                                    _jobProduct["hil_part"] = Uq;
                                    _jobProduct["hil_priority"] = _serviceBOM.Contains("hil_priority") ? (string)_serviceBOM["hil_priority"] : string.Empty;

                                    _jobProduct["msdyn_workorderincident"] = entCol.Entities[0].ToEntityReference();

                                    _jobProduct["msdyn_workorder"] = entCol.Entities[0].GetAttributeValue<EntityReference>("msdyn_workorder");

                                    _jobProduct["hil_maxquantity"] = Convert.ToDecimal(_serviceBOM.Contains("hil_quantity") ? _serviceBOM.GetAttributeValue<int>("hil_quantity") : 1);
                                    _jobProduct["msdyn_quantity"] = Convert.ToDouble(1);

                                    if (_serviceBOM.Contains("hil_chargeableornot"))
                                    {
                                        OptionSetValue _chargeableOS = _serviceBOM.GetAttributeValue<OptionSetValue>("hil_chargeableornot");
                                        _jobProduct["hil_chargeableornot"] = _chargeableOS;
                                        if (_chargeableOS.Value == 1)
                                        {
                                            _jobProduct["hil_warrantystatus"] = new OptionSetValue(2);
                                        }
                                    }
                                    if (Pdt1.hil_Amount != null)
                                    {
                                        _jobProduct["msdyn_totalamount"] = Pdt1.hil_Amount;
                                        _jobProduct["hil_partamount"] = Pdt1.hil_Amount.Value;
                                    }
                                    if (_serviceBOM.Contains("hil_isserialized"))
                                    {
                                        _jobProduct["hil_isserialized"] = _serviceBOM.GetAttributeValue<OptionSetValue>("hil_isserialized");
                                    }
                                }
                                service.Create(_jobProduct);
                            }

                            context.OutputParameters["status_code"] = "200";
                            context.OutputParameters["status_description"] = "OK.";
                        }
                        catch (Exception ex)
                        {
                            context.OutputParameters["status_code"] = "204";
                            context.OutputParameters["status_description"] = ex.Message;
                        }
                    }
                    else
                    {
                        context.OutputParameters["status_code"] = "204";
                        context.OutputParameters["status_description"] = "Job Incident Id does not exist.";
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
