using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.Dataverse.CustomConnector.Claim_Module
{
    public class GetClaimOverheadPriceList : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            decimal _rate = 0;
            bool _statusCode = false;

            try
            {
                string _jobId = context.InputParameters.Contains("JobID") ? context.InputParameters["JobID"].ToString() : null;
                int _quantity = context.InputParameters.Contains("Quantity") ? Convert.ToInt32(context.InputParameters["Quantity"].ToString()) : 0;

                if (string.IsNullOrEmpty(_jobId) || _quantity == 0)
                {
                    context.OutputParameters["StatusMessage"] = "Invalid Inputs.";
                    context.OutputParameters["StatusCode"] = _statusCode;
                    context.OutputParameters["Rate"] = _rate;
                    return;
                }
                else
                {
                    Entity entJob = service.Retrieve("msdyn_workorder", new Guid(_jobId), new ColumnSet("hil_salesoffice", "hil_owneraccount", "hil_productcategory"));
                    EntityReference _channelPartner = null;
                    EntityReference _salesOffice = null;
                    EntityReference _productDivision = null;
                    if (entJob != null) {
                        _channelPartner = entJob.GetAttributeValue<EntityReference>("hil_owneraccount");
                        _salesOffice = entJob.GetAttributeValue<EntityReference>("hil_salesoffice");
                        _productDivision = entJob.GetAttributeValue<EntityReference>("hil_productcategory");

                        string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_nonstandardcharges'>
                        <attribute name='hil_nonstandardchargesid' />
                        <attribute name='hil_rate' />
                        <order attribute='hil_channelpartner' descending='true' />
                        <order attribute='hil_salesoffice' descending='true' />
                        <filter type='and'>
                        <condition attribute='statecode' operator='eq' value='0' />
                            <filter type='or'>
                                <condition attribute='hil_channelpartner' operator='null' />
                                <condition attribute='hil_channelpartner' operator='eq' value='{_channelPartner.Id}' />
                            </filter>
                        <condition attribute='hil_division' operator='eq' value='{_productDivision.Id}' />
                            <filter type='or'>
                                <condition attribute='hil_salesoffice' operator='null' />
                                <condition attribute='hil_salesoffice' operator='eq' value='{_salesOffice.Id}' />
                            </filter>
                            <condition attribute='hil_quantityupto' operator='ge' value='{_quantity}' />
                            <condition attribute='hil_quantityto' operator='le' value='{_quantity}' />
                        </filter>
                        </entity>
                        </fetch>";
                        EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entCol.Entities.Count > 0) {
                            _rate = entCol.Entities[0].GetAttributeValue<Money>("hil_rate").Value;
                            _statusCode = true;
                            context.OutputParameters["StatusMessage"] = "Success";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["StatusMessage"] = ex.Message;
            }
            context.OutputParameters["StatusCode"] = _statusCode;
            context.OutputParameters["Rate"] = _rate;
        }
    }
}
