using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Home_Advisory
{
    public class GetEnquiryProductType : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
 
            List<EnquiryProductType> lstEnquiryProdType = new List<EnquiryProductType>();
            string JsonResponse = "";
            try
            {
                string EnquiryCategory = Convert.ToString(context.InputParameters["EnquiryCategory"]);

                if (string.IsNullOrWhiteSpace(EnquiryCategory))
                {
                    lstEnquiryProdType.Add(new EnquiryProductType()
                    {
                        EnquiryProductCode = "ERROR",
                        EnquiryProductName = "Enquiry Category is required."
                    });
                    JsonResponse = JsonSerializer.Serialize(lstEnquiryProdType);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                if (service != null)
                {
                    string _fetchxml = string.Empty;

                    if (EnquiryCategory == "1")
                    {
                        _fetchxml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='hil_typeofproduct'>
                            <attribute name='hil_typeofproductid' />
                            <attribute name='hil_name' />
                            <attribute name='hil_typeofenquiry' />
                            <attribute name='hil_index' />
                            <attribute name='hil_code' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                              <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            <link-entity name='hil_enquirytype' from='hil_enquirytypeid' to='hil_typeofenquiry' visible='false' link-type='outer' alias='eq'>
                              <attribute name='hil_enquirytypename' />
                              <attribute name='hil_enquirytypecode' />
                            </link-entity>
                          </entity>
                        </fetch>";
                        EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchxml));
                        if (entCol.Entities.Count > 0)
                        {
                            foreach (Entity ent in entCol.Entities)
                            {
                                lstEnquiryProdType.Add(new EnquiryProductType()
                                {
                                    EnquiryCategory = EnquiryCategory,
                                    EnquiryProductCode = ent.GetAttributeValue<string>("hil_code"),
                                    EnquiryProductName = ent.GetAttributeValue<string>("hil_name"),
                                    EnquiryTypeName = ent.GetAttributeValue<AliasedValue>("eq.hil_enquirytypename").Value.ToString(),
                                    EnquiryTypeCode = ent.GetAttributeValue<AliasedValue>("eq.hil_enquirytypecode").Value.ToString(),
                                });
                                JsonResponse = JsonSerializer.Serialize(lstEnquiryProdType);
                            }
                            context.OutputParameters["data"] = JsonResponse;
                            return;
                        }
                        else
                        {
                            lstEnquiryProdType.Add(new EnquiryProductType() { EnquiryProductCode = "ERROR", EnquiryProductName = "No Data found." });
                            JsonResponse = JsonSerializer.Serialize(lstEnquiryProdType);
                            context.OutputParameters["data"] = JsonResponse;
                            return;
                        }
                    }
                    else if (EnquiryCategory == "2")
                    {
                        _fetchxml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='product'>
                            <attribute name='name' />
                            <attribute name='productnumber' />
                            <attribute name='hil_sapcode' />
                            <order attribute='name' descending='false' />
                            <filter type='and'>
                              <condition attribute='hil_hierarchylevel' operator='eq' value='2' />
                              <condition attribute='producttypecode' operator='eq' value='1' />
                            </filter>
                          </entity>
                        </fetch>";
                        EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchxml));
                        if (entCol.Entities.Count > 0)
                        {
                            foreach (Entity ent in entCol.Entities)
                            {
                                lstEnquiryProdType.Add(new EnquiryProductType()
                                {
                                    EnquiryCategory = EnquiryCategory,
                                    EnquiryProductCode = ent.GetAttributeValue<string>("hil_sapcode"),
                                    EnquiryProductName = ent.GetAttributeValue<string>("name"),
                                });
                                JsonResponse = JsonSerializer.Serialize(lstEnquiryProdType);
                            }
                            context.OutputParameters["data"] = JsonResponse;
                            return;
                        }
                        else
                        {
                            lstEnquiryProdType.Add(new EnquiryProductType() { EnquiryProductCode = "ERROR", EnquiryProductName = "No Data found." });
                            JsonResponse = JsonSerializer.Serialize(lstEnquiryProdType);
                            context.OutputParameters["data"] = JsonResponse;
                            return;
                        }

                    }
                    else
                    {
                        lstEnquiryProdType.Add(new EnquiryProductType() { EnquiryProductCode = "ERROR", EnquiryProductName = "Enquiry Category not equal to 1 or 2" });
                        JsonResponse = JsonSerializer.Serialize(lstEnquiryProdType);
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                }
                else
                {
                    lstEnquiryProdType.Add(new EnquiryProductType() { EnquiryProductCode = "ERROR", EnquiryProductName = "D365 Service Unavailable" });
                    JsonResponse = JsonSerializer.Serialize(lstEnquiryProdType);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }

            }
            catch (Exception ex)
            {
                lstEnquiryProdType.Add(new EnquiryProductType() { EnquiryProductCode = "ERROR", EnquiryProductName = "D365 Internal Server Error : " + ex.Message });
                JsonResponse = JsonSerializer.Serialize(lstEnquiryProdType);
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
        }

        public class EnquiryProductType
        {
            public string EnquiryCategory { get; set; }
            public string EnquiryTypeName { get; set; }
            public string EnquiryTypeCode { get; set; }
            public string EnquiryProductCode { get; set; }
            public string EnquiryProductName { get; set; }

        }
    }
}
