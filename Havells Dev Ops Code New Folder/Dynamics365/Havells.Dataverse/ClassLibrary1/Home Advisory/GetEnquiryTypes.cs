using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.Customer_Asset.Home_Advisory
{
    public class GetEnquiryTypes : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            List<EnquiryType> lstEnquiryType = new List<EnquiryType>();

            string enquiryCategory = Convert.ToString(context.InputParameters["EnquiryCategory"]);
            if (string.IsNullOrWhiteSpace(enquiryCategory))
            {
                lstEnquiryType.Add(new EnquiryType() { EnquiryTypeCode = "ERROR", EnquiryTypeName = "EnquiryCategory is required." });
                string jsonErrorResult = JsonSerializer.Serialize(lstEnquiryType);
                context.OutputParameters["data"] = jsonErrorResult;
                return;
            }

            //	if (context.InputParameters.Contains("EnquiryCategory") && !string.IsNullOrWhiteSpace(context.InputParameters["EnquiryCategory"].ToString()))            
            //  string enquiryCategory = context.InputParameters["EnquiryCategory"].ToString();

            if (enquiryCategory.Length <= 9 && APValidate.NumericValue(enquiryCategory))
            {
                try
                {
                    string fetchXml = $@"
                            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_enquirytype'>
                            <attribute name='hil_enquirytypename' />
                            <attribute name='hil_enquirytypecode' />
                            <order attribute='hil_enquirytypename' descending='false' />
                            <filter type='and'>
                            <condition attribute='statecode' operator='eq' value='0' />
                            <condition attribute='hil_enquirycategory' operator='eq' value='{enquiryCategory}' />
                            </filter>
                            </entity>
                            </fetch>";

                    EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(fetchXml));

                    if (entCol.Entities.Count > 0)
                    {
                        foreach (Entity ent in entCol.Entities)
                        {
                            lstEnquiryType.Add(new EnquiryType()
                            {
                                EnquiryTypeName = ent.GetAttributeValue<string>("hil_enquirytypename"),
                                EnquiryTypeCode = ent.GetAttributeValue<int>("hil_enquirytypecode").ToString(),
                                EnquiryCategory = enquiryCategory
                            });
                        }
                    }
                    else
                    {
                        lstEnquiryType.Add(new EnquiryType() { EnquiryTypeCode = "ERROR", EnquiryTypeName = "No Data found." });
                    }

                    string jsonResult = JsonSerializer.Serialize(lstEnquiryType);
                    context.OutputParameters["data"] = jsonResult;
                }

                catch (Exception ex)
                {
                    lstEnquiryType.Add(new EnquiryType() { EnquiryTypeCode = "ERROR", EnquiryTypeName = "D365 Internal Server Error : " + ex.Message });
                    string jsonErrorResult = JsonSerializer.Serialize(lstEnquiryType);
                    context.OutputParameters["data"] = jsonErrorResult;
                }
            }
            else
            {
                lstEnquiryType.Add(new EnquiryType() { EnquiryTypeCode = "ERROR", EnquiryTypeName = "EnquiryCategory is not Valid" });
                string jsonErrorResult = JsonSerializer.Serialize(lstEnquiryType);
                context.OutputParameters["data"] = jsonErrorResult;
            }

            //else
            //{
            //    lstEnquiryType.Add(new EnquiryType() { EnquiryTypeCode = "ERROR", EnquiryTypeName = "EnquiryCategory is required " });
            //    string jsonErrorResult = JsonSerializer.Serialize(lstEnquiryType);
            //    context.OutputParameters["data"] = jsonErrorResult;
            //}
        }
    }
    public class EnquiryType
    {
        public string EnquiryCategory { get; set; }
        public string EnquiryTypeName { get; set; }
        public string EnquiryTypeCode { get; set; }
    }
}

