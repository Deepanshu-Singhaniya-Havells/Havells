using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Havells.Dataverse.CustomConnector.Customer_Asset
{
    public class IoTInstallationLocationEnum : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            HashTableDTO objIoTSalutationEnum;
            List<HashTableDTO> lstIoTSalutationEnum = new List<HashTableDTO>();
            string JsonResponse = "";
            try
            {
                if (service != null)
                {
                    var attributeRequest = new RetrieveAttributeRequest
                    {
                        EntityLogicalName = "msdyn_customerasset",
                        LogicalName = "hil_product",
                        RetrieveAsIfPublished = true
                    };

                    var attributeResponse = (RetrieveAttributeResponse)service.Execute(attributeRequest);
                    var attributeMetadata = (EnumAttributeMetadata)attributeResponse.AttributeMetadata;

                    var optionList = (from o in attributeMetadata.OptionSet.Options
                                      select new { Value = o.Value, Text = o.Label.UserLocalizedLabel.Label }).ToList();
                    foreach (var option in optionList)
                    {
                        lstIoTSalutationEnum.Add(new HashTableDTO() { Value = option.Value, Label = option.Text, StatusCode = "200", StatusDescription = "OK" });
                    }
                    JsonResponse = JsonConvert.SerializeObject(lstIoTSalutationEnum);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                else
                {
                    objIoTSalutationEnum = new HashTableDTO { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    lstIoTSalutationEnum.Add(objIoTSalutationEnum);
                    JsonResponse = JsonConvert.SerializeObject(lstIoTSalutationEnum);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
            }
            catch (Exception ex)
            {
                objIoTSalutationEnum = new HashTableDTO { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                lstIoTSalutationEnum.Add(objIoTSalutationEnum);
                JsonResponse = JsonConvert.SerializeObject(lstIoTSalutationEnum);
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
        }
        public class HashTableDTO
        {
            public string Label { get; set; }
            public int? Value { get; set; }
            public string Extension { get; set; }
            public string StatusCode { get; set; }
            public string StatusDescription { get; set; }
        }

    }
}