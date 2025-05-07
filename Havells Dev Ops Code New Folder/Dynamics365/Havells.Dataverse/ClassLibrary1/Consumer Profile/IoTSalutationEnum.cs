using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Havells.Dataverse.CustomConnector.Consumer_Profile
{
    public class IoTSalutationEnum : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            List<HashTableDTO> lstIoTSalutationEnum = new List<HashTableDTO>();
            string JsonResponse = string.Empty;
            try
            {
                var salutationData = GetSalutationEnum(service);
                JsonResponse = JsonConvert.SerializeObject(salutationData);
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
            catch (Exception ex)
            {
                var objIoTSalutationEnum = new HashTableDTO
                {
                    StatusCode = "500",
                    StatusDescription = "D365 Internal Server Error : " + ex.Message.ToLower()

                };
                lstIoTSalutationEnum.Add(objIoTSalutationEnum);
                JsonResponse = JsonConvert.SerializeObject(lstIoTSalutationEnum);
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
        }
        public List<HashTableDTO> GetSalutationEnum(IOrganizationService service)
        {
            HashTableDTO objIoTSalutationEnum;
            List<HashTableDTO> lstIoTSalutationEnum = new List<HashTableDTO>();
            try
            {
                if (service != null)
                {
                    var attributeRequest = new RetrieveAttributeRequest
                    {
                        EntityLogicalName = "contact",
                        LogicalName = "hil_salutation",
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
                }
                else
                {
                    objIoTSalutationEnum = new HashTableDTO { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    lstIoTSalutationEnum.Add(objIoTSalutationEnum);
                }
            }
            catch (Exception ex)
            {
                objIoTSalutationEnum = new HashTableDTO { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                lstIoTSalutationEnum.Add(objIoTSalutationEnum);
            }
            return lstIoTSalutationEnum;
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
