using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Text.Json;


namespace Havells.Dataverse.CustomConnector.AMC
{
    public class GetWarrantyContent : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            string JsonResponse = "";
            bool IsValidRequest = true;
            StringBuilder errorMessage = new StringBuilder();
            string jsonString = Convert.ToString(context.InputParameters["reqdata"]);
            var data = JsonSerializer.Deserialize<reqdata>(jsonString);
            string SourceType = data.SourceType;
            if (string.IsNullOrWhiteSpace(SourceType))
            {
                JsonResponse = JsonSerializer.Serialize(new RequestStatus
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = "Please enter the Source Type."
                });
                context.OutputParameters["data"] = JsonResponse;
                return;

            }
            if (SourceType != "6")
            {
                JsonResponse = JsonSerializer.Serialize(new RequestStatus
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = "Please enter valid Source Type."
                });
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
            string LoginUserId = Convert.ToString(context.InputParameters["LoginUserId"]);
            string UserToken = Convert.ToString(context.InputParameters["UserToken"]);
            if (LoginUserId.Length != 10)
            {
                JsonResponse = JsonSerializer.Serialize(new RequestStatus
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = "Invalid mobile number"
                });
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
            if (!IsValidRequest)
            {
                JsonResponse = JsonSerializer.Serialize(new RequestStatus
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = errorMessage.ToString()
                });
                context.OutputParameters["data"] = JsonResponse;
                return;
            }

            JsonResponse = JsonSerializer.Serialize((GetWarrantyContentdata(service, SourceType, LoginUserId)));
            context.OutputParameters["data"] = JsonResponse;
        }

        public WarrantyContentRes GetWarrantyContentdata(IOrganizationService service, string SourceType, string LoginUserId)
        {
            WarrantyContentRes warrantyContent = new WarrantyContentRes();
            EntityCollection entcoll;
            try
            {
                if (service != null)
                {
                    string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='hil_productcatalog'>
                                    <attribute name='hil_productcatalogid' />
                                    <attribute name='hil_plantclink' />
                                    <attribute name='hil_productcode' />
                                    <order attribute='hil_productcode' descending='false' />
                                    <filter type='and'>
                                        <condition attribute='hil_eligibleforamc' operator='eq' value='1' />
                                    </filter>
                                    <link-entity name='product' from='productid' to='hil_productcode' link-type='inner' alias='ag'>
                                        <filter type='and'>
                                            <condition attribute='hil_hierarchylevel' operator='eq' value='2' />
                                        </filter>
                                    </link-entity>
                                    </entity>
                                    </fetch>";
                    entcoll = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entcoll.Entities.Count > 0)
                    {
                        List<AMCCatagory> lstAMCCatagory = new List<AMCCatagory>();
                        foreach (Entity ent in entcoll.Entities)
                        {
                            AMCCatagory objAMCCatagory = new AMCCatagory();
                            objAMCCatagory.CategoryId = ent.Id;
                            objAMCCatagory.CategoryName = ent.Contains("hil_productcode") ? ent.GetAttributeValue<EntityReference>("hil_productcode").Name : null;
                            objAMCCatagory.Icon = ent.Contains("hil_plantclink") ? ent.GetAttributeValue<string>("hil_plantclink") : null;
                            lstAMCCatagory.Add(objAMCCatagory);
                        }
                        warrantyContent.AMCCatagories = lstAMCCatagory;
                    }
                    _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='hil_warrantymediagallery'>
                                    <attribute name='hil_warrantymediagalleryid' />
                                    <attribute name='hil_name' />
                                    <attribute name='hil_content' />
                                    <attribute name='hil_imagepath' />
                                    <attribute name='hil_displayindex' />
                                    <attribute name='hil_category' />
                                        <order attribute='hil_category' descending='false' />
                                        <order attribute='hil_displayindex' descending='false' />
                                    </entity>
                                    </fetch>";
                    entcoll = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entcoll.Entities.Count > 0)
                    {
                        List<WarrantyMediaContent> lstWarrantyMediaContent = new List<WarrantyMediaContent>();
                        List<WarrantyDiscountBanner> lstWarrantyDiscountBanner = new List<WarrantyDiscountBanner>();
                        foreach (Entity ent in entcoll.Entities)
                        {
                            int Category = ent.Contains("hil_category") ? ent.GetAttributeValue<OptionSetValue>("hil_category").Value : 0;
                            if (Category == 1) //Banner
                            {
                                WarrantyDiscountBanner objWarrantyDiscountBanner = new WarrantyDiscountBanner();
                                objWarrantyDiscountBanner.Index = ent.Contains("hil_displayindex") ? ent.GetAttributeValue<int>("hil_displayindex") : 0;
                                objWarrantyDiscountBanner.URL = ent.Contains("hil_imagepath") ? ent.GetAttributeValue<string>("hil_imagepath") : null;
                                lstWarrantyDiscountBanner.Add(objWarrantyDiscountBanner);
                            }
                            else if (Category == 2) // Category
                            {
                                WarrantyMediaContent objWarrantyMediaContent = new WarrantyMediaContent();
                                objWarrantyMediaContent.Icon = ent.Contains("hil_imagepath") ? ent.GetAttributeValue<string>("hil_imagepath") : null;
                                objWarrantyMediaContent.Content = ent.Contains("hil_content") ? ent.GetAttributeValue<string>("hil_content") : null;
                                lstWarrantyMediaContent.Add(objWarrantyMediaContent);
                            }
                        }
                        warrantyContent.WarrantyMediaContents = lstWarrantyMediaContent;
                        warrantyContent.WarrantyDiscountBanners = lstWarrantyDiscountBanner;
                    }
                    RequestStatus requestStatus = new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.OK,
                        Message = "Success",
                    };
                    warrantyContent.StatusCode = requestStatus.StatusCode;
                    warrantyContent.Message = requestStatus.Message;
                    return (warrantyContent);
                }
                else
                {
                    var requestStatus = new RequestStatus()
                    {
                        StatusCode = (int)HttpStatusCode.ServiceUnavailable,
                        Message = "D365 service unavailable.",
                    };
                    warrantyContent.StatusCode = requestStatus.StatusCode;
                    warrantyContent.Message = requestStatus.Message;
                    return (warrantyContent);
                }
            }
            catch (Exception ex)
            {
                RequestStatus requestStatus = new RequestStatus()
                {
                    StatusCode = (int)HttpStatusCode.BadRequest,
                    Message = "D365 internal server error :",
                };
                warrantyContent.StatusCode = requestStatus.StatusCode;
                warrantyContent.Message = requestStatus.Message;
                return (warrantyContent);
            }
        }

    }
    public class WarrantyContentRes : TokenExpires
    {
        public List<AMCCatagory> AMCCatagories { get; set; }
        public List<WarrantyMediaContent> WarrantyMediaContents { get; set; }
        public List<WarrantyDiscountBanner> WarrantyDiscountBanners { get; set; }
    }
    public class AMCCatagory
    {
        public string Icon { get; set; }
        public string CategoryName { get; set; }
        public Guid CategoryId { get; set; }
    }
    public class WarrantyMediaContent
    {
        public string Icon { get; set; }
        public string Content { get; set; }
    }
    public class WarrantyDiscountBanner
    {
        public int Index { get; set; }
        public string URL { get; set; }
    }
    public class TokenExpires
    {
        public int StatusCode { get; set; }
        public string Message { get; set; }
    }
    public class RequestStatus
    {
        public int StatusCode { get; set; }
        public string Message { get; set; }
    }

    public class reqdata
    {
        public string SourceType { get; set; }
    }
    public class RootObject
    {
        public reqdata reqdata { get; set; }
    }
}
