using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.NatureOfComplaint
{
    public class NatureOfComplaintByProdSubcategory : IPlugin
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
            Guid ProductSubCategoryId = Guid.Empty;
            string Source = null;
            List<IoTNatureofComplaint> lstNatureOfComplaint = new List<IoTNatureofComplaint>();

            if (context.InputParameters.Contains("ProductSubCategoryId") && context.InputParameters["ProductSubCategoryId"] is string)
            {
                bool isValidProductSubCategoryId = Guid.TryParse(context.InputParameters["ProductSubCategoryId"].ToString(), out ProductSubCategoryId);
                if (!isValidProductSubCategoryId)
                {
                    lstNatureOfComplaint.Add(new IoTNatureofComplaint
                    {
                        StatusCode = "204",
                        StatusDescription = "Invalid Product SubCategory GUID."
                    });
                    JsonResponse = JsonSerializer.Serialize(lstNatureOfComplaint);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }

                ProductSubCategoryId = new Guid(context.InputParameters["ProductSubCategoryId"].ToString());
                if (context.InputParameters.Contains("Source"))
                {
                    if (!string.IsNullOrWhiteSpace(Convert.ToString(context.InputParameters["Source"])))
                        Source = context.InputParameters["Source"].ToString();
                }
                _tracingService.Trace("JsonResponse");
                JsonResponse = JsonSerializer.Serialize(GetNatureOfComplaintByProdSubcategory(service, ProductSubCategoryId, Source));
                context.OutputParameters["data"] = JsonResponse;
            }
        }

        public List<IoTNatureofComplaint> GetNatureOfComplaintByProdSubcategory(IOrganizationService service, Guid ProductSubCategoryId, string Source)
        {
            IoTNatureofComplaint objNatureOfComplaint;
            List<IoTNatureofComplaint> lstNatureOfComplaint = new List<IoTNatureofComplaint>();
            EntityCollection entcoll;
            try
            {
                if (service != null)
                {
                    string fetchXml = string.Empty;
                    if (Source == null)
                    {
                        fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                            <entity name='hil_natureofcomplaint'>
                            <attribute name='hil_name' />
                            <attribute name='hil_natureofcomplaintid' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_relatedproduct' operator='eq' value='{" + ProductSubCategoryId + @"}' />
                            </filter>
                            </entity>
                            </fetch>";
                    }
                    else
                    {
                        fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                            <entity name='hil_natureofcomplaint'>
                            <attribute name='hil_name' />
                            <attribute name='hil_natureofcomplaintid' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_relatedproduct' operator='eq' value='{" + ProductSubCategoryId + @"}' />
                                <condition attribute='hil_callsubtype' operator='in'>
                                <value uiname='AMC Call' uitype='hil_callsubtype'>{55A71A52-3C0B-E911-A94E-000D3AF06CD4}</value>
                                <value uiname='Breakdown' uitype='hil_callsubtype'>{6560565A-3C0B-E911-A94E-000D3AF06CD4}</value>
                                <value uiname='Demo' uitype='hil_callsubtype'>{AE1B2B71-3C0B-E911-A94E-000D3AF06CD4}</value>
                                <value uiname='Installation' uitype='hil_callsubtype'>{E3129D79-3C0B-E911-A94E-000D3AF06CD4}</value>
                                <value uiname='PMS' uitype='hil_callsubtype'>{E2129D79-3C0B-E911-A94E-000D3AF06CD4}</value>
                                </condition>
                            </filter>
                            </entity>
                            </fetch>";
                    }
                    entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (entcoll.Entities.Count == 0)
                    {
                        objNatureOfComplaint = new IoTNatureofComplaint { StatusCode = "204", StatusDescription = "No Nature of Complaint is mapped with Serial Number." };
                        lstNatureOfComplaint.Add(objNatureOfComplaint);
                    }
                    else
                    {
                        foreach (Entity ent in entcoll.Entities)
                        {
                            objNatureOfComplaint = new IoTNatureofComplaint();
                            objNatureOfComplaint.Guid = ent.GetAttributeValue<Guid>("hil_natureofcomplaintid");
                            objNatureOfComplaint.Name = ent.GetAttributeValue<string>("hil_name");
                            objNatureOfComplaint.ProductSubCategoryId = ProductSubCategoryId;
                            objNatureOfComplaint.SerialNumber = "";
                            objNatureOfComplaint.StatusCode = "200";
                            objNatureOfComplaint.StatusDescription = "OK";
                            lstNatureOfComplaint.Add(objNatureOfComplaint);
                        }
                    }
                    return lstNatureOfComplaint;
                }
                else
                {
                    objNatureOfComplaint = new IoTNatureofComplaint { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    lstNatureOfComplaint.Add(objNatureOfComplaint);
                    return lstNatureOfComplaint;
                }
            }
            catch (Exception ex)
            {
                objNatureOfComplaint = new IoTNatureofComplaint { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                lstNatureOfComplaint.Add(objNatureOfComplaint);
                return lstNatureOfComplaint;
            }

        }
    }
    public class IoTNatureofComplaint
    {
        public string SerialNumber { get; set; }
        public Guid ProductSubCategoryId { get; set; }
        public string Name { get; set; }
        public Guid Guid { get; set; }
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
        public string Source { get; set; }
    }
}
