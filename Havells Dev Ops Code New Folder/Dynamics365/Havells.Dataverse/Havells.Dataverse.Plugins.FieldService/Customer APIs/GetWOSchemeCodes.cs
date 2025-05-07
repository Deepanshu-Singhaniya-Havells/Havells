using Havells.Dataverse.Plugins.FieldService.Models;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace Havells.Dataverse.Plugins.FieldService.Customer_APIs
{
    public class GetWOSchemeCodes : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            WOSchemesResult wOSchemesResult = new WOSchemesResult()
            {
                lstWOSchemes = new List<WOSchemes>()
            };
            string JsonResponse = "";
            try
            {               
                if (context.InputParameters.Contains("WorkOrderId") && context.InputParameters["WorkOrderId"] is string)
                {
                    #region Validate Params
                    Guid WorkOrderId = Guid.Empty;
                    bool isValidWorkOrderId = Guid.TryParse(context.InputParameters["WorkOrderId"].ToString(), out WorkOrderId);
                    if (!isValidWorkOrderId)
                    {
                        JsonResponse = JsonConvert.SerializeObject(new WOSchemesResult
                        {
                            result = new CustomerResult
                            {
                                ResultStatus = false,
                                ResultMessage = "Work Order GuId is required."
                            }
                        });
                        context.OutputParameters["Response"] = JsonResponse;
                        return;
                    }
                    #endregion
                    DateTime _PurchaseDate;
                    Guid _CallSubType = Guid.Empty;
                    Guid _SalesOffice = Guid.Empty;
                    Guid _ProdSubCatg = Guid.Empty;
                    OptionSetValue _CallerType = null;
                    DateTime _CreatedOn;
                    if (service != null)
                    {
                        #region Get Work Order Details
                        Entity entWO = service.Retrieve("msdyn_workorder", WorkOrderId, new ColumnSet("hil_purchasedate", "hil_callsubtype", "createdon", "hil_salesoffice", "hil_productsubcategory", "hil_callertype"));
                        if (entWO != null)
                        {
                            if (entWO.Attributes.Contains("hil_callertype"))
                            {
                                _CallerType = entWO.GetAttributeValue<OptionSetValue>("hil_callertype");
                            }
                            if (entWO.Attributes.Contains("hil_purchasedate"))
                            {
                                _PurchaseDate = entWO.GetAttributeValue<DateTime>("hil_purchasedate").AddMinutes(330);
                            }
                            else
                            {
                                _PurchaseDate = new DateTime(1900, 1, 1);
                            }
                            if (entWO.Attributes.Contains("hil_callsubtype"))
                            {
                                _CallSubType = entWO.GetAttributeValue<EntityReference>("hil_callsubtype").Id;
                            }
                            else
                            {
                                JsonResponse = JsonConvert.SerializeObject(new WOSchemesResult
                                {
                                    result = new CustomerResult
                                    {
                                        ResultStatus = false,
                                        ResultMessage = "Call Sub Type is not defined in Work Order."
                                    }
                                });
                                context.OutputParameters["Response"] = JsonResponse;
                                return;
                            }
                            if (entWO.Attributes.Contains("hil_salesoffice"))
                            {
                                _SalesOffice = entWO.GetAttributeValue<EntityReference>("hil_salesoffice").Id;
                            }
                            else
                            {
                                JsonResponse = JsonConvert.SerializeObject(new WOSchemesResult
                                {
                                    result = new CustomerResult
                                    {
                                        ResultStatus = false,
                                        ResultMessage = "Sales Office is not defined in Work Order."
                                    }
                                });
                                context.OutputParameters["Response"] = JsonResponse;
                                return;
                            }
                            if (entWO.Attributes.Contains("hil_productsubcategory"))
                            {
                                _ProdSubCatg = entWO.GetAttributeValue<EntityReference>("hil_productsubcategory").Id;
                            }
                            else
                            {
                                JsonResponse = JsonConvert.SerializeObject(new WOSchemesResult
                                {
                                    result = new CustomerResult
                                    {
                                        ResultStatus = false,
                                        ResultMessage = "Product Sub Category is not defined in Work Order."
                                    }
                                });
                                context.OutputParameters["Response"] = JsonResponse;
                                return;
                            }
                            _CreatedOn = entWO.GetAttributeValue<DateTime>("createdon").AddMinutes(330);
                            string _purchaseDateValue = _PurchaseDate.Year.ToString() + "-" + _PurchaseDate.Month.ToString().PadLeft(2, '0') + "-" + _PurchaseDate.Day.ToString().PadLeft(2, '0');
                            string _createdOnValue = _CreatedOn.Year.ToString() + "-" + _CreatedOn.Month.ToString().PadLeft(2, '0') + "-" + _CreatedOn.Day.ToString().PadLeft(2, '0');

                            string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                                <entity name='hil_schemeincentive'>
                                                <attribute name='hil_schemeincentiveid' />
                                                <attribute name='hil_name' />
                                                <order attribute='hil_name' descending='false' />
                                                <filter type='and'>
                                                    <condition attribute='hil_schemeexpirydate' operator='on-or-after' value='" + _createdOnValue + @"' />
                                                    <condition attribute='hil_fromdate' operator='on-or-before' value='" + _purchaseDateValue + @"' />
                                                    <condition attribute='hil_todate' operator='on-or-after' value='" + _purchaseDateValue + @"' />
                                                    <condition attribute='hil_callsubtype' operator='eq' value='{" + _CallSubType + @"}' />
                                                    <condition attribute='hil_productsubcategory' operator='eq' value='{" + _ProdSubCatg + @"}' />
                                                    <condition attribute='hil_salesoffice' operator='in'>
                                                <value >{" + _SalesOffice + @"}</value>
                                                </condition>
                                                <condition attribute='statecode' operator='eq' value='0' />";
                            if (_CallerType != null)
                                fetchXML = fetchXML + @"<condition attribute='hil_callertype' operator='eq' value='" + _CallerType.Value.ToString() + @"' />";
                            fetchXML = fetchXML + @"</filter></entity></fetch>";

                            EntityCollection entColConsumer = service.RetrieveMultiple(new FetchExpression(fetchXML));
                            if (entColConsumer.Entities.Count > 0)
                            {
                                List<WOSchemes> lstWOSchemes = new List<WOSchemes>();
                                foreach (Entity ent in entColConsumer.Entities)
                                {
                                    WOSchemes wOSchemes = new WOSchemes();
                                    if (ent.Attributes.Contains("hil_name"))
                                    {
                                        wOSchemes.SchemeId = ent.Id;
                                        wOSchemes.SchemeName = ent.GetAttributeValue<string>("hil_name");
                                    }
                                    lstWOSchemes.Add(wOSchemes);
                                }
                                wOSchemesResult.lstWOSchemes = lstWOSchemes;
                            }
                            else
                            {
                                fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='hil_schemeincentive'>
                                            <attribute name='hil_schemeincentiveid' />
                                            <attribute name='hil_name' />
                                            <order attribute='hil_name' descending='false' />
                                            <filter type='and'>
                                                <condition attribute='hil_schemeexpirydate' operator='on-or-after' value='" + _createdOnValue + @"' />
                                                <condition attribute='hil_fromdate' operator='on-or-before' value='" + _purchaseDateValue + @"' />
                                                <condition attribute='hil_todate' operator='on-or-after' value='" + _purchaseDateValue + @"' />
                                                <condition attribute='hil_callsubtype' operator='eq' value='{" + _CallSubType + @"}' />
                                                <condition attribute='hil_productsubcategory' operator='eq' value='{" + _ProdSubCatg + @"}' />
                                                <condition attribute='hil_salesoffice' operator='in'>
                                            <value >{90503976-8FD1-EA11-A813-000D3AF0563C}</value>
                                            </condition>
                                            <condition attribute='statecode' operator='eq' value='0' />";
                                if (_CallerType != null)
                                    fetchXML = fetchXML + @"<condition attribute='hil_callertype' operator='eq' value='" + _CallerType.Value.ToString() + @"' />";
                                fetchXML = fetchXML + @"</filter></entity></fetch>";

                                entColConsumer = service.RetrieveMultiple(new FetchExpression(fetchXML));
                                if (entColConsumer.Entities.Count > 0)
                                {
                                    List<WOSchemes> lstWOSchemes = new List<WOSchemes>();
                                    foreach (Entity ent in entColConsumer.Entities)
                                    {
                                        WOSchemes wOSchemes = new WOSchemes();
                                        if (ent.Attributes.Contains("hil_name"))
                                        {
                                            wOSchemes.SchemeId = ent.Id;
                                            wOSchemes.SchemeName = ent.GetAttributeValue<string>("hil_name");
                                        }
                                        lstWOSchemes.Add(wOSchemes);
                                    }
                                    wOSchemesResult.lstWOSchemes = lstWOSchemes;
                                }
                            }
                        }
                        else
                        {
                            JsonResponse = JsonConvert.SerializeObject(new WOSchemesResult
                            {
                                result = new CustomerResult
                                {
                                    ResultStatus = false,
                                    ResultMessage = "Work Order Id does not exist !!! Something went wrong."
                                }
                            });
                            context.OutputParameters["Response"] = JsonResponse;
                            return;
                        }
                        if (wOSchemesResult.lstWOSchemes.Count > 0)
                        {
                            wOSchemesResult.result = new CustomerResult
                            {
                                ResultStatus = true,
                                ResultMessage = "OK"
                            };
                        }
                        else
                        {
                            wOSchemesResult.result = new CustomerResult
                            {
                                ResultStatus = true,
                                ResultMessage = "No Record found."
                            };
                        }
                        context.OutputParameters["Response"] = JsonConvert.SerializeObject(wOSchemesResult);
                        #endregion
                    }
                    else
                    {
                        JsonResponse = JsonConvert.SerializeObject(new WOSchemesResult
                        {
                            result = new CustomerResult
                            {
                                ResultStatus = false,
                                ResultMessage = "D365 Service Unavailable"
                            }
                        });
                        context.OutputParameters["Response"] = JsonResponse;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                JsonResponse = JsonConvert.SerializeObject(new WOSchemesResult
                {
                    result = new CustomerResult
                    {
                        ResultStatus = false,
                        ResultMessage = ex.Message
                    }
                });
                context.OutputParameters["Response"] = JsonResponse;
                return;
            }
        }
    }
}