using Havells.Dataverse.Plugins.FieldService.Models;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;

namespace Havells.Dataverse.Plugins.FieldService.Customer_APIs
{
    public class ValidateAMCReceiptAmount : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            DateTime _invoiceDate = new DateTime(1900, 1, 1);
            string JsonResponse = "";
            try
            {
                if (context.InputParameters.Contains("JobId") && context.InputParameters["JobId"] is string
                    && context.InputParameters.Contains("ReceiptAmount") && context.InputParameters["ReceiptAmount"] is decimal)
                {
                    Guid JobId = Guid.Empty;
                    bool isValidJobId = Guid.TryParse(context.InputParameters["JobId"].ToString(), out JobId);
                    decimal ReceiptAmount = Convert.ToDecimal(context.InputParameters["ReceiptAmount"].ToString());

                    if (!isValidJobId)
                    {
                        JsonResponse = JsonConvert.SerializeObject(new CustomerResultWithMessageType
                        {
                            ResultMessageType = "INFO",
                            ResultStatus = false,
                            ResultMessage = "Invalid Job Id."
                        });
                        context.OutputParameters["Response"] = JsonResponse;
                        return;
                    }
                    if (ReceiptAmount <= 0)
                    {
                        JsonResponse = JsonConvert.SerializeObject(new CustomerResultWithMessageType
                        {
                            ResultMessageType = "INFO",
                            ResultStatus = false,
                            ResultMessage = "Receipt Amount should be greater than zero."
                        });
                        context.OutputParameters["Response"] = JsonResponse;
                        return;
                    }
                    if (service != null)
                    {
                        string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='msdyn_workorderincident'>
                                    <attribute name='msdyn_name' />
                                    <filter type='and'>
                                        <condition attribute='msdyn_workorder' operator='eq' value='{JobId}' />
                                        <condition attribute='statecode' operator='eq' value='0' />
                                    </filter>
                                    <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='msdyn_customerasset' visible='false' link-type='outer' alias='ca'>
                                        <attribute name='hil_invoicedate' />
                                    </link-entity>
                                    <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' visible='false' link-type='outer' alias='wo'>
                                        <attribute name='createdon' />
                                        <attribute name='hil_actualcharges' />
                                        <attribute name='hil_callsubtype' />
                                        <attribute name='hil_productcategory' />
                                    </link-entity>
                                    </entity>
                                    </fetch>";

                        EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(fetchXML));
                        if (entCol.Entities.Count > 0)
                        {
                            if (!entCol.Entities[0].Attributes.Contains("wo.hil_callsubtype"))
                            {
                                JsonResponse = JsonConvert.SerializeObject(new CustomerResultWithMessageType
                                {
                                    ResultMessageType = "INFO",
                                    ResultStatus = false,
                                    ResultMessage = "Call Subtype is required."
                                });
                                context.OutputParameters["Response"] = JsonResponse;
                                return;
                            }
                            if (entCol.Entities[0].Attributes.Contains("ca.hil_invoicedate"))
                            {
                                _invoiceDate = (DateTime)(entCol.Entities[0].GetAttributeValue<AliasedValue>("ca.hil_invoicedate").Value);
                            }
                            EntityReference entTemp = (EntityReference)entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.hil_callsubtype").Value;
                            EntityReference entProdCatg = (EntityReference)entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.hil_productcategory").Value;

                            if (entTemp.Id != new Guid("55A71A52-3C0B-E911-A94E-000D3AF06CD4"))
                            {
                                JsonResponse = JsonConvert.SerializeObject(new CustomerResultWithMessageType
                                {
                                    ResultMessageType = "INFO",
                                    ResultStatus = true,
                                    ResultMessage = "OK"
                                });
                                context.OutputParameters["Response"] = JsonResponse;
                                return;
                            }

                            decimal _payableAmount = 0;
                            if (entCol.Entities[0].Attributes.Contains("wo.hil_actualcharges"))
                            {
                                _payableAmount = ((Money)(entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.hil_actualcharges").Value)).Value;
                            }
                            else
                            {
                                JsonResponse = JsonConvert.SerializeObject(new CustomerResultWithMessageType
                                {
                                    ResultMessageType = "INFO",
                                    ResultStatus = true,
                                    ResultMessage = "OK"
                                });
                                context.OutputParameters["Response"] = JsonResponse;
                                return;
                            }
                            DateTime _jobDate = (DateTime)entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.createdon").Value;
                            //_asOn Definition :: AMC Job Create date is concidered for Applying Discount rate becoz Product ageing also calculated from AMC Job Create Date
                            string _asOn = _jobDate.Year.ToString() + "-" + _jobDate.Month.ToString().PadLeft(2, '0') + "-" + _jobDate.Day.ToString().PadLeft(2, '0');
                            int _dayDiff = Convert.ToInt32(Math.Round((_jobDate - _invoiceDate).TotalDays, 0));
                            if (_dayDiff < 0)
                            {
                                JsonResponse = JsonConvert.SerializeObject(new CustomerResultWithMessageType
                                {
                                    ResultMessageType = "WARNING",
                                    ResultStatus = false,
                                    ResultMessage = "Product Age is -4 days. Job is created prior to Asset Invoice Date."
                                });
                                context.OutputParameters["Response"] = JsonResponse;
                                return;
                            }
                            decimal _stdDiscPer = 0;
                            decimal _spcDiscPer = 0;
                            decimal _stdDiscAmount = 0;
                            decimal _spcDiscAmount = 0;
                            //03B5A2D6-CC64-ED11-9562-6045BDAC526A - AMC Sale - FSM (Source)
                            fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_amcdiscountmatrix'>
                                <attribute name='hil_amcdiscountmatrixid' />
                                <attribute name='hil_name' />
                                <attribute name='createdon' />
                                <attribute name='hil_discounttype' />
                                <attribute name='hil_discper' />
                                <order attribute='hil_name' descending='false' />
                                <filter type='and'>
                                    <condition attribute='statecode' operator='eq' value='0' />
                                    <condition attribute='hil_appliedto' operator='eq' value='{{03B5A2D6-CC64-ED11-9562-6045BDAC526A}}' />
                                    <condition attribute='hil_productaegingstart' operator='le' value='{_dayDiff.ToString()}' />
                                    <condition attribute='hil_productageingend' operator='ge' value='{_dayDiff.ToString()}' />
                                    <condition attribute='hil_validfrom' operator='on-or-before' value='{_asOn}' />
                                    <condition attribute='hil_validto' operator='on-or-after' value='{_asOn}' />
                                    <condition attribute='hil_productcategory' operator='eq' value='{entProdCatg.Id}' />
                                </filter>
                                </entity>
                                </fetch>";

                            EntityCollection entCol1 = service.RetrieveMultiple(new FetchExpression(fetchXML));
                            if (entCol1.Entities.Count > 0)
                            {
                                foreach (Entity ent in entCol1.Entities)
                                {
                                    if (ent.GetAttributeValue<OptionSetValue>("hil_discounttype").Value == 1)
                                    {
                                        _stdDiscPer = ent.GetAttributeValue<decimal>("hil_discper");
                                    }
                                    else
                                    {
                                        _spcDiscPer = ent.GetAttributeValue<decimal>("hil_discper");
                                    }
                                }
                                _stdDiscAmount = Math.Round((_payableAmount - (_payableAmount * _stdDiscPer) / 100), 2); //Max Limit (90)
                                _spcDiscAmount = Math.Round(_payableAmount - (_payableAmount * (_stdDiscPer + _spcDiscPer)) / 100, 2); //Min Limit (85)
                                if (ReceiptAmount >= _spcDiscAmount && ReceiptAmount < _stdDiscAmount)
                                {
                                    decimal _additionaldisc = Math.Round(_stdDiscAmount - Convert.ToDecimal(ReceiptAmount), 2);

                                    JsonResponse = JsonConvert.SerializeObject(new CustomerResultWithMessageType
                                    {
                                        ResultMessageType = "CONFIRMATION",
                                        ResultStatus = false,
                                        ResultMessage = "To offer additional discount (Rs. " + _additionaldisc.ToString() + ") above Standard Discount, you need to take BSH approval. Click 'Yes' if approval already taken Or Click 'No'."
                                    });
                                    context.OutputParameters["Response"] = JsonResponse;
                                    return;
                                }
                                else if (ReceiptAmount < _spcDiscAmount)
                                {
                                    JsonResponse = JsonConvert.SerializeObject(new CustomerResultWithMessageType
                                    {
                                        ResultMessageType = "WARNING",
                                        ResultStatus = false,
                                        ResultMessage = "As per AMC discount policy, you are allowed to collect minimum Rs. " + _stdDiscAmount.ToString() + "."
                                    });
                                    context.OutputParameters["Response"] = JsonResponse;
                                    return;
                                }
                                else
                                {
                                    JsonResponse = JsonConvert.SerializeObject(new CustomerResultWithMessageType
                                    {
                                        ResultMessageType = "INFO",
                                        ResultStatus = false,
                                        ResultMessage = "OK"
                                    });
                                    context.OutputParameters["Response"] = JsonResponse;
                                    return;
                                }
                            }
                            else
                            {
                                if (_payableAmount != ReceiptAmount)
                                {
                                    JsonResponse = JsonConvert.SerializeObject(new CustomerResultWithMessageType
                                    {
                                        ResultMessageType = "WARNING",
                                        ResultStatus = false,
                                        ResultMessage = "No AMC Discount Policy is defined in System !!! Receipt Amount can't be less than Payable Amount."
                                    });
                                    context.OutputParameters["Response"] = JsonResponse;
                                    return;
                                }
                                else
                                {
                                    JsonResponse = JsonConvert.SerializeObject(new CustomerResultWithMessageType
                                    {
                                        ResultMessageType = "INFO",
                                        ResultStatus = false,
                                        ResultMessage = "OK"
                                    });
                                    context.OutputParameters["Response"] = JsonResponse;
                                    return;
                                }
                            }
                        }
                        else
                        {
                            JsonResponse = JsonConvert.SerializeObject(new CustomerResultWithMessageType
                            {
                                ResultMessageType = "WARNING",
                                ResultStatus = false,
                                ResultMessage = "No Work Order Incident found."
                            });
                            context.OutputParameters["Response"] = JsonResponse;
                            return;
                        }
                    }
                    else
                    {
                        JsonResponse = JsonConvert.SerializeObject(new CustomerResultWithMessageType
                        {
                            ResultMessageType = "WARNING",
                            ResultStatus = false,
                            ResultMessage = "D365 Service Unavailable."
                        });
                        context.OutputParameters["Response"] = JsonResponse;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                JsonResponse = JsonConvert.SerializeObject(new CustomerResultWithMessageType
                {
                    ResultMessageType = "WARNING",
                    ResultStatus = false,
                    ResultMessage = ex.Message
                });
                context.OutputParameters["Response"] = JsonResponse;
                return;
            }
        }
    }
}
