using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin.WorkOrder
{
    public class PreUpdateFieldValidations : IPlugin
    {
        static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "msdyn_workorder" && context.MessageName.ToUpper() == "UPDATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    string _previousValue = string.Empty, _currentValue = string.Empty;
                    //decimal _payableAmount = 0;
                    //decimal _receiptAmount = 0;
                    Entity preImageEntity = (Entity)context.PreEntityImages["image"];

                    EntityReference jobSubStatus = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("msdyn_substatus")).GetAttributeValue<EntityReference>("msdyn_substatus");
                    if (jobSubStatus.Id.ToString() == JobSubStatus.Canceled|| jobSubStatus.Id.ToString() == JobSubStatus.Closed || jobSubStatus.Id.ToString() == JobSubStatus.WorkDone)
                    {
                        if (entity.Contains("hil_customerref") && preImageEntity.Contains("hil_customerref"))
                        {
                            _currentValue = entity.GetAttributeValue<EntityReference>("hil_customerref").Id.ToString();
                            _previousValue = preImageEntity.GetAttributeValue<EntityReference>("hil_customerref").Id.ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_productcatsubcatmapping") && preImageEntity.Contains("hil_productcatsubcatmapping"))
                        {
                            _currentValue = entity.GetAttributeValue<EntityReference>("hil_productcatsubcatmapping").Id.ToString();
                            _previousValue = preImageEntity.GetAttributeValue<EntityReference>("hil_productcatsubcatmapping").Id.ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_productcategory") && preImageEntity.Contains("hil_productcategory"))
                        {
                            _currentValue = entity.GetAttributeValue<EntityReference>("hil_productcategory").Id.ToString();
                            _previousValue = preImageEntity.GetAttributeValue<EntityReference>("hil_productcategory").Id.ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_productsubcategory") && preImageEntity.Contains("hil_productsubcategory"))
                        {
                            _currentValue = entity.GetAttributeValue<EntityReference>("hil_productsubcategory").Id.ToString();
                            _previousValue = preImageEntity.GetAttributeValue<EntityReference>("hil_productsubcategory").Id.ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_consumertype") && preImageEntity.Contains("hil_consumertype"))
                        {
                            _currentValue = entity.GetAttributeValue<EntityReference>("hil_consumertype").Id.ToString();
                            _previousValue = preImageEntity.GetAttributeValue<EntityReference>("hil_consumertype").Id.ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_consumercategory") && preImageEntity.Contains("hil_consumercategory"))
                        {
                            _currentValue = entity.GetAttributeValue<EntityReference>("hil_consumercategory").Id.ToString();
                            _previousValue = preImageEntity.GetAttributeValue<EntityReference>("hil_consumercategory").Id.ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_natureofcomplaint") && preImageEntity.Contains("hil_natureofcomplaint"))
                        {
                            _currentValue = entity.GetAttributeValue<EntityReference>("hil_natureofcomplaint").Id.ToString();
                            _previousValue = preImageEntity.GetAttributeValue<EntityReference>("hil_natureofcomplaint").Id.ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_callsubtype") && preImageEntity.Contains("hil_callsubtype"))
                        {
                            _currentValue = entity.GetAttributeValue<EntityReference>("hil_callsubtype").Id.ToString();
                            _previousValue = preImageEntity.GetAttributeValue<EntityReference>("hil_callsubtype").Id.ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_quantity") && preImageEntity.Contains("hil_quantity"))
                        {
                            _currentValue = entity.GetAttributeValue<int>("hil_quantity").ToString();
                            _previousValue = preImageEntity.GetAttributeValue<int>("hil_quantity").ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_callertype") && preImageEntity.Contains("hil_callertype"))
                        {
                            _currentValue = entity.GetAttributeValue<OptionSetValue>("hil_callertype").Value.ToString();
                            _previousValue = preImageEntity.GetAttributeValue<OptionSetValue>("hil_callertype").Value.ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_newserialnumber") && preImageEntity.Contains("hil_newserialnumber"))
                        {
                            _currentValue = entity.GetAttributeValue<string>("hil_newserialnumber");
                            _previousValue = preImageEntity.GetAttributeValue<string>("hil_newserialnumber");
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_customercomplaintdescription") && preImageEntity.Contains("hil_customercomplaintdescription"))
                        {
                            _currentValue = entity.GetAttributeValue<string>("hil_customercomplaintdescription");
                            _previousValue = preImageEntity.GetAttributeValue<string>("hil_customercomplaintdescription");
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_expecteddeliverydate") && preImageEntity.Contains("hil_expecteddeliverydate"))
                        {
                            _currentValue = entity.GetAttributeValue<DateTime>("hil_expecteddeliverydate").ToString();
                            _previousValue = preImageEntity.GetAttributeValue<DateTime>("hil_expecteddeliverydate").ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_sourceofjob") && preImageEntity.Contains("hil_sourceofjob"))
                        {
                            _currentValue = entity.GetAttributeValue<OptionSetValue>("hil_sourceofjob").Value.ToString();
                            _previousValue = preImageEntity.GetAttributeValue<OptionSetValue>("hil_sourceofjob").Value.ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_owneraccount") && preImageEntity.Contains("hil_owneraccount"))
                        {
                            _currentValue = entity.GetAttributeValue<EntityReference>("hil_owneraccount").Id.ToString();
                            _previousValue = preImageEntity.GetAttributeValue<EntityReference>("hil_owneraccount").Id.ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("msdyn_customerasset") && preImageEntity.Contains("msdyn_customerasset"))
                        {
                            _currentValue = entity.GetAttributeValue<EntityReference>("msdyn_customerasset").Id.ToString();
                            _previousValue = preImageEntity.GetAttributeValue<EntityReference>("msdyn_customerasset").Id.ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_address") && preImageEntity.Contains("hil_address"))
                        {
                            _currentValue = entity.GetAttributeValue<EntityReference>("hil_address").Id.ToString();
                            _previousValue = preImageEntity.GetAttributeValue<EntityReference>("hil_address").Id.ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_mobilenumber") && preImageEntity.Contains("hil_mobilenumber"))
                        {
                            _currentValue = entity.GetAttributeValue<string>("hil_mobilenumber");
                            _previousValue = preImageEntity.GetAttributeValue<string>("hil_mobilenumber");
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_jobclosuredon") && preImageEntity.Contains("hil_jobclosuredon"))
                        {
                            _currentValue = entity.GetAttributeValue<DateTime>("hil_jobclosuredon").ToString();
                            _previousValue = preImageEntity.GetAttributeValue<DateTime>("hil_jobclosuredon").ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("msdyn_timeclosed") && preImageEntity.Contains("msdyn_timeclosed"))
                        {
                            _currentValue = entity.GetAttributeValue<DateTime>("msdyn_timeclosed").ToString();
                            _previousValue = preImageEntity.GetAttributeValue<DateTime>("msdyn_timeclosed").ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_tattime") && preImageEntity.Contains("hil_tattime"))
                        {
                            _currentValue = entity.GetAttributeValue<DateTime>("hil_tattime").ToString();
                            _previousValue = preImageEntity.GetAttributeValue<DateTime>("hil_tattime").ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied!!! Modification is not allowed.");
                            }
                        }
                        else if (entity.Contains("hil_schemecode") && preImageEntity.Contains("hil_schemecode"))
                        {
                            _currentValue = entity.GetAttributeValue<EntityReference>("hil_schemecode").Id.ToString();
                            _previousValue = preImageEntity.GetAttributeValue<EntityReference>("hil_schemecode").Id.ToString();
                            if (_currentValue != _previousValue)
                            {
                                throw new InvalidPluginExecutionException("Access denied. Modification is not allowed.");
                            }
                        }
                    }

                    //Check Receipt Amount and Payable Amount 
                    //if (entity.Contains("hil_receiptamount")) {
                    //    _receiptAmount = entity.GetAttributeValue<Money>("hil_receiptamount").Value;
                    //    Entity _entPayableAmount = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_actualcharges"));
                    //    _payableAmount = _entPayableAmount.GetAttributeValue<Money>("hil_actualcharges").Value;
                    //    if (_payableAmount < _receiptAmount) {
                    //        throw new InvalidPluginExecutionException(@"Data Validations !!! \n Payable Amount can't be less than Receipt Amount.");
                    //    }
                    //}
                    //if (entity.Contains("hil_actualcharges"))
                    //{
                    //    _payableAmount = entity.GetAttributeValue<Money>("hil_actualcharges").Value;
                    //    Entity _entReceiptAmount = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_receiptamount"));
                    //    _receiptAmount = _entReceiptAmount.GetAttributeValue<Money>("hil_receiptamount").Value;
                    //    if (_payableAmount < _receiptAmount)
                    //    {
                    //        throw new InvalidPluginExecutionException(@"Data Validations !!! \n Payable Amount can't be less than Receipt Amount.");
                    //    }
                    //}
                }
            }
            catch (InvalidPluginExecutionException e)
            {
                throw new InvalidPluginExecutionException("Data Validations !!! \n" + e.Message);
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
            #endregion
        }
    }
}
