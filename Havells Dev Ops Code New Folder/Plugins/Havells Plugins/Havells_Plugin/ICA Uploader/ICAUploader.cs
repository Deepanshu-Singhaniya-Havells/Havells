using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Text.RegularExpressions;
using Havells_Plugin;
using System.Text;
using Microsoft.Crm.Sdk.Messages;

namespace Havells_Plugin.ICA_Uploader
{
    public class PreCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            Entity entity = (Entity)context.InputParameters["Target"];
            Guid pmsJobId = Guid.Empty;
            try
            {
                string _callSubType = entity.GetAttributeValue<string>("hil_callsubtype");
                string _prodSubCatg = entity.GetAttributeValue<string>("hil_productsubcategory");
                string _natureOfComlaint = entity.GetAttributeValue<string>("hil_natureofcomplaint");
                string _observation = entity.GetAttributeValue<string>("hil_observation");
                string _cause = entity.GetAttributeValue<string>("hil_cause");
                string _causeAction = entity.GetAttributeValue<string>("hil_causeaction");
                decimal _amount = entity.GetAttributeValue<decimal>("hil_actionprice");

                StringBuilder _errorRemarks = null;
                string _retMsg;

                bool _errorOccured = false;
                Guid _prodSubCatgGuId = Guid.Empty;
                Guid _callSubTypGuId = Guid.Empty;
                Guid _natureOfComlaintGuId = Guid.Empty;
                Guid _observationGuId = Guid.Empty;
                Guid _causeGuId = Guid.Empty;
                Guid _causeActionGuId = Guid.Empty;
                QueryExpression _queryExp;
                EntityCollection _entcoll;

                #region Validating Inputs
                _errorRemarks = new StringBuilder();
                _queryExp = new QueryExpression("product");
                _queryExp.ColumnSet = new ColumnSet("name");
                _queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                _queryExp.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 3); // Material Group
                _queryExp.Criteria.AddCondition("producttypecode", ConditionOperator.Equal, 1); // Finished Goods
                _queryExp.Criteria.AddCondition("name", ConditionOperator.Equal, _prodSubCatg);
                _entcoll = service.RetrieveMultiple(_queryExp);
                if (_entcoll.Entities.Count == 0)
                {
                    _errorRemarks.AppendLine("Product SubCategory does not exist.");
                    _errorOccured = true;
                }
                else
                {
                    _prodSubCatgGuId = _entcoll.Entities[0].Id;
                }
                _queryExp = new QueryExpression("hil_callsubtype");
                _queryExp.ColumnSet = new ColumnSet("hil_name");
                _queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                _queryExp.Criteria.AddCondition("hil_name", ConditionOperator.Equal, _callSubType);
                _entcoll = service.RetrieveMultiple(_queryExp);
                if (_entcoll.Entities.Count == 0)
                {
                    _errorRemarks.AppendLine("Call Subtype does not exist.");
                    _errorOccured = true;
                }
                else
                {
                    _callSubTypGuId = _entcoll.Entities[0].Id;
                }
                #endregion
                if (!_errorOccured)
                {
                    _natureOfComlaintGuId = GetNatureOfComplaint(_callSubTypGuId, _prodSubCatgGuId, _natureOfComlaint, service, out _retMsg);
                    if (_natureOfComlaintGuId != Guid.Empty)
                    {
                        if (_retMsg != "OK") { _errorRemarks.AppendLine(_retMsg); }
                        _observationGuId = GetObservation(_prodSubCatgGuId, _observation, service, out _retMsg);
                        if (_observationGuId != Guid.Empty)
                        {
                            if (_retMsg != "OK") { _errorRemarks.AppendLine(_retMsg); }
                            _causeGuId = GetCause(_observationGuId, _prodSubCatgGuId, _cause, service, out _retMsg);
                            if (_causeGuId != Guid.Empty)
                            {
                                if (_retMsg != "OK") { _errorRemarks.AppendLine(_retMsg); }
                                _causeActionGuId = GetService(_causeAction, _amount, service, out _retMsg);
                                if (_causeActionGuId != Guid.Empty)
                                {
                                    if (_retMsg != "OK") { _errorRemarks.AppendLine(_retMsg); }
                                    Guid _causeServiceGuId = GetCauseService(_causeGuId, _causeActionGuId, "", service, out _retMsg);
                                    if (_causeServiceGuId == Guid.Empty)
                                    {
                                        _errorOccured = true;
                                        _errorRemarks.AppendLine(_retMsg);
                                    }
                                    else {
                                        _errorRemarks.AppendLine(_retMsg);
                                    }
                                }
                                else
                                {
                                    _errorOccured = true;
                                    _errorRemarks.AppendLine(_retMsg);
                                }
                            }
                            else
                            {
                                _errorOccured = true;
                                _errorRemarks.AppendLine(_retMsg);
                            }
                        }
                        else
                        {
                            _errorOccured = true;
                            _errorRemarks.AppendLine(_retMsg);
                        }
                    }
                    else
                    {
                        _errorOccured = true;
                        _errorRemarks.AppendLine(_retMsg);
                    }
                }
                entity["hil_name"] = _errorRemarks.ToString();
                entity["hil_icastatus"] = _errorOccured;
            }
            catch (Exception ex)
            {
                entity["hil_name"] = ex.Message;
                entity["hil_icastatu"] = false;
                throw new InvalidPluginExecutionException("  ***Havells_Plugin.ICAUploader.PreCreate.Execute***  " + ex.Message);
            }
        }

        public Guid GetNatureOfComplaint(Guid _callSubTypeGuId, Guid _prodSubcategoryGuId, string _nocName, IOrganizationService _service, out string _retMsg)
        {
            QueryExpression _queryExp;
            EntityCollection _entcoll;
            Guid _retValue = Guid.Empty;
            _retMsg = "OK";
            try
            {
                _queryExp = new QueryExpression("hil_natureofcomplaint");
                _queryExp.ColumnSet = new ColumnSet("statecode");
                _queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                _queryExp.Criteria.AddCondition("hil_name", ConditionOperator.Equal, _nocName);
                _queryExp.Criteria.AddCondition("hil_callsubtype", ConditionOperator.Equal, _callSubTypeGuId);
                _queryExp.Criteria.AddCondition("hil_relatedproduct", ConditionOperator.Equal, _prodSubcategoryGuId);
                _entcoll = _service.RetrieveMultiple(_queryExp);
                if (_entcoll.Entities.Count > 0)
                {
                    if (_entcoll.Entities[0].GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                    {
                        SetStateRequest setStateRequest = new SetStateRequest()
                        {
                            EntityMoniker = new EntityReference
                            {
                                Id = _entcoll.Entities[0].Id,
                                LogicalName = "hil_natureofcomplaint",
                            },
                            State = new OptionSetValue(0),
                            Status = new OptionSetValue(1)
                        };
                        _service.Execute(setStateRequest);
                        _retMsg = "Nature Of Complaint was Deactive. It's Activate Now";
                    }
                    _retValue = _entcoll.Entities[0].Id;
                }
                else
                {
                    Entity entObj = new Entity("hil_natureofcomplaint");
                    entObj.Attributes["hil_name"] = _nocName;
                    entObj.Attributes["hil_callsubtype"] = new EntityReference("hil_callsubtype", _callSubTypeGuId);
                    entObj.Attributes["hil_relatedproduct"] = new EntityReference("product", _prodSubcategoryGuId);
                    _retValue = _service.Create(entObj);
                }
            }
            catch (Exception ex)
            {
                _retMsg = ex.Message;
            }
            return _retValue;
        }
        public Guid GetObservation(Guid _prodSubcategoryGuId, string _observationName, IOrganizationService _service, out string _retMsg)
        {
            QueryExpression _queryExp;
            EntityCollection _entcoll;
            Guid _retValue = Guid.Empty;
            _retMsg = "OK";
            try
            {
                _queryExp = new QueryExpression("hil_observation");
                _queryExp.ColumnSet = new ColumnSet("statecode");
                _queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                _queryExp.Criteria.AddCondition("hil_name", ConditionOperator.Equal, _observationName);
                _queryExp.Criteria.AddCondition("hil_relatedproduct", ConditionOperator.Equal, _prodSubcategoryGuId);
                _entcoll = _service.RetrieveMultiple(_queryExp);
                if (_entcoll.Entities.Count > 0)
                {
                    if (_entcoll.Entities[0].GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                    {
                        SetStateRequest setStateRequest = new SetStateRequest()
                        {
                            EntityMoniker = new EntityReference
                            {
                                Id = _entcoll.Entities[0].Id,
                                LogicalName = "hil_observation",
                            },
                            State = new OptionSetValue(0),
                            Status = new OptionSetValue(1)
                        };
                        _service.Execute(setStateRequest);
                        _retMsg = "Observation was Deactive. It's Activate Now";
                    }
                    _retValue = _entcoll.Entities[0].Id;
                }
                else
                {
                    Entity entObj = new Entity("hil_observation");
                    entObj.Attributes["hil_name"] = _observationName;
                    entObj.Attributes["hil_relatedproduct"] = new EntityReference("product", _prodSubcategoryGuId);
                    _retValue = _service.Create(entObj);
                }
            }
            catch (Exception ex)
            {
                _retMsg = ex.Message;
            }
            return _retValue;
        }
        public Guid GetCause(Guid _observationGuId, Guid _prodSubcategoryGuId, string _causeName, IOrganizationService _service, out string _retMsg)
        {
            QueryExpression _queryExp;
            EntityCollection _entcoll;
            Guid _retValue = Guid.Empty;
            _retMsg = "OK";
            try
            {
                _queryExp = new QueryExpression("msdyn_incidenttype");
                _queryExp.ColumnSet = new ColumnSet("statecode");
                _queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                _queryExp.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, _causeName);
                _queryExp.Criteria.AddCondition("hil_observation", ConditionOperator.Equal, _observationGuId);
                _queryExp.Criteria.AddCondition("hil_model", ConditionOperator.Equal, _prodSubcategoryGuId);
                _entcoll = _service.RetrieveMultiple(_queryExp);
                if (_entcoll.Entities.Count > 0)
                {
                    if (_entcoll.Entities[0].GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                    {
                        SetStateRequest setStateRequest = new SetStateRequest()
                        {
                            EntityMoniker = new EntityReference
                            {
                                Id = _entcoll.Entities[0].Id,
                                LogicalName = "msdyn_incidenttype",
                            },
                            State = new OptionSetValue(0),
                            Status = new OptionSetValue(1)
                        };
                        _service.Execute(setStateRequest);
                        _retMsg = "Cause was Deactive. It's Activate Now";
                    }
                    _retValue = _entcoll.Entities[0].Id;
                }
                else
                {
                    Entity entObj = new Entity("msdyn_incidenttype");
                    entObj.Attributes["msdyn_name"] = _causeName;
                    entObj.Attributes["hil_observation"] = new EntityReference("hil_observation", _observationGuId);
                    entObj.Attributes["hil_model"] = new EntityReference("product", _prodSubcategoryGuId);
                    _retValue = _service.Create(entObj);
                }
            }
            catch (Exception ex)
            {
                _retMsg = ex.Message;
            }
            return _retValue;
        }
        public Guid GetService(string _serviceName, decimal _amount, IOrganizationService _service, out string _retMsg)
        {
            QueryExpression _queryExp;
            EntityCollection _entcoll;
            Guid _retValue = Guid.Empty;
            _retMsg = "OK";
            try
            {
                string _uniqueKey = _serviceName;
                if (_amount > 0)
                {
                    _uniqueKey = _serviceName + Convert.ToInt32(_amount).ToString();
                }

                _queryExp = new QueryExpression("product");
                _queryExp.ColumnSet = new ColumnSet("statecode");
                _queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                FilterExpression _fltExp = _queryExp.Criteria.AddFilter(LogicalOperator.Or);
                _fltExp.AddCondition("productnumber", ConditionOperator.Equal, _serviceName);
                _fltExp.AddCondition("hil_productcode", ConditionOperator.Equal, _uniqueKey);
                _entcoll = _service.RetrieveMultiple(_queryExp);
                if (_entcoll.Entities.Count > 0)
                {
                    _retValue = _entcoll.Entities[0].Id;
                }
                else
                {
                    Entity entObj = new Entity("product");
                    entObj.Attributes["name"] = _serviceName;
                    entObj.Attributes["productnumber"] = _uniqueKey;
                    entObj.Attributes["hil_productcode"] = _uniqueKey;
                    entObj.Attributes["hil_uniquekey"] = _uniqueKey;
                    entObj.Attributes["productstructure"] = new OptionSetValue(1);
                    entObj.Attributes["hil_amount"] = new Money(_amount);
                    entObj.Attributes["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("68A6A9CA-6BEB-E811-A96C-000D3AF05828"));
                    entObj.Attributes["pricelevelid"] = new EntityReference("pricelevel", new Guid("BB52E023-9609-E911-A94D-000D3AF0694E"));
                    entObj.Attributes["defaultuomid"] = new EntityReference("uom", new Guid("0359D51B-D7CF-43B1-87F6-FC13A2C1DEC8"));
                    entObj.Attributes["msdyn_fieldserviceproducttype"] = new OptionSetValue(690970002);
                    entObj.Attributes["hil_hierarchylevel"] = new OptionSetValue(910590000);
                    entObj.Attributes["producttypecode"] = new OptionSetValue(4);
                    entObj.Attributes["defaultuomscheduleid"] = new EntityReference("uomschedule", new Guid("AF39A94C-F79F-4E6D-9A9E-20F2948FE185"));
                    _retValue = _service.Create(entObj);
                }
            }
            catch (Exception ex)
            {
                _retMsg = ex.Message;
            }
            return _retValue;
        }
        public Guid GetCauseService(Guid _causeGuId, Guid _serviceGuId, string _causeServiceName, IOrganizationService _service, out string _retMsg)
        {
            QueryExpression _queryExp;
            EntityCollection _entcoll;
            Guid _retValue = Guid.Empty;
            _retMsg = "OK";
            try
            {
                _queryExp = new QueryExpression("msdyn_incidenttypeservice");
                _queryExp.ColumnSet = new ColumnSet("statecode");
                _queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                _queryExp.Criteria.AddCondition("msdyn_incidenttype", ConditionOperator.Equal, _causeGuId);
                _queryExp.Criteria.AddCondition("msdyn_service", ConditionOperator.Equal, _serviceGuId);
                _entcoll = _service.RetrieveMultiple(_queryExp);
                if (_entcoll.Entities.Count > 0)
                {
                    if (_entcoll.Entities[0].GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                    {
                        SetStateRequest setStateRequest = new SetStateRequest()
                        {
                            EntityMoniker = new EntityReference
                            {
                                Id = _entcoll.Entities[0].Id,
                                LogicalName = "msdyn_incidenttypeservice",
                            },
                            State = new OptionSetValue(0),
                            Status = new OptionSetValue(1)
                        };
                        _service.Execute(setStateRequest);
                        _retMsg = "Cause Service was Deactive. It's Activate Now";
                        _retValue = _entcoll.Entities[0].Id;
                    }
                    else
                    {
                        _retMsg = "Duplicate Cause Service found.";
                    }
                }
                else
                {
                    Entity entObj = new Entity("msdyn_incidenttypeservice");
                    entObj.Attributes["msdyn_name"] = _causeServiceName;
                    entObj.Attributes["msdyn_duration"] = 30; // 30 Minutes
                    entObj.Attributes["msdyn_unit"] = new EntityReference("uom", new Guid("0359D51B-D7CF-43B1-87F6-FC13A2C1DEC8")); //Primary Unit
                    entObj.Attributes["msdyn_incidenttype"] = new EntityReference("msdyn_incidenttype", _causeGuId);
                    entObj.Attributes["msdyn_service"] = new EntityReference("product", _serviceGuId);
                    _retValue = _service.Create(entObj);
                }
            }
            catch (Exception ex)
            {
                _retMsg = ex.Message;
            }
            return _retValue;
        }
    }
}
