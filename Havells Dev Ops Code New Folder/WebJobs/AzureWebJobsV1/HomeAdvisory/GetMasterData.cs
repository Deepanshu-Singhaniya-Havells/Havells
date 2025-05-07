using System.Text;
using System.Globalization;
using System.Collections.Generic;
using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RestSharp;
using Newtonsoft.Json;
using System.Net;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace HomeAdvisory
{
    public static class OnDemandJobs
    {
        static public void CancellationReasonandConstructionType(IOrganizationService service)
        {
            try
            {
                Integration integration = common.IntegrationConfiguration(service, "GetLookupMasterDataMasterForHomeAdvisory");

                string _authInfo = integration.Auth;
                _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));
                String sUrl = integration.uri;

                var client = new RestClient(sUrl);

                client.Timeout = -1;
                var request = new RestRequest(Method.GET);
                request.AddHeader("Authorization", "Basic " + _authInfo);
                IRestResponse response = client.Execute(request);
                Console.WriteLine(response.Content);

                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<Root>(response.Content);

                foreach (LookUpRes look in obj.Results)
                {
                    if (look.LookupType == "CloseEnquiry")
                    {
                        QueryExpression query = new QueryExpression("hil_cancellationreason");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, look.Code.ToString());
                        EntityCollection entCol = service.RetrieveMultiple(query);

                        Entity _entity = new Entity("hil_cancellationreason");
                        _entity["hil_code"] = look.Code.ToString();
                        _entity["hil_sequence"] = look.Sequence.ToString();
                        _entity["hil_name"] = look.Name.ToString();
                        _entity["hil_lookuptype"] = look.LookupType.ToString();
                        _entity["hil_lookupid"] = look.LookupID.ToString();
                        _entity["hil_mastertypeid"] = look.MasterTypeID.ToString();
                        _entity["hil_mdmtimestamp"] = (look.ModifiedDate != "" && look.ModifiedDate != null) ? Convert.ToDateTime(look.ModifiedDate) : Convert.ToDateTime(look.CreatedDate);
                        if (entCol.Entities.Count == 0)
                        {
                            service.Create(_entity);
                        }
                        else if (entCol.Entities.Count > 0)
                        {
                            _entity.Id = entCol[0].Id;
                            service.Update(_entity);
                        }
                    }
                    else if (look.LookupType == "ConstructionType")
                    {
                        QueryExpression query = new QueryExpression("hil_constructiontype");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, look.Code.ToString());
                        EntityCollection entCol = service.RetrieveMultiple(query);
                        if (entCol.Entities.Count == 0)
                        {

                            Entity _entity = new Entity("hil_constructiontype");
                            _entity["hil_name"] = look.Name.ToString();
                            _entity["hil_code"] = look.Code.ToString();

                            service.Create(_entity);
                        }
                        else if (entCol.Entities.Count > 0)
                        {

                            Entity _entity = new Entity("hil_constructiontype");
                            _entity["hil_name"] = look.Name.ToString();
                            _entity["hil_code"] = look.Code.ToString();

                            service.Update(_entity);
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("D365 Internal Error " + ex.Message);
            }
        }
    }
    public static class GetMasterData
    {
        public static void getEnquiryType(IOrganizationService service, String timestamp)
        {

            Console.WriteLine("****************************************************");
            Console.WriteLine("getEnquiryType...");
            Console.WriteLine("****************************************************");
            try
            {
                QueryExpression query = new QueryExpression("hil_enquirytype");
                query.ColumnSet = new ColumnSet("hil_mdmtimestamp");
                query.TopCount = 1;
                query.AddOrder("hil_mdmtimestamp", OrderType.Descending);
                EntityCollection entiColl = service.RetrieveMultiple(query);
                if (entiColl.Entities.Count > 0)
                {
                    if (entiColl[0].Contains("hil_mdmtimestamp"))
                    {
                        DateTime _cTimeStamp = entiColl[0].GetAttributeValue<DateTime>("hil_mdmtimestamp").AddMinutes(330);
                        timestamp = _cTimeStamp.Year.ToString() + _cTimeStamp.Month.ToString().PadLeft(2, '0') + _cTimeStamp.Day.ToString().PadLeft(2, '0') + _cTimeStamp.Hour.ToString().PadLeft(2, '0') + _cTimeStamp.Minute.ToString().PadLeft(2, '0') + _cTimeStamp.Second.ToString().PadLeft(2, '0');
                    }
                }
                EnquryType enqtyp = new EnquryType();
                Integration integration = common.IntegrationConfiguration(service, "GetEnquiryTypeMaster");

                string _authInfo = integration.Auth;
                _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));
                String sUrl = integration.uri + timestamp;
                var client = new RestClient(sUrl);// URL + "GetEnquiryType?enquiryDate=" + timestamp);
                client.Timeout = -1;
                var request = new RestSharp.RestRequest(Method.GET);
                request.AddHeader("Authorization", "Basic " + _authInfo);
                IRestResponse response = client.Execute(request);
                //Console.WriteLine(response.Content);
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<EnquryType>(response.Content);
                Console.WriteLine("obj.Results.Count... " + obj.Results.Count);
                for (int i = 0; i < obj.Results.Count; i++)
                {
                    #region fetch Enquery type...
                    query = new QueryExpression("hil_enquirytype");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_enquirytypecode", ConditionOperator.Equal, obj.Results[i].EnquiryTypeCode.Trim());
                    EntityCollection entCol = service.RetrieveMultiple(query);
                    DateTime cc = obj.Results[i].ModifiedDate != "" && obj.Results[i].ModifiedDate != null ? Convert.ToDateTime(obj.Results[i].ModifiedDate) : Convert.ToDateTime(obj.Results[i].CreatedDate.ToString());
                    Entity enqury = new Entity("hil_enquirytype");
                    if (obj.Results[i].EnquiryTypeCode != null)
                    {
                        enqury["hil_enquirytypecode"] = Int32.Parse(obj.Results[i].EnquiryTypeCode);
                    }
                    if (obj.Results[i].EnquiryTypeDesc != null)
                    {
                        enqury["hil_enquirytypename"] = obj.Results[i].EnquiryTypeDesc.ToString().Trim();
                    }
                    Guid enqid = Guid.Empty;
                    enqury["hil_mdmtimestamp"] = obj.Results[i].ModifiedDate != "" && obj.Results[i].ModifiedDate != null ? Convert.ToDateTime(obj.Results[i].ModifiedDate) : Convert.ToDateTime(obj.Results[i].CreatedDate.ToString());
                    if (entCol.Entities.Count == 0)
                    {
                        enqid = service.Create(enqury);
                    }
                    else
                    {
                        enqid = entCol[0].Id;
                        enqury.Id = enqid;
                        service.Update(enqury);

                        Console.WriteLine("Enquery Type Updated Count = " + i);
                    }
                    if (obj.Results[i].IsActive)
                    {
                        SetStateRequest setStateRequest = new SetStateRequest()
                        {
                            EntityMoniker = new EntityReference
                            {
                                Id = enqid,
                                LogicalName = enqury.LogicalName,
                            },
                            State = new OptionSetValue(0), //active
                            Status = new OptionSetValue(1) //active
                        };
                        service.Execute(setStateRequest);
                    }
                    else
                    {
                        SetStateRequest setStateRequest = new SetStateRequest()
                        {
                            EntityMoniker = new EntityReference
                            {
                                Id = enqid,
                                LogicalName = enqury.LogicalName,
                            },
                            State = new OptionSetValue(1), //deactive
                            Status = new OptionSetValue(2) //deactive
                        };
                        service.Execute(setStateRequest);
                    }
                    Console.WriteLine("Enquery Type Created Count = " + i);
                    #endregion
                }
                Console.WriteLine("Enquery Type Done...");
            }
            catch (Exception ex)
            {
                Console.WriteLine("D365 Internal Error " + ex.Message);
            }
            Console.WriteLine("getEnquiryType Completed Sucessfully....");
        }
        public static void getProductType(IOrganizationService service, String timestamp)
        {
            Console.WriteLine("****************************************************");
            Console.WriteLine("getEnquiryType...");
            Console.WriteLine("****************************************************");
            try
            {
                string _globalOptionSetName = "hil_typeofproduct";
                //Guid enqid = Guid.Empty;
                bool _publish = false;
                QueryExpression query = new QueryExpression("hil_typeofproduct");
                query.ColumnSet = new ColumnSet("hil_mdmtimestamp");
                query.TopCount = 1;
                query.AddOrder("hil_mdmtimestamp", OrderType.Descending);
                EntityCollection entiColl = service.RetrieveMultiple(query);
                if (entiColl.Entities.Count > 0)
                {
                    if (entiColl[0].Contains("hil_mdmtimestamp"))
                    {
                        DateTime _cTimeStamp = entiColl[0].GetAttributeValue<DateTime>("hil_mdmtimestamp").AddMinutes(330);
                        timestamp = _cTimeStamp.Year.ToString() + _cTimeStamp.Month.ToString().PadLeft(2, '0') + _cTimeStamp.Day.ToString().PadLeft(2, '0') + _cTimeStamp.Hour.ToString().PadLeft(2, '0') + _cTimeStamp.Minute.ToString().PadLeft(2, '0') + _cTimeStamp.Second.ToString().PadLeft(2, '0');
                    }
                }
                ProductType productTypedetails = new ProductType();

                Integration integration = common.IntegrationConfiguration(service, "GetAdvisoryProductTypeMaster");

                string _authInfo = integration.Auth;
                _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));
                String sUrl = integration.uri + timestamp;

                var client = new RestClient(sUrl);

                client.Timeout = -1;
                var request = new RestRequest(Method.GET);
                request.AddHeader("Authorization", "Basic " + _authInfo);
                IRestResponse response = client.Execute(request);

                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<ProductType>(response.Content);
                Console.WriteLine("obj.Results.Count  " + obj.Results.Count);
                for (int i = 0; i < obj.Results.Count; i++)
                {
                    query = new QueryExpression("hil_typeofproduct");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, obj.Results[i].AdvisoryProductTypeCode.Trim());
                    EntityCollection entCol = service.RetrieveMultiple(query);
                    Entity ProductType = new Entity("hil_typeofproduct");
                    // enqury.Id = EnqId;
                    Guid enqid = Guid.Empty;
                    if (obj.Results[i].AdvisoryProductTypeCode != null)
                    {
                        ProductType["hil_code"] = obj.Results[i].AdvisoryProductTypeCode.ToString().Trim();
                    }
                    if (obj.Results[i].AdvisoryProductType != null)
                    {
                        ProductType["hil_name"] = obj.Results[i].AdvisoryProductType.ToString().Trim();
                    }
                    ProductType["hil_mdmtimestamp"] = obj.Results[i].ModifiedDate != "" && obj.Results[i].ModifiedDate != null ? Convert.ToDateTime(obj.Results[i].ModifiedDate) : Convert.ToDateTime(obj.Results[i].CreatedDate.ToString());

                    if (entCol.Entities.Count == 0)
                    {
                        enqid = service.Create(ProductType);

                        Console.WriteLine("Product Type Created Count = " + i);
                    }
                    else if (entCol.Entities.Count > 0)
                    {
                        enqid = entCol[0].Id;
                        ProductType.Id = enqid;
                        service.Update(ProductType);

                        Console.WriteLine("Product Type Updated Count = " + i);
                    }
                    if (obj.Results[i].IsActive)
                    {
                        SetStateRequest setStateRequest = new SetStateRequest()
                        {
                            EntityMoniker = new EntityReference
                            {
                                Id = enqid,
                                LogicalName = ProductType.LogicalName,
                            },
                            State = new OptionSetValue(0), //active
                            Status = new OptionSetValue(1) //active
                        };
                        service.Execute(setStateRequest);
                    }
                    else
                    {
                        SetStateRequest setStateRequest = new SetStateRequest()
                        {
                            EntityMoniker = new EntityReference
                            {
                                Id = enqid,
                                LogicalName = ProductType.LogicalName,
                            },
                            State = new OptionSetValue(1), //deactive
                            Status = new OptionSetValue(2) //deactive
                        };
                        service.Execute(setStateRequest);
                    }
                    string _index = service.Retrieve("hil_typeofproduct", enqid, new ColumnSet("hil_index")).GetAttributeValue<string>("hil_index");
                    #region Create Options in GlobalOptionSet
                    RetrieveOptionSetRequest retrieveOptionSetRequest =
                    new RetrieveOptionSetRequest
                    {
                        Name = _globalOptionSetName
                    };
                    // Execute the request.
                    RetrieveOptionSetResponse retrieveOptionSetResponse = (RetrieveOptionSetResponse)service.Execute(retrieveOptionSetRequest);
                    // Access the retrieved OptionSetMetadata.
                    OptionSetMetadata retrievedOptionSetMetadata =
                    (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;
                    // Get the current options list for the retrieved attribute.
                    OptionMetadata[] optionList =
                    retrievedOptionSetMetadata.Options.ToArray();
                    if (Array.Find(optionList, o => o.Value == Convert.ToInt32(_index)) == null)
                    {
                        InsertOptionValueRequest insertOptionValueRequest =
                        new InsertOptionValueRequest
                        {
                            OptionSetName = _globalOptionSetName,
                            Label = new Label(obj.Results[i].AdvisoryProductType.ToString().Trim(), 1033),
                            Value = Convert.ToInt32(_index)
                        };
                        int _insertedOptionValue = ((InsertOptionValueResponse)service.Execute(
                         insertOptionValueRequest)).NewOptionValue;
                        _publish = true;
                    }
                    else
                    {
                        UpdateOptionValueRequest updateOptionValueRequest =
                        new UpdateOptionValueRequest
                        {
                            OptionSetName = _globalOptionSetName,
                            Value = Convert.ToInt32(_index),
                            Label = new Label(obj.Results[i].AdvisoryProductType.ToString().Trim(), 1033)
                        };
                        service.Execute(updateOptionValueRequest);
                        _publish = true;
                    }
                    if (_publish)
                    {
                        //Publish the OptionSet
                        PublishXmlRequest pxReq2 = new PublishXmlRequest { ParameterXml = String.Format("<importexportxml><optionsets><optionset>{0}</optionset></optionsets></importexportxml>", _globalOptionSetName) };
                        service.Execute(pxReq2);
                    }
                    #endregion
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("D365 Internal Error " + ex.Message);
            }
            Console.WriteLine("getProductType Completed Sucessfully....");
        }
        public static void GetProductMapping(IOrganizationService service, String timestamp)
        {
            Console.WriteLine("****************************************************");
            Console.WriteLine("getEnquiryType...");
            Console.WriteLine("****************************************************");
            try
            {
                QueryExpression query = new QueryExpression("hil_typeofproduct");
                query.ColumnSet = new ColumnSet("hil_mdmtimestamp");
                query.TopCount = 1;
                query.AddOrder("hil_mdmtimestamp", OrderType.Descending);
                EntityCollection entiColl = service.RetrieveMultiple(query);
                if (entiColl.Entities.Count > 0)
                {
                    if (entiColl[0].Contains("hil_mdmtimestamp"))
                    {
                        DateTime _cTimeStamp = entiColl[0].GetAttributeValue<DateTime>("hil_mdmtimestamp").AddMinutes(330);
                        timestamp = _cTimeStamp.Year.ToString() + _cTimeStamp.Month.ToString().PadLeft(2, '0') + _cTimeStamp.Day.ToString().PadLeft(2, '0') + _cTimeStamp.Hour.ToString().PadLeft(2, '0') + _cTimeStamp.Minute.ToString().PadLeft(2, '0') + _cTimeStamp.Second.ToString().PadLeft(2, '0');
                    }
                }
                GetEnquiryProductMapping productTypedetails = new GetEnquiryProductMapping();
                Integration integration = common.IntegrationConfiguration(service, "GetEnquiryProductMappingMaster");
                string _authInfo = integration.Auth;
                _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));
                String sUrl = integration.uri + timestamp;
                var client = new RestClient(sUrl);
                client.Timeout = -1;
                var request = new RestRequest(Method.GET);
                request.AddHeader("Authorization", "Basic " + _authInfo);
                IRestResponse response = client.Execute(request);

                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<GetEnquiryProductMapping>(response.Content);
                Console.WriteLine("obj.Results.Count  " + obj.Results.Count);
                for (int i = 0; i < obj.Results.Count; i++)
                {
                    query = new QueryExpression("hil_enquirytype");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_enquirytypecode", ConditionOperator.Equal, obj.Results[i].EnquiryTypeCode.Trim());
                    EntityCollection entCol = service.RetrieveMultiple(query);
                    if (entCol.Entities.Count > 0)
                    {
                        query = new QueryExpression("hil_typeofproduct");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, obj.Results[i].AdvisoryProductTypeCode.Trim());
                        EntityCollection prodCol = service.RetrieveMultiple(query);
                        if (prodCol.Entities.Count > 0)
                        {
                            Entity ProductType = new Entity("hil_typeofproduct");
                            ProductType.Id = prodCol[0].Id;
                            ProductType["hil_typeofenquiry"] = new EntityReference("hil_enquirytype", entCol.Entities[0].Id);
                            service.Update(ProductType);
                        }
                    }
                    Console.WriteLine("Mapping Updated Count = " + i);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("D365 Internal Error " + ex.Message);
            }
            Console.WriteLine("getProductType Completed Sucessfully....");
        }
        public static void GetCustomersType(IOrganizationService service, String timestamp)
        {
            Console.WriteLine("****************************************************");
            Console.WriteLine("getEnquiryType...");
            Console.WriteLine("****************************************************");
            try
            {
                QueryExpression query = new QueryExpression("hil_typeofcustomer");
                query.ColumnSet = new ColumnSet("hil_mdmtimestamp");
                query.TopCount = 1;
                query.AddOrder("hil_mdmtimestamp", OrderType.Descending);
                EntityCollection entiColl = service.RetrieveMultiple(query);
                if (entiColl.Entities.Count > 0)
                {
                    if (entiColl[0].Contains("hil_mdmtimestamp"))
                    {
                        DateTime _cTimeStamp = entiColl[0].GetAttributeValue<DateTime>("hil_mdmtimestamp").AddMinutes(330);
                        timestamp = _cTimeStamp.Year.ToString() + _cTimeStamp.Month.ToString().PadLeft(2, '0') + _cTimeStamp.Day.ToString().PadLeft(2, '0') + _cTimeStamp.Hour.ToString().PadLeft(2, '0') + _cTimeStamp.Minute.ToString().PadLeft(2, '0') + _cTimeStamp.Second.ToString().PadLeft(2, '0');
                    }
                }
                GetPropertyType productTypedetails = new GetPropertyType();


                Integration integration = common.IntegrationConfiguration(service, "GetCustomersTypeMaster");

                string _authInfo = integration.Auth;
                _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));
                String sUrl = integration.uri + timestamp;

                var client = new RestClient(sUrl);

                // var client = new RestClient(URL + "GetCustomersType?enquiryDate=" + timestamp);
                client.Timeout = -1;
                var request = new RestRequest(Method.GET);
                request.AddHeader("Authorization", "Basic " + _authInfo);
                IRestResponse response = client.Execute(request);
                // Console.WriteLine(response.Content);
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<GetPropertyType>(response.Content);
                Console.WriteLine("obj.Results.Count  " + obj.Results.Count);
                for (int i = 0; i < obj.Results.Count; i++)
                {
                    query = new QueryExpression("hil_typeofcustomer");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, obj.Results[i].Code.Trim());
                    EntityCollection entCol = service.RetrieveMultiple(query);
                    if (entCol.Entities.Count == 0)
                    {
                        Entity CustomerType = new Entity("hil_typeofcustomer");
                        // enqury.Id = EnqId;
                        if (obj.Results[i].Code != null)
                        {
                            CustomerType["hil_code"] = obj.Results[i].Code.ToString().Trim();
                        }
                        if (obj.Results[i].Name != null)
                        {
                            CustomerType["hil_name"] = obj.Results[i].Name.ToString().Trim();
                        }
                        CustomerType["hil_mdmtimestamp"] = obj.Results[i].ModifiedDate != "" && obj.Results[i].ModifiedDate != null ? Convert.ToDateTime(obj.Results[i].ModifiedDate) : Convert.ToDateTime(obj.Results[i].CreatedDate.ToString());
                        Guid enqid = service.Create(CustomerType);
                        if (obj.Results[i].IsActive)
                        {
                            SetStateRequest setStateRequest = new SetStateRequest()
                            {
                                EntityMoniker = new EntityReference
                                {
                                    Id = enqid,
                                    LogicalName = CustomerType.LogicalName,
                                },
                                State = new OptionSetValue(0), //active
                                Status = new OptionSetValue(1) //active
                            };
                            service.Execute(setStateRequest);
                        }
                        else
                        {
                            SetStateRequest setStateRequest = new SetStateRequest()
                            {
                                EntityMoniker = new EntityReference
                                {
                                    Id = enqid,
                                    LogicalName = CustomerType.LogicalName,
                                },
                                State = new OptionSetValue(1), //deactive
                                Status = new OptionSetValue(2) //deactive
                            };
                            service.Execute(setStateRequest);
                        }
                    }
                    else if (entCol.Entities.Count > 0)
                    {
                        Entity CustomerType = new Entity("hil_typeofcustomer");
                        if (obj.Results[i].Code != null)
                        {
                            CustomerType["hil_code"] = obj.Results[i].Code.ToString().Trim();
                        }
                        if (obj.Results[i].Name != null)
                        {
                            CustomerType["hil_name"] = obj.Results[i].Name.ToString().Trim();
                        }
                        CustomerType.Id = entCol.Entities[0].Id;
                        CustomerType["hil_mdmtimestamp"] = obj.Results[i].ModifiedDate != "" && obj.Results[i].ModifiedDate != null ? Convert.ToDateTime(obj.Results[i].ModifiedDate) : Convert.ToDateTime(obj.Results[i].CreatedDate.ToString());
                        service.Update(CustomerType);
                        if (obj.Results[i].IsActive)
                        {
                            SetStateRequest setStateRequest = new SetStateRequest()
                            {
                                EntityMoniker = new EntityReference
                                {
                                    Id = CustomerType.Id,
                                    LogicalName = CustomerType.LogicalName,
                                },
                                State = new OptionSetValue(0), //active
                                Status = new OptionSetValue(1) //active
                            };
                            service.Execute(setStateRequest);
                        }
                        else
                        {
                            SetStateRequest setStateRequest = new SetStateRequest()
                            {
                                EntityMoniker = new EntityReference
                                {
                                    Id = CustomerType.Id,
                                    LogicalName = CustomerType.LogicalName,
                                },
                                State = new OptionSetValue(1), //deactive
                                Status = new OptionSetValue(2) //deactive
                            };
                            service.Execute(setStateRequest);
                        }
                    }
                    Console.WriteLine("Customer Type Created/Updated Count = " + i);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("D365 Internal Error " + ex.Message);
            }
            Console.WriteLine("getProductType Completed Sucessfully....");
        }
        public static void getPropertytypeData(IOrganizationService service, String timestamp)
        {
            Console.WriteLine("****************************************************");
            Console.WriteLine("Advisory Master...");
            Console.WriteLine("****************************************************");
            try
            {
                QueryExpression query = new QueryExpression("hil_propertytype");
                query.ColumnSet = new ColumnSet("hil_mdmtimestamp");
                query.TopCount = 1;
                query.AddOrder("hil_mdmtimestamp", OrderType.Descending);
                EntityCollection entiColl = service.RetrieveMultiple(query);
                if (entiColl.Entities.Count > 0)
                {
                    if (entiColl[0].Contains("hil_mdmtimestamp"))
                    {
                        DateTime _cTimeStamp = entiColl[0].GetAttributeValue<DateTime>("hil_mdmtimestamp").AddMinutes(330);
                        timestamp = _cTimeStamp.Year.ToString() + _cTimeStamp.Month.ToString().PadLeft(2, '0') + _cTimeStamp.Day.ToString().PadLeft(2, '0') + _cTimeStamp.Hour.ToString().PadLeft(2, '0') + _cTimeStamp.Minute.ToString().PadLeft(2, '0') + _cTimeStamp.Second.ToString().PadLeft(2, '0');
                    }
                }
                GetPropertyType enqtyp = new GetPropertyType();


                Integration integration = common.IntegrationConfiguration(service, "GetPropertyTypeMaster");

                string _authInfo = integration.Auth;
                _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));
                String sUrl = integration.uri + timestamp;

                var client = new RestClient(sUrl);


                //var client = new RestClient(URL + "GetPropertyType?enquiryDate=" + timestamp);
                client.Timeout = -1;
                var request = new RestRequest(Method.GET);

                request.AddHeader("Authorization", "Basic " + _authInfo);
                IRestResponse response = client.Execute(request);
                Console.WriteLine(response.Content);
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<GetPropertyType>(response.Content);
                Console.WriteLine("response.Count " + obj.Results.Count);
                for (int i = 0; i < obj.Results.Count; i++)
                {
                    query = new QueryExpression("hil_propertytype");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, obj.Results[i].Code.Trim());
                    EntityCollection entCol = service.RetrieveMultiple(query);
                    if (entCol.Entities.Count == 0)
                    {
                        Entity propertType = new Entity("hil_propertytype");
                        if (obj.Results[i].Code != null)
                        {
                            propertType["hil_code"] = obj.Results[i].Code.ToString().Trim();
                        }
                        if (obj.Results[i].Name != null)
                        {
                            propertType["hil_name"] = obj.Results[i].Name.ToString().Trim();
                        }
                        propertType["hil_mdmtimestamp"] = obj.Results[i].ModifiedDate != "" && obj.Results[i].ModifiedDate != null ? Convert.ToDateTime(obj.Results[i].ModifiedDate) : Convert.ToDateTime(obj.Results[i].CreatedDate.ToString());
                        Guid enqid = service.Create(propertType);
                        if (obj.Results[i].IsActive)
                        {
                            SetStateRequest setStateRequest = new SetStateRequest()
                            {
                                EntityMoniker = new EntityReference
                                {
                                    Id = enqid,
                                    LogicalName = propertType.LogicalName,
                                },
                                State = new OptionSetValue(0), //active
                                Status = new OptionSetValue(1) //active
                            };
                            service.Execute(setStateRequest);
                        }
                        else
                        {
                            SetStateRequest setStateRequest = new SetStateRequest()
                            {
                                EntityMoniker = new EntityReference
                                {
                                    Id = enqid,
                                    LogicalName = propertType.LogicalName,
                                },
                                State = new OptionSetValue(1), //deactive
                                Status = new OptionSetValue(2) //deactive
                            };
                            service.Execute(setStateRequest);
                        }

                    }
                    else if (entCol.Entities.Count > 0)
                    {

                        Entity propertType = new Entity("hil_propertytype");
                        if (obj.Results[i].Code != null)
                        {
                            propertType["hil_code"] = obj.Results[i].Code.ToString().Trim();
                        }
                        if (obj.Results[i].Name != null)
                        {
                            propertType["hil_name"] = obj.Results[i].Name.ToString().Trim();
                        }
                        propertType.Id = entCol.Entities[0].Id;
                        propertType["hil_mdmtimestamp"] = obj.Results[i].ModifiedDate != "" && obj.Results[i].ModifiedDate != null ? Convert.ToDateTime(obj.Results[i].ModifiedDate) : Convert.ToDateTime(obj.Results[i].CreatedDate.ToString());
                        service.Update(propertType);
                        if (obj.Results[i].IsActive)
                        {
                            SetStateRequest setStateRequest = new SetStateRequest()
                            {
                                EntityMoniker = new EntityReference
                                {
                                    Id = propertType.Id,
                                    LogicalName = propertType.LogicalName,
                                },
                                State = new OptionSetValue(0), //active
                                Status = new OptionSetValue(1) //active
                            };
                            service.Execute(setStateRequest);
                        }
                        else
                        {
                            SetStateRequest setStateRequest = new SetStateRequest()
                            {
                                EntityMoniker = new EntityReference
                                {
                                    Id = propertType.Id,
                                    LogicalName = propertType.LogicalName,
                                },
                                State = new OptionSetValue(1), //deactive
                                Status = new OptionSetValue(2) //deactive
                            };
                            service.Execute(setStateRequest);
                        }

                    }
                    Console.WriteLine("Property Type Created/Updated Count = " + i);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("D365 Internal Error " + ex.Message);
            }
            Console.WriteLine("Advisor Master Completed Sucessfully....");
        }
        public static void AdvisoryMaster(IOrganizationService _service, String timestamp)
        {
            Console.WriteLine("****************************************************");
            Console.WriteLine("Advisory Master...");
            Console.WriteLine("****************************************************");
            try
            {

                QueryExpression query = new QueryExpression("hil_advisormaster");
                query.ColumnSet = new ColumnSet("hil_mdmtimestamp");
                query.TopCount = 1;
                query.AddOrder("hil_mdmtimestamp", OrderType.Descending);
                EntityCollection entiColl = _service.RetrieveMultiple(query);
                Console.WriteLine("1");
                if (entiColl.Entities.Count > 0)
                {
                    Console.WriteLine("2");
                    if (entiColl[0].Contains("hil_mdmtimestamp"))
                    {
                        Console.WriteLine("2.1");
                        DateTime _cTimeStamp = entiColl[0].GetAttributeValue<DateTime>("hil_mdmtimestamp").AddMinutes(330);
                        timestamp = _cTimeStamp.Year.ToString() + _cTimeStamp.Month.ToString().PadLeft(2, '0') + _cTimeStamp.Day.ToString().PadLeft(2, '0') + _cTimeStamp.Hour.ToString().PadLeft(2, '0') + _cTimeStamp.Minute.ToString().PadLeft(2, '0') + _cTimeStamp.Second.ToString().PadLeft(2, '0');
                        Console.WriteLine("2.2" + timestamp.ToString());
                    }
                }
                Console.WriteLine("3");
                TextInfo tInfo = new CultureInfo("en-US", false).TextInfo;
                Console.WriteLine("3.1");
                WorkflowData advdivsetup = new WorkflowData();
                Integration integration = common.IntegrationConfiguration(_service, "GetWorkFlowDataMaster");
                Console.WriteLine("3.2");
                string _authInfo = integration.Auth;
                _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));
                String sUrl = integration.uri + timestamp;
                var client = new RestClient(sUrl);
                client.Timeout = -1;
                var request = new RestRequest(Method.GET);
                request.AddHeader("Authorization", "Basic " + _authInfo);
                IRestResponse response = client.Execute(request);
                Console.WriteLine("3.3");
                Console.WriteLine("Response \n" + response.Content);
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<WorkflowData>(response.Content);
                Console.WriteLine("response.Count " + obj.Results.Count);
                for (int i = 0; i < obj.Results.Count; i++)
                {
                    EntityReference hil_enquirytype = null;
                    EntityReference hil_producttypecode = null;
                    EntityReference hil_salesoffice = null;
                    #region fetch Enquery type...
                    query = new QueryExpression("hil_enquirytype");
                    query.ColumnSet.AddColumns("hil_enquirytypename", "hil_enquirytypecode");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    query.Criteria.AddCondition("hil_enquirytypecode", ConditionOperator.Equal, obj.Results[i].EnquiryTypeCode.Trim());
                    EntityCollection entCol = _service.RetrieveMultiple(query);
                    if (entCol.Entities.Count > 0)
                    {
                        hil_enquirytype = entCol[0].ToEntityReference();
                    }
                    else
                    {
                        Console.WriteLine(i + " hil_enquirytype " + obj.Results[i].EnquiryTypeCode.Trim() + " not found");
                    }
                    #endregion
                    #region fetch Product type...
                    query = new QueryExpression("hil_typeofproduct");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, obj.Results[i].AdvisoryProductTypeCode.Trim());
                    entCol = _service.RetrieveMultiple(query);
                    if (entCol.Entities.Count > 0)
                    {
                        hil_producttypecode = entCol[0].ToEntityReference();
                        //AdvDivisionsetup["hil_producttypecode"] = new EntityReference("hil_typeofproduct", entCol[0].Id);
                    }
                    else
                    {
                        Console.WriteLine(i + " hil_typeofproduct " + obj.Results[i].AdvisoryProductTypeCode.Trim() + "  not found");
                    }
                    #endregion
                    #region fetch Sales Office...
                    query = new QueryExpression("hil_salesoffice");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    query.Criteria.AddCondition("hil_salesofficecode", ConditionOperator.Equal, obj.Results[i].SaleOfficeCode.Trim());
                    entCol = _service.RetrieveMultiple(query);
                    if (entCol.Entities.Count > 0)
                    {
                        hil_salesoffice = entCol[0].ToEntityReference();
                        //AdvDivisionsetup["hil_salesoffice"] = new EntityReference("hil_salesoffice", entCol[0].Id);
                    }
                    else
                    {
                        Console.WriteLine(i + " hil_salesofficecode  " + obj.Results[i].SaleOfficeCode.Trim() + " not found");
                    }
                    #endregion
                    if (hil_enquirytype == null || hil_producttypecode == null || hil_salesoffice == null)
                    {
                        //  Console.WriteLine(i + "hil_enquirytype, hil_producttypecode, hil_salesoffice is not found...");
                    }
                    else
                    {
                        query = new QueryExpression("hil_advisormaster");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, obj.Results[i].UserCode.Trim());
                        query.Criteria.AddCondition("hil_divisioncode", ConditionOperator.Equal, hil_enquirytype.Id);
                        query.Criteria.AddCondition("hil_producttypecode", ConditionOperator.Equal, hil_producttypecode.Id);
                        query.Criteria.AddCondition("hil_salesoffice", ConditionOperator.Equal, hil_salesoffice.Id);

                        entCol = _service.RetrieveMultiple(query);
                        Console.WriteLine("hil_advisormaster Count " + entCol.Entities.Count);
                        if (entCol.Entities.Count == 0)
                        {
                            entCol = new EntityCollection();
                            Entity AdvDivisionsetup = new Entity("hil_advisormaster");
                            AdvDivisionsetup["hil_name"] = tInfo.ToTitleCase(obj.Results[i].UserName.ToLower().Trim());
                            AdvDivisionsetup["hil_code"] = obj.Results[i].UserCode.Trim();
                            AdvDivisionsetup["hil_email"] = obj.Results[i].EmailId.ToLower().Trim();
                            AdvDivisionsetup["hil_advisormobilenumber"] = obj.Results[i].MobileNumber.Trim();
                            AdvDivisionsetup["hil_mdmtimestamp"] = (obj.Results[i].ModifiedDate == null) ? Convert.ToDateTime(obj.Results[i].CreatedDate) : Convert.ToDateTime(obj.Results[i].ModifiedDate);
                            AdvDivisionsetup["hil_divisioncode"] = hil_enquirytype;
                            AdvDivisionsetup["hil_producttypecode"] = hil_producttypecode;
                            AdvDivisionsetup["hil_salesoffice"] = hil_salesoffice;
                            Guid AdvDivisionsetupid = _service.Create(AdvDivisionsetup);
                            Console.WriteLine("Created " + i);
                        }
                        else
                        {
                            Entity AdvDivisionsetup = new Entity("hil_advisormaster");
                            AdvDivisionsetup.Id = entCol[0].Id;
                            AdvDivisionsetup["hil_name"] = tInfo.ToTitleCase(obj.Results[i].UserName.ToLower().Trim());
                            AdvDivisionsetup["hil_code"] = obj.Results[i].UserCode.Trim();
                            AdvDivisionsetup["hil_email"] = obj.Results[i].EmailId.ToLower().Trim();
                            AdvDivisionsetup["hil_advisormobilenumber"] = obj.Results[i].MobileNumber.Trim();
                            AdvDivisionsetup["hil_mdmtimestamp"] = (obj.Results[i].ModifiedDate == null) ? Convert.ToDateTime(obj.Results[i].CreatedDate) : Convert.ToDateTime(obj.Results[i].ModifiedDate);
                            AdvDivisionsetup["hil_divisioncode"] = hil_enquirytype;
                            AdvDivisionsetup["hil_producttypecode"] = hil_producttypecode;
                            AdvDivisionsetup["hil_salesoffice"] = hil_salesoffice;
                            _service.Update(AdvDivisionsetup);
                            Console.WriteLine("Updated " + i);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("D365 Internal Error " + ex.Message);
            }
            Console.WriteLine("Advisor Master Completed Sucessfully....");
        }
    }
}
