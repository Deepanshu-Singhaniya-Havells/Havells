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
    public class WebJobs
    {
        #region Create EnquryType
        public static void getEnquiryType(IOrganizationService service, string URL, String timestamp)
        {
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
                        //timestamp = _cTimeStamp.Year.ToString() + _cTimeStamp.Month.ToString().PadLeft(2, '0') + _cTimeStamp.Day.ToString().PadLeft(2, '0') + _cTimeStamp.Hour.ToString().PadLeft(2, '0') + _cTimeStamp.Minute.ToString().PadLeft(2, '0') + _cTimeStamp.Second.ToString().PadLeft(2, '0');
                    }
                }

                EnquryType enqtyp = new EnquryType();

                var client = new RestClient(URL + "GetEnquiryType?enquiryDate=" + timestamp);
                client.Timeout = -1;
                var request = new RestSharp.RestRequest(Method.GET);
                request.AddHeader("Authorization", "Basic RDM2NV9IQVZFTExTOkRFVkQzNjVAMTIzNA==");
                IRestResponse response = client.Execute(request);
                //Console.WriteLine(response.Content);
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<EnquryType>(response.Content);

                //Console.WriteLine();
                for (int i = 0; i < obj.Results.Count; i++)
                {
                    #region fetch Enquery type...
                    query = new QueryExpression("hil_enquirytype");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_enquirytypecode", ConditionOperator.Equal, obj.Results[i].EnquiryTypeCode.Trim());
                    EntityCollection entCol = service.RetrieveMultiple(query);
                    DateTime cc = obj.Results[i].ModifiedDate != "" && obj.Results[i].ModifiedDate != null ? Convert.ToDateTime(obj.Results[i].ModifiedDate) : Convert.ToDateTime(obj.Results[i].CreatedDate.ToString());
                    if (entCol.Entities.Count == 0)
                    {
                        Entity enqury = new Entity("hil_enquirytype");
                        if (obj.Results[i].EnquiryTypeCode != null)
                        {
                            enqury["hil_enquirytypecode"] = Int32.Parse(obj.Results[i].EnquiryTypeCode);
                        }
                        if (obj.Results[i].EnquiryTypeDesc != null)
                        {
                            enqury["hil_enquirytypename"] = obj.Results[i].EnquiryTypeDesc.ToString().Trim();
                        }

                        enqury["hil_mdmtimestamp"] = obj.Results[i].ModifiedDate != "" && obj.Results[i].ModifiedDate != null ? Convert.ToDateTime(obj.Results[i].ModifiedDate) : Convert.ToDateTime(obj.Results[i].CreatedDate.ToString());
                        Guid enqid = service.Create(enqury);

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
                    }
                    else
                    {
                        Entity enqury = new Entity("hil_enquirytype");
                        if (obj.Results[i].EnquiryTypeCode != null)
                        {
                            enqury["hil_enquirytypecode"] = Int32.Parse(obj.Results[i].EnquiryTypeCode);
                        }
                        if (obj.Results[i].EnquiryTypeDesc != null)
                        {
                            enqury["hil_enquirytypename"] = obj.Results[i].EnquiryTypeDesc.ToString().Trim();
                        }
                        enqury["hil_mdmtimestamp"] = obj.Results[i].ModifiedDate != "" && obj.Results[i].ModifiedDate != null ? Convert.ToDateTime(obj.Results[i].ModifiedDate) : Convert.ToDateTime(obj.Results[i].CreatedDate.ToString());
                        enqury.Id = entCol[0].Id;
                        service.Update(enqury);
                        if (obj.Results[i].IsActive)
                        {
                            SetStateRequest setStateRequest = new SetStateRequest()
                            {
                                EntityMoniker = new EntityReference
                                {
                                    Id = enqury.Id,
                                    LogicalName = "hil_enquirytype",
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
                                    Id = enqury.Id,
                                    LogicalName = "hil_natureofcomplaint",
                                },
                                State = new OptionSetValue(1), //deactive
                                Status = new OptionSetValue(2) //deactive
                            };
                            service.Execute(setStateRequest);
                        }
                        Console.WriteLine("Enquery Type Updated Count = " + i);
                    }

                    #endregion
                }
                Console.WriteLine("Enquery Type Done...");
            }
            catch (Exception ex)
            {
                Console.WriteLine("D365 Internal Error " + ex.Message);
            }

        }
        #endregion
        #region Create/Update productType
        public static void getProductType(IOrganizationService service, string URL, String timestamp)
        {
            try
            {
                string _globalOptionSetName = "hil_typeofproduct";
                Guid enqid=Guid.Empty;
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
                var client = new RestClient(URL + "GetAdvisoryProductType?enquiryDate=" + timestamp);
                client.Timeout = -1;
                var request = new RestRequest(Method.GET);
                request.AddHeader("Authorization", "Basic RDM2NV9IQVZFTExTOkRFVkQzNjVAMTIzNA==");
                IRestResponse response = client.Execute(request);
                Console.WriteLine(response.Content);
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<ProductType>(response.Content);
                for (int i = 0; i < obj.Results.Count; i++)
                {
                    query = new QueryExpression("hil_typeofproduct");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, obj.Results[i].AdvisoryProductTypeCode.Trim());
                    EntityCollection entCol = service.RetrieveMultiple(query);
                    if (entCol.Entities.Count == 0)
                    {
                        Entity ProductType = new Entity("hil_typeofproduct");
                        // enqury.Id = EnqId;
                        if (obj.Results[i].AdvisoryProductTypeCode != null)
                        {
                            ProductType["hil_code"] = obj.Results[i].AdvisoryProductTypeCode.ToString().Trim();
                        }
                        if (obj.Results[i].AdvisoryProductType != null)
                        {
                            ProductType["hil_name"] = obj.Results[i].AdvisoryProductType.ToString().Trim();
                        }
                        ProductType["hil_mdmtimestamp"] = obj.Results[i].ModifiedDate != "" && obj.Results[i].ModifiedDate != null ? Convert.ToDateTime(obj.Results[i].ModifiedDate) : Convert.ToDateTime(obj.Results[i].CreatedDate.ToString());
                        enqid = service.Create(ProductType);
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
                        Console.WriteLine("Product Type Created Count = " + i);
                    }
                    else if (entCol.Entities.Count > 0)
                    {
                        Entity ProductType = new Entity("hil_typeofproduct");
                        if (obj.Results[i].AdvisoryProductTypeCode != null)
                        {
                            ProductType["hil_code"] = obj.Results[i].AdvisoryProductTypeCode.ToString().Trim();
                        }
                        if (obj.Results[i].AdvisoryProductType != null)
                        {
                            ProductType["hil_name"] = obj.Results[i].AdvisoryProductType.ToString().Trim();
                        }
                        ProductType.Id = entCol[0].Id;
                        enqid = entCol[0].Id;
                        ProductType["hil_mdmtimestamp"] = obj.Results[i].ModifiedDate != "" && obj.Results[i].ModifiedDate != null ? Convert.ToDateTime(obj.Results[i].ModifiedDate) : Convert.ToDateTime(obj.Results[i].CreatedDate.ToString());
                        service.Update(ProductType);

                        if (obj.Results[i].IsActive)
                        {
                            SetStateRequest setStateRequest = new SetStateRequest()
                            {
                                EntityMoniker = new EntityReference
                                {
                                    Id = ProductType.Id,
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
                                    Id = ProductType.Id,
                                    LogicalName = ProductType.LogicalName,
                                },
                                State = new OptionSetValue(1), //deactive
                                Status = new OptionSetValue(2) //deactive
                            };
                            service.Execute(setStateRequest);
                        }
                        Console.WriteLine("Product Type Updated Count = " + i);
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
        }
        #endregion
        #region Create/update GetEnquiryProductMapping
        public static void GetProductMapping(IOrganizationService service, string URL, String timestamp)
        {
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
                var client = new RestClient(URL + "GetEnquiryProductMapping?enquiryDate=" + timestamp);
                client.Timeout = -1;
                var request = new RestRequest(Method.GET);
                request.AddHeader("Authorization", "Basic RDM2NV9IQVZFTExTOkRFVkQzNjVAMTIzNA==");
                IRestResponse response = client.Execute(request);
                Console.WriteLine(response.Content);
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<GetEnquiryProductMapping>(response.Content);
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
        }
        #endregion
        #region create GetCustomersType from middleware
        public static void GetCustomersType(IOrganizationService service, string URL, String timestamp)
        {
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
                var client = new RestClient(URL + "GetCustomersType?enquiryDate=" + timestamp);
                client.Timeout = -1;
                var request = new RestRequest(Method.GET);
                request.AddHeader("Authorization", "Basic RDM2NV9IQVZFTExTOkRFVkQzNjVAMTIzNA==");
                IRestResponse response = client.Execute(request);
                Console.WriteLine(response.Content);
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<GetPropertyType>(response.Content);
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
        }
        #endregion
        #region create GetPropertyType
        public static void getPropertytypeData(IOrganizationService service, string URL, String timestamp)
        {
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
                var client = new RestClient(URL + "GetPropertyType?enquiryDate=" + timestamp);
                client.Timeout = -1;
                var request = new RestRequest(Method.GET);
                var byteArray = Encoding.ASCII.GetBytes("username:password1234");

                request.AddHeader("Authorization", "Basic RDM2NV9IQVZFTExTOkRFVkQzNjVAMTIzNA==");
                IRestResponse response = client.Execute(request);
                Console.WriteLine(response.Content);
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<GetPropertyType>(response.Content);
                // Console.WriteLine();
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
        }
        #endregion
        #region AdvisoryMaster
        public static void AdvisoryMaster(IOrganizationService _service, string URL, String timestamp)
        {
            try
            {
                TextInfo tInfo = new CultureInfo("en-US", false).TextInfo;
                WorkflowData advdivsetup = new WorkflowData();
                var client = new RestClient(URL + "GetWorkFlowData?enquiryDate=" + timestamp);
                client.Timeout = -1;
                var request = new RestRequest(Method.GET);
                request.AddHeader("Authorization", "Basic RDM2NV9IQVZFTExTOkRFVkQzNjVAMTIzNA==");
                IRestResponse response = client.Execute(request);
                Console.WriteLine(response.Content);
                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<WorkflowData>(response.Content);
                // Console.WriteLine();
                for (int i = 0; i < obj.Results.Count; i++)
                {
                    QueryExpression query = new QueryExpression();
                    //query.ColumnSet = new ColumnSet(false);
                    //query.Criteria = new FilterExpression(LogicalOperator.And);
                    //query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, obj.Results[i].EnquiryTypeCode.Trim());

                    //EntityCollection entCol = _service.RetrieveMultiple(query);
                    //if (entCol.Entities.Count > 0)
                    //{

                    //}

                    // query = new QueryExpression("hil_advisormaster");
                    //query.ColumnSet = new ColumnSet(false);
                    //query.Criteria = new FilterExpression(LogicalOperator.And);
                    //query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, obj.Results[i].UserCode.Trim());
                    //query.Criteria.AddCondition("hil_divisioncode", ConditionOperator.Equal, obj.Results[i].EnquiryTypeCode.Trim());
                    //query.Criteria.AddCondition("hil_producttypecode", ConditionOperator.Equal, obj.Results[i].AdvisoryProductTypeCode.Trim());

                    //entCol = _service.RetrieveMultiple(query);
                    //if (entCol.Entities.Count == 0)
                    {
                        EntityCollection entCol = new EntityCollection();
                        Entity AdvDivisionsetup = new Entity("hil_advisormaster");
                        AdvDivisionsetup["hil_name"] = tInfo.ToTitleCase(obj.Results[i].UserName.ToLower().Trim());
                        AdvDivisionsetup["hil_code"] = obj.Results[i].UserCode.Trim();
                        AdvDivisionsetup["hil_email"] = obj.Results[i].EmailId.ToLower().Trim();
                        AdvDivisionsetup["hil_advisormobilenumber"] = obj.Results[i].MobileNumber.Trim();
                        AdvDivisionsetup["hil_mdmtimestamp"] = (obj.Results[i].ModifiedDate == null) ? Convert.ToDateTime(obj.Results[i].CreatedDate) : Convert.ToDateTime(obj.Results[i].ModifiedDate);
                        #region fetch Enquery type...
                        query = new QueryExpression("hil_enquirytype");
                        query.ColumnSet.AddColumns("hil_enquirytypename", "hil_enquirytypecode");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        query.Criteria.AddCondition("hil_enquirytypecode", ConditionOperator.Equal, obj.Results[i].EnquiryTypeCode.Trim());
                        entCol = _service.RetrieveMultiple(query);
                        if (entCol.Entities.Count > 0)
                        {
                            AdvDivisionsetup["hil_divisioncode"] = new EntityReference("hil_advisoryenquiry", entCol[0].Id);
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
                            AdvDivisionsetup["hil_producttypecode"] = new EntityReference("hil_typeofproduct", entCol[0].Id);
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
                            AdvDivisionsetup["hil_salesoffice"] = new EntityReference("hil_salesoffice", entCol[0].Id);
                        }
                        #endregion
                        Guid AdvDivisionsetupid = _service.Create(AdvDivisionsetup);
                        Console.WriteLine("Created " + i);
                    }
                    //else
                    //{
                    //    Entity AdvDivisionsetup = new Entity("hil_advisormaster");
                    //    AdvDivisionsetup.Id = entCol[0].Id;
                    //    AdvDivisionsetup["hil_name"] = tInfo.ToTitleCase(obj.Results[i].UserName.ToLower().Trim());
                    //    AdvDivisionsetup["hil_code"] = obj.Results[i].UserCode.Trim();
                    //    AdvDivisionsetup["hil_email"] = obj.Results[i].EmailId.ToLower().Trim();
                    //    AdvDivisionsetup["hil_advisormobilenumber"] = obj.Results[i].MobileNumber.Trim();
                    //    AdvDivisionsetup["hil_mdmtimestamp"] = (obj.Results[i].ModifiedDate == null) ? Convert.ToDateTime(obj.Results[i].CreatedDate) : Convert.ToDateTime(obj.Results[i].ModifiedDate);
                    //    #region fetch Enquery type...
                    //    query = new QueryExpression("hil_enquirytype");
                    //    query.ColumnSet.AddColumns("hil_enquirytypename", "hil_enquirytypecode");
                    //    query.Criteria = new FilterExpression(LogicalOperator.And);
                    //    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    //    query.Criteria.AddCondition("hil_enquirytypecode", ConditionOperator.Equal, obj.Results[i].EnquiryTypeCode.Trim());
                    //    entCol = _service.RetrieveMultiple(query);
                    //    if (entCol.Entities.Count > 0)
                    //    {
                    //        AdvDivisionsetup["hil_divisioncode"] = new EntityReference("hil_advisoryenquiry", entCol[0].Id);
                    //    }
                    //    #endregion
                    //    #region fetch Product type...
                    //    query = new QueryExpression("hil_typeofproduct");
                    //    query.ColumnSet = new ColumnSet(false);
                    //    query.Criteria = new FilterExpression(LogicalOperator.And);
                    //    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    //    query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, obj.Results[i].AdvisoryProductTypeCode.Trim());
                    //    entCol = _service.RetrieveMultiple(query);
                    //    if (entCol.Entities.Count > 0)
                    //    {
                    //        AdvDivisionsetup["hil_producttypecode"] = new EntityReference("hil_typeofproduct", entCol[0].Id);
                    //    }
                    //    #endregion
                    //    #region fetch Sales Office...
                    //    query = new QueryExpression("hil_salesoffice");
                    //    query.ColumnSet = new ColumnSet(false);
                    //    query.Criteria = new FilterExpression(LogicalOperator.And);
                    //    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    //    query.Criteria.AddCondition("hil_salesofficecode", ConditionOperator.Equal, obj.Results[i].SaleOfficeCode.Trim());
                    //    entCol = _service.RetrieveMultiple(query);
                    //    if (entCol.Entities.Count > 0)
                    //    {
                    //        AdvDivisionsetup["hil_salesoffice"] = new EntityReference("hil_salesoffice", entCol[0].Id);
                    //    }
                    //    #endregion
                    //    _service.Update(AdvDivisionsetup);
                    //    Console.WriteLine("Updated " + i);
                    //}
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("D365 Internal Error " + ex.Message);
            }
        }
        #endregion
        #region CANCELLATION REASON...
        static public void CancellationReasonandConstructionType(IOrganizationService service)
        {
            try
            {
                //string sUserName = "SFA_Havells";
                //string sPassword = "DEVSFA@1234";
                //String sUrl = "middlewaredev";

                //string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));

                //WebClient webClient = new WebClient();
                //webClient.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
                //var jsonData = webClient.DownloadData(sUrl);


                var client = new RestClient("https://middlewaredev.havells.com:50001/RESTAdapter/MDMService/Core/Lookup/GetLookupMasterData");
                client.Timeout = -1;
                var request = new RestRequest(Method.GET);
                request.AddHeader("Authorization", "Basic U0ZBX0hhdmVsbHM6REVWU0ZBQDEyMzQ=");
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
                        if (entCol.Entities.Count == 0)
                        {
                            Entity _entity = new Entity("hil_cancellationreason");
                            _entity["hil_code"] = look.Code.ToString();
                            _entity["hil_sequence"] = look.Sequence.ToString();
                            _entity["hil_name"] = look.Name.ToString();
                            _entity["hil_lookuptype"] = look.LookupType.ToString();
                            _entity["hil_lookupid"] = look.LookupID.ToString();
                            _entity["hil_mastertypeid"] = look.MasterTypeID.ToString();
                            _entity["hil_mdmtimestamp"] = (look.ModifiedDate != "" && look.ModifiedDate != null) ? Convert.ToDateTime(look.ModifiedDate) : Convert.ToDateTime(look.CreatedDate);

                            service.Create(_entity);
                        }
                        else if (entCol.Entities.Count > 0)
                        {
                            Entity _entity = new Entity("hil_cancellationreason");
                            _entity["hil_code"] = look.Code.ToString();
                            _entity["hil_sequence"] = look.Sequence.ToString();
                            _entity["hil_name"] = look.Name.ToString();
                            _entity["hil_lookuptype"] = look.LookupType.ToString();
                            _entity["hil_lookupid"] = look.LookupID.ToString();
                            _entity["hil_mastertypeid"] = look.MasterTypeID.ToString();
                            _entity["hil_mdmtimestamp"] = (look.ModifiedDate != "" && look.ModifiedDate != null) ? Convert.ToDateTime(look.ModifiedDate) : Convert.ToDateTime(look.CreatedDate);

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
        #endregion
        #region
        //static public void EnquirySendtoSFA(IOrganizationService service) {
        //    QueryExpression query = new QueryExpression("hil_homeadvisoryline");
        //    query.ColumnSet = new ColumnSet(true);
        //    query.Criteria = new FilterExpression(LogicalOperator.And);
        //    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
        //    query.Criteria.AddCondition("hil_syncwithsfa", ConditionOperator.Equal, false);
        //    EntityCollection entCol = service.RetrieveMultiple(query);
        //    foreach(Entity postImage in entCol.Entities)
        //    {
        //        EnquiryReq _enq = new EnquiryReq();
        //        EnquiryLineReq _enqLine = new EnquiryLineReq();
        //        EntityReference _EnqueryRef = postImage.GetAttributeValue<EntityReference>("hil_advisoryenquery");
        //        Entity _enqueryEntity = service.Retrieve(_EnqueryRef.LogicalName, _EnqueryRef.Id, new ColumnSet(true));

        //        _enqLine.EnquiryId = postImage.Contains("hil_name") ? postImage.GetAttributeValue<String>("hil_name") : "";
        //        _enqLine.ParentEnquiryId = postImage.Contains("hil_advisoryenquery") ? postImage.GetAttributeValue<EntityReference>("hil_advisoryenquery").Name : "";
        //        _enqLine.EnquiryLineGuid = postImage.Id.ToString();
        //        if (postImage.Contains("hil_typeofenquiiry"))
        //        {
        //            Entity _enqueryTypeEntity = service.Retrieve(postImage.GetAttributeValue<EntityReference>("hil_typeofenquiiry").LogicalName, postImage.GetAttributeValue<EntityReference>("hil_typeofenquiiry").Id, new ColumnSet("hil_enquirytypecode"));
        //            _enqLine.EnquiryTypeCode = _enqueryTypeEntity.Contains("hil_enquirytypecode") ? _enqueryTypeEntity.GetAttributeValue<int>("hil_enquirytypecode").ToString() : "";
        //        }
        //        if (postImage.Contains("hil_typeofproduct"))
        //        {
        //            Entity _ProductTypeEntity = service.Retrieve(postImage.GetAttributeValue<EntityReference>("hil_typeofproduct").LogicalName, postImage.GetAttributeValue<EntityReference>("hil_typeofproduct").Id, new ColumnSet("hil_code"));
        //            _enqLine.ProducTypeCode = _ProductTypeEntity.Contains("hil_code") ? _ProductTypeEntity.GetAttributeValue<String>("hil_code") : "";
        //        }
        //        if (postImage.Contains("hil_assignedadvisor"))
        //        {
        //            Entity _AdvisorEntity = service.Retrieve(postImage.GetAttributeValue<EntityReference>("hil_assignedadvisor").LogicalName, postImage.GetAttributeValue<EntityReference>("hil_assignedadvisor").Id, new ColumnSet("hil_code"));
        //            _enqLine.AdvisorCode = _AdvisorEntity.Contains("hil_code") ? _AdvisorEntity.GetAttributeValue<String>("hil_code") : "";
        //        }
        //        _enqLine.EnquiryStatus = postImage.Contains("hil_enquirystauts") ? postImage.GetAttributeValue<OptionSetValue>("hil_enquirystauts").Value.ToString() : "";
        //        _enqLine.AppointmentId = postImage.Contains("hil_appointmentid") ? postImage.GetAttributeValue<String>("hil_appointmentid") : "";
        //        _enqLine.AppointmentType = postImage.Contains("hil_appointmenttypes") ? postImage.GetAttributeValue<OptionSetValue>("hil_appointmenttypes").Value.ToString() : "";
        //        if (postImage.Contains("hil_appointmentdate"))
        //        {
        //            DateTime dateTime = postImage.GetAttributeValue<DateTime>("hil_appointmentdate");
        //            string _datestr = dateTime.Year.ToString() + dateTime.Month.ToString().PadLeft(2, '0') + dateTime.Day.ToString().PadLeft(2, '0') + dateTime.Hour.ToString().PadLeft(2, '0') + dateTime.Minute.ToString().PadLeft(2, '0') + dateTime.Second.ToString().PadLeft(2, '0');
        //            _enqLine.AppointmentDate = _datestr;
        //            dateTime = postImage.GetAttributeValue<DateTime>("hil_appointmentdate").AddMinutes(30);
        //            _datestr = dateTime.Year.ToString() + dateTime.Month.ToString().PadLeft(2, '0') + dateTime.Day.ToString().PadLeft(2, '0') + dateTime.Hour.ToString().PadLeft(2, '0') + dateTime.Minute.ToString().PadLeft(2, '0') + dateTime.Second.ToString().PadLeft(2, '0');
        //            _enqLine.AppointmentEndDate = _datestr;
        //        }
        //        else
        //        {
        //            _enqLine.AppointmentDate = "19000101000000";
        //            _enqLine.AppointmentEndDate = "19000101000000";
        //        }

        //        _enqLine.AppointmentStatus = postImage.Contains("hil_appointmentstatus") ? postImage.GetAttributeValue<OptionSetValue>("hil_appointmentstatus").Value.ToString() : "1";
        //        _enqLine.VideoCallURL = postImage.Contains("hil_videocallurl") ? postImage.GetAttributeValue<String>("hil_videocallurl") : "";
        //        _enqLine.CustomerRemarks = postImage.Contains("hil_customerremark") ? postImage.GetAttributeValue<String>("hil_customerremark") : "";
        //        _enqLine.PinCode = postImage.Contains("hil_pincode") ? postImage.GetAttributeValue<EntityReference>("hil_pincode").Name.ToString() : "";
        //        List<EnquiryLineReq> EnqLines = new List<EnquiryLineReq>();
        //        EnqLines.Add(_enqLine);
        //        _enq.EnquiryId = _enqueryEntity.Contains("hil_name") ? _enqueryEntity.GetAttributeValue<String>("hil_name") : "";
        //        _enq.Area = _enqueryEntity.Contains("hil_areasqrt") ? _enqueryEntity.GetAttributeValue<String>("hil_areasqrt") : "";
        //        if (_enqueryEntity.Contains("hil_typeofcustomer"))
        //        {
        //            Entity _CustomerTypeEntity = service.Retrieve(_enqueryEntity.GetAttributeValue<EntityReference>("hil_typeofcustomer").LogicalName, _enqueryEntity.GetAttributeValue<EntityReference>("hil_typeofcustomer").Id, new ColumnSet("hil_code"));
        //            _enq.CustomerTypeCode = _CustomerTypeEntity.Contains("hil_code") ? _CustomerTypeEntity.GetAttributeValue<String>("hil_code") : "";
        //        }
        //        if (_enqueryEntity.Contains("hil_propertytype"))
        //        {
        //            Entity _propertyTypeEntity = service.Retrieve(_enqueryEntity.GetAttributeValue<EntityReference>("hil_propertytype").LogicalName, _enqueryEntity.GetAttributeValue<EntityReference>("hil_propertytype").Id, new ColumnSet("hil_code"));
        //            _enq.PropertyTypeCode = _propertyTypeEntity.Contains("hil_code") ? _propertyTypeEntity.GetAttributeValue<String>("hil_code") : "";
        //        }
        //        if (_enqueryEntity.Contains("hil_constructiontype"))
        //        {
        //            Entity _constructionTypeEntity = service.Retrieve(_enqueryEntity.GetAttributeValue<EntityReference>("hil_constructiontype").LogicalName, _enqueryEntity.GetAttributeValue<EntityReference>("hil_constructiontype").Id, new ColumnSet("hil_code"));
        //            _enq.ConstructionTypeCode = _constructionTypeEntity.Contains("hil_code") ? _constructionTypeEntity.GetAttributeValue<String>("hil_code") : "";
        //        }
        //        _enq.Rooftop = _enqueryEntity.GetAttributeValue<bool>("hil_rooftop") ? "true" : "false";
        //        _enq.AssetType = _enqueryEntity.Contains("hil_assettype") ? _enqueryEntity.GetAttributeValue<String>("hil_assettype") : "";
        //        _enq.CustomerName = _enqueryEntity.Contains("hil_customer") ? _enqueryEntity.GetAttributeValue<EntityReference>("hil_customer").Name : "";
        //        _enq.MobileNumber = _enqueryEntity.Contains("hil_mobilenumber") ? _enqueryEntity.GetAttributeValue<String>("hil_mobilenumber") : "";
        //        _enq.EmailId = _enqueryEntity.Contains("hil_emailid") ? _enqueryEntity.GetAttributeValue<String>("hil_emailid") : "";
        //        if (_enqueryEntity.Contains("hil_city"))
        //        {
        //            Entity _cityEntity = service.Retrieve(_enqueryEntity.GetAttributeValue<EntityReference>("hil_city").LogicalName, _enqueryEntity.GetAttributeValue<EntityReference>("hil_city").Id, new ColumnSet("hil_citycode"));
        //            _enq.CityCode = _cityEntity.Contains("hil_citycode") ? _cityEntity.GetAttributeValue<String>("hil_citycode") : "";
        //        }
        //        if (_enqueryEntity.Contains("hil_state"))
        //        {
        //            Entity _stateEntity = service.Retrieve(_enqueryEntity.GetAttributeValue<EntityReference>("hil_state").LogicalName, _enqueryEntity.GetAttributeValue<EntityReference>("hil_state").Id, new ColumnSet("hil_statecode"));
        //            _enq.StateCode = _stateEntity.Contains("hil_statecode") ? _stateEntity.GetAttributeValue<String>("hil_statecode") : "";
        //        }
        //        _enq.EnquiryGuid = _enqueryEntity.Id.ToString();
        //        _enq.PINCode = _enqueryEntity.Contains("hil_pincode") ? _enqueryEntity.GetAttributeValue<EntityReference>("hil_pincode").Name : "";
        //        _enq.TDS = _enqueryEntity.Contains("hil_tds") ? _enqueryEntity.GetAttributeValue<String>("hil_tds") : "";
        //        _enq.EnquiryLines = EnqLines;
        //        _enq.SourceOfCreation = "1";
        //        EnqueryRequestReq enqReq = new EnqueryRequestReq();
        //        enqReq.Data = _enq;

        //        string data = JsonConvert.SerializeObject(enqReq);

        //        var client = new RestClient("https://middlewaredev.havells.com:50001/RESTAdapter/CMSService/HomeAdvisor/SaveEnquiryDetails");
        //        client.Timeout = -1;
        //        var request = new RestRequest(Method.POST);
        //        request.AddHeader("Authorization", "Basic RDM2NV9IYXZlbGxzOkRFVkQzNjVAMTIzNA==");
        //        request.AddHeader("Content-Type", "application/json");
        //        request.AddHeader("Cookie", "JSESSIONID=9bTRiem7PB9UAgthLrIEPd12LBPxdwGenHkA_SAPNRuxfvJoXdvQT2lLsHROL2TZ; saplb_*=(J2EE7969920)7969951");
        //        request.AddParameter("application/json",data, ParameterType.RequestBody);
        //        IRestResponse response = client.Execute(request);
        //        Console.WriteLine(response.Content);

        //        EnqueryResponseReq result = JsonConvert.DeserializeObject<EnqueryResponseReq>(response.Content);

        //        if (result.IsSuccess)
        //        {
        //            Entity _line = new Entity("hil_homeadvisoryline");
        //            _line.Id = postImage.Id;
        //            _line["hil_syncwithsfa"] = true;
        //            service.Update(_line);
        //        }
        //        else
        //        {
        //            Console.WriteLine("Error : " + result.Message);
        //        }
        //    }

        //}
        #endregion
        #region SaveEnquiryDocument
        public static void SaveEnquiryDocument(IOrganizationService service)
        {
            try
            {
                QueryExpression _query = new QueryExpression("hil_attachment");
                _query.ColumnSet = new ColumnSet(true);
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                _query.Criteria.AddCondition("hil_syncwithsfa", ConditionOperator.Equal, false);
                EntityCollection Found = service.RetrieveMultiple(_query);
                foreach (Entity postImage in Found.Entities)
                {
                    string cc = postImage.GetAttributeValue<OptionSetValue>("hil_sourceofdocument").Value.ToString();
                    if (postImage.GetAttributeValue<EntityReference>("regardingobjectid").LogicalName == "hil_homeadvisoryline" && postImage.GetAttributeValue<OptionSetValue>("hil_sourceofdocument").Value != 1)
                    {
                        Attachments attachments = new Attachments();
                        Entity _enqueryEntity = service.Retrieve(postImage.GetAttributeValue<EntityReference>("regardingobjectid").LogicalName, postImage.GetAttributeValue<EntityReference>("regardingobjectid").Id, new ColumnSet("hil_name"));
                        attachments.EnquiryId = _enqueryEntity.GetAttributeValue<String>("hil_name");
                        attachments.FileName = postImage.GetAttributeValue<string>("subject");
                        attachments.Subject = postImage.GetAttributeValue<string>("subject");
                        attachments.BlobURL = postImage.GetAttributeValue<String>("hil_docurl").ToString();
                        attachments.MIMEType = "";
                        attachments.DocGuid = _enqueryEntity.Id.ToString();
                        attachments.IsDelete = postImage.GetAttributeValue<bool>("hil_isdeleted");
                        attachments.DocType = postImage.GetAttributeValue<EntityReference>("hil_documenttype").Name;
                        attachments.Size = Convert.ToInt32(postImage.GetAttributeValue<double>("hil_docsize")).ToString();
                        List<Attachments> attachmentsList = new List<Attachments>();
                        attachmentsList.Add(attachments);
                        AttachmentsList data = new AttachmentsList();
                        data.Data = attachmentsList;
                        string ss = JsonConvert.SerializeObject(data);
                        var client = new RestClient("https://middlewaredev.havells.com:50001/RESTAdapter/CMSService/HomeAdvisor/SaveEnquiryDocument");
                        client.Timeout = -1;
                        var request = new RestRequest(Method.POST);
                        request.AddHeader("Authorization", "Basic RDM2NV9IYXZlbGxzOkRFVkQzNjVAMTIzNA==");
                        request.AddHeader("Content-Type", "application/json");
                        request.AddParameter("application/json", ss, ParameterType.RequestBody);
                        IRestResponse response = client.Execute(request);
                        Console.WriteLine(response.Content);

                        EnqueryResponseReq result = JsonConvert.DeserializeObject<EnqueryResponseReq>(response.Content);

                        if (result.IsSuccess)
                        {
                            Entity _line = new Entity("hil_attachment");
                            _line.Id = postImage.Id;
                            _line["hil_syncwithsfa"] = true;
                            service.Update(_line);
                        }
                        else
                        {
                            Console.WriteLine("Error : " + result.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex.Message);
            }
        }
        #endregion
        #region SaveEnquiryDetails
        public static void SaveEnquiryDetails(IOrganizationService service)
        {
            try
            {
                QueryExpression _query = new QueryExpression("hil_homeadvisoryline");
                _query.ColumnSet = new ColumnSet(true);
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                _query.Criteria.AddCondition("hil_syncwithsfa", ConditionOperator.Equal, false);
                EntityCollection Found = service.RetrieveMultiple(_query);

                Entity postImage = service.Retrieve("hil_homeadvisoryline", new Guid("0D6B7043-168D-EB11-B1AC-6045BD724A95"), new ColumnSet(true));
                //foreach (Entity postImage in Found.Entities)
                {
                    EnquirySendSFA _enq = new EnquirySendSFA();
                    EnquiryLine _enqLine = new EnquiryLine();
                    EntityReference _EnqueryRef = postImage.GetAttributeValue<EntityReference>("hil_advisoryenquery");
                    Entity _enqueryEntity = service.Retrieve(_EnqueryRef.LogicalName, _EnqueryRef.Id, new ColumnSet(true));

                    _enqLine.EnquiryId = postImage.Contains("hil_name") ? postImage.GetAttributeValue<String>("hil_name") : "";
                    _enqLine.ParentEnquiryId = postImage.Contains("hil_advisoryenquery") ? postImage.GetAttributeValue<EntityReference>("hil_advisoryenquery").Name : "";
                    _enqLine.EnquiryLineGuid = postImage.Id.ToString();
                    if (postImage.Contains("hil_typeofenquiiry"))
                    {
                        Entity _enqueryTypeEntity = service.Retrieve(postImage.GetAttributeValue<EntityReference>("hil_typeofenquiiry").LogicalName, postImage.GetAttributeValue<EntityReference>("hil_typeofenquiiry").Id, new ColumnSet("hil_enquirytypecode"));
                        _enqLine.EnquiryTypeCode = _enqueryTypeEntity.Contains("hil_enquirytypecode") ? _enqueryTypeEntity.GetAttributeValue<int>("hil_enquirytypecode").ToString() : "";
                    }
                    if (postImage.Contains("hil_typeofproduct"))
                    {
                        Entity _ProductTypeEntity = service.Retrieve(postImage.GetAttributeValue<EntityReference>("hil_typeofproduct").LogicalName, postImage.GetAttributeValue<EntityReference>("hil_typeofproduct").Id, new ColumnSet("hil_code"));
                        _enqLine.ProducTypeCode = _ProductTypeEntity.Contains("hil_code") ? _ProductTypeEntity.GetAttributeValue<String>("hil_code") : "";
                    }
                    if (postImage.Contains("hil_assignedadvisor"))
                    {
                        Entity _AdvisorEntity = service.Retrieve(postImage.GetAttributeValue<EntityReference>("hil_assignedadvisor").LogicalName, postImage.GetAttributeValue<EntityReference>("hil_assignedadvisor").Id, new ColumnSet("hil_code"));
                        _enqLine.AdvisorCode = _AdvisorEntity.Contains("hil_code") ? _AdvisorEntity.GetAttributeValue<String>("hil_code") : "";
                    }
                    _enqLine.EnquiryStatus = postImage.Contains("hil_enquirystauts") ? postImage.GetAttributeValue<OptionSetValue>("hil_enquirystauts").Value.ToString() : "";
                    _enqLine.AppointmentId = postImage.Contains("hil_appointmentid") ? postImage.GetAttributeValue<String>("hil_appointmentid") : "";
                    _enqLine.AppointmentType = postImage.Contains("hil_appointmenttypes") ? postImage.GetAttributeValue<OptionSetValue>("hil_appointmenttypes").Value.ToString() : "";
                    if (postImage.Contains("hil_appointmentdate"))
                    {


                        DateTime dateTime = postImage.GetAttributeValue<DateTime>("hil_appointmentdate").AddMinutes(330);
                        string _datestr = dateTime.Year.ToString() + dateTime.Month.ToString().PadLeft(2, '0') + dateTime.Day.ToString().PadLeft(2, '0') + dateTime.Hour.ToString().PadLeft(2, '0') + dateTime.Minute.ToString().PadLeft(2, '0') + dateTime.Second.ToString().PadLeft(2, '0');
                        _enqLine.AppointmentDate = _datestr;
                    }
                    else
                    {
                        _enqLine.AppointmentDate = "19000101000000";
                    }
                    if (postImage.Contains("hil_appointmentdate"))
                    {
                        DateTime dateTime = postImage.GetAttributeValue<DateTime>("hil_appointmentenddate").AddMinutes(330);
                        string _datestr = dateTime.Year.ToString() + dateTime.Month.ToString().PadLeft(2, '0') + dateTime.Day.ToString().PadLeft(2, '0') + dateTime.Hour.ToString().PadLeft(2, '0') + dateTime.Minute.ToString().PadLeft(2, '0') + dateTime.Second.ToString().PadLeft(2, '0');
                        _enqLine.AppointmentEndDate = _datestr;
                    }
                    else
                    {
                        _enqLine.AppointmentEndDate = "19000101000000";
                    }
                    _enqLine.AppointmentStatus = postImage.Contains("hil_appointmentstatus") ? postImage.GetAttributeValue<OptionSetValue>("hil_appointmentstatus").Value.ToString() : "";
                    _enqLine.VideoCallURL = postImage.Contains("hil_videocallurl") ? postImage.GetAttributeValue<String>("hil_videocallurl") : "";
                    _enqLine.CustomerRemarks = postImage.Contains("hil_customerremark") ? postImage.GetAttributeValue<String>("hil_customerremark") : "";
                    _enqLine.PinCode = postImage.Contains("hil_pincode") ? postImage.GetAttributeValue<EntityReference>("hil_pincode").Name.ToString() : "";
                    List<EnquiryLine> EnqLines = new List<EnquiryLine>();
                    EnqLines.Add(_enqLine);
                    _enq.EnquiryId = _enqueryEntity.Contains("hil_name") ? _enqueryEntity.GetAttributeValue<String>("hil_name") : "";
                    _enq.Area = _enqueryEntity.Contains("hil_areasqrt") ? _enqueryEntity.GetAttributeValue<String>("hil_areasqrt") : "";
                    if (_enqueryEntity.Contains("hil_typeofcustomer"))
                    {
                        Entity _CustomerTypeEntity = service.Retrieve(_enqueryEntity.GetAttributeValue<EntityReference>("hil_typeofcustomer").LogicalName, _enqueryEntity.GetAttributeValue<EntityReference>("hil_typeofcustomer").Id, new ColumnSet("hil_code"));
                        _enq.CustomerTypeCode = _CustomerTypeEntity.Contains("hil_code") ? _CustomerTypeEntity.GetAttributeValue<String>("hil_code") : "";
                    }
                    if (_enqueryEntity.Contains("hil_propertytype"))
                    {
                        Entity _propertyTypeEntity = service.Retrieve(_enqueryEntity.GetAttributeValue<EntityReference>("hil_propertytype").LogicalName, _enqueryEntity.GetAttributeValue<EntityReference>("hil_propertytype").Id, new ColumnSet("hil_code"));
                        _enq.PropertyTypeCode = _propertyTypeEntity.Contains("hil_code") ? _propertyTypeEntity.GetAttributeValue<String>("hil_code") : "";
                    }
                    if (_enqueryEntity.Contains("hil_constructiontype"))
                    {
                        Entity _constructionTypeEntity = service.Retrieve(_enqueryEntity.GetAttributeValue<EntityReference>("hil_constructiontype").LogicalName, _enqueryEntity.GetAttributeValue<EntityReference>("hil_constructiontype").Id, new ColumnSet("hil_code"));
                        _enq.ConstructionTypeCode = _constructionTypeEntity.Contains("hil_code") ? _constructionTypeEntity.GetAttributeValue<String>("hil_code") : "";
                    }
                    _enq.Rooftop = _enqueryEntity.GetAttributeValue<bool>("hil_rooftop") ? "true" : "false";
                    _enq.AssetType = _enqueryEntity.Contains("hil_assettype") ? _enqueryEntity.GetAttributeValue<String>("hil_assettype") : "";
                    _enq.CustomerName = _enqueryEntity.Contains("hil_customer") ? _enqueryEntity.GetAttributeValue<EntityReference>("hil_customer").Name : "";
                    _enq.MobileNumber = _enqueryEntity.Contains("hil_mobilenumber") ? _enqueryEntity.GetAttributeValue<String>("hil_mobilenumber") : "";
                    _enq.EmailId = _enqueryEntity.Contains("hil_emailid") ? _enqueryEntity.GetAttributeValue<String>("hil_emailid") : "";
                    if (_enqueryEntity.Contains("hil_city"))
                    {
                        Entity _cityEntity = service.Retrieve(_enqueryEntity.GetAttributeValue<EntityReference>("hil_city").LogicalName, _enqueryEntity.GetAttributeValue<EntityReference>("hil_city").Id, new ColumnSet("hil_citycode"));
                        _enq.CityCode = _cityEntity.Contains("hil_citycode") ? _cityEntity.GetAttributeValue<String>("hil_citycode") : "";
                    }
                    if (_enqueryEntity.Contains("hil_state"))
                    {
                        Entity _stateEntity = service.Retrieve(_enqueryEntity.GetAttributeValue<EntityReference>("hil_state").LogicalName, _enqueryEntity.GetAttributeValue<EntityReference>("hil_state").Id, new ColumnSet("hil_statecode"));
                        _enq.StateCode = _stateEntity.Contains("hil_statecode") ? _stateEntity.GetAttributeValue<String>("hil_statecode") : "";
                    }
                    _enq.EnquiryGuid = _enqueryEntity.Id.ToString();
                    _enq.PINCode = _enqueryEntity.Contains("hil_pincode") ? _enqueryEntity.GetAttributeValue<EntityReference>("hil_pincode").Name : "";
                    _enq.TDS = _enqueryEntity.Contains("hil_tds") ? _enqueryEntity.GetAttributeValue<String>("hil_tds") : "";
                    _enq.EnquiryLines = EnqLines;

                    _enq.SourceOfCreation = _enqueryEntity.Contains("hil_sourceofcreation") ? _enqueryEntity.GetAttributeValue<OptionSetValue>("hil_sourceofcreation").Value.ToString() : "";

                    EnqueryRequest enqReq = new EnqueryRequest();

                    enqReq.Data = _enq;
                    string data = JsonConvert.SerializeObject(enqReq);

                    var client = new RestClient("https://middlewaredev.havells.com:50001/RESTAdapter/CMSService/HomeAdvisor/SaveEnquiryDetails");
                    client.Timeout = -1;
                    var request = new RestRequest(Method.POST);
                    request.AddHeader("Authorization", "Basic RDM2NV9IYXZlbGxzOkRFVkQzNjVAMTIzNA==");
                    request.AddHeader("Content-Type", "application/json");
                    request.AddParameter("application/json", data, ParameterType.RequestBody);
                    IRestResponse response = client.Execute(request);
                    Console.WriteLine(response.Content);

                    EnqueryResponseReq result = JsonConvert.DeserializeObject<EnqueryResponseReq>(response.Content);

                    if (result.IsSuccess)
                    {
                        Entity _line = new Entity("hil_homeadvisoryline");
                        _line.Id = postImage.Id;
                        _line["hil_syncwithsfa"] = true;
                        service.Update(_line);
                    }
                    else
                    {
                        Console.WriteLine("Error : " + result.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error " + ex.Message);
            }
        }
        #endregion
    }

    public class LookUpRes
    {
        public int LookupID { get; set; }
        public string Name { get; set; }
        public string Code { get; set; }
        public int Sequence { get; set; }
        public int MasterTypeID { get; set; }
        public string LookupType { get; set; }
        public String CreatedDate { get; set; }
        public int CreatedBy { get; set; }
        public String ModifiedDate { get; set; }
        public int? ModifiedBy { get; set; }
        public bool IsActive { get; set; }
    }

    public class Root
    {
        public object Result { get; set; }
        public List<LookUpRes> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }

    public class EnquiryLineReq
    {
        public string ParentEnquiryId { get; set; }
        public string EnquiryId { get; set; }
        public string EnquiryLineGuid { get; set; }
        public string EnquiryTypeCode { get; set; }
        public string ProducTypeCode { get; set; }
        public string PinCode { get; set; }
        public string AdvisorCode { get; set; }
        public string CustomerRemarks { get; set; }
        public string AppointmentType { get; set; }
        public string AppointmentDate { get; set; }
        public string AppointmentEndDate { get; set; }
        public string AppointmentId { get; set; }
        public string AppointmentStatus { get; set; }
        public string VideoCallURL { get; set; }
        public string EnquiryStatus { get; set; }
    }
    public class EnquiryReq
    {
        public string EnquiryId { get; set; }
        public string CustomerName { get; set; }
        public string MobileNumber { get; set; }
        public string EmailId { get; set; }
        public string CustomerTypeCode { get; set; }
        public string PropertyTypeCode { get; set; }
        public string ConstructionTypeCode { get; set; }
        public string StateCode { get; set; }
        public string CityCode { get; set; }
        public string PINCode { get; set; }
        public string Area { get; set; }
        public string Rooftop { get; set; }
        public string TDS { get; set; }
        public string AssetType { get; set; }
        public string SourceOfCreation { get; set; }
        public string EnquiryGuid { get; set; }
        public List<EnquiryLineReq> EnquiryLines { get; set; }
    }
    public class EnqueryRequestReq
    {
        public EnquiryReq Data { get; set; }
    }
    public class EnqueryResponseReq
    {
        public int ResponseStatusCode { get; set; }
        public string Message { get; set; }
        public object ServiceToken { get; set; }
        public List<object> Results { get; set; }
        public bool Result { get; set; }
        public bool IsDeleteInsert { get; set; }
        public object EarlierChannelType { get; set; }
        public object MaxLastSyncDate { get; set; }
        public bool IsSuccess { get; set; }
    }
    public class EnquiryLine
    {
        public string ParentEnquiryId { get; set; }
        public string EnquiryId { get; set; }
        public string EnquiryLineGuid { get; set; }
        public string EnquiryTypeCode { get; set; }
        public string ProducTypeCode { get; set; }
        public string PinCode { get; set; }
        public string AdvisorCode { get; set; }
        public string CustomerRemarks { get; set; }
        public string AppointmentType { get; set; }
        public string AppointmentDate { get; set; }
        public string AppointmentEndDate { get; set; }
        public string AppointmentId { get; set; }
        public string AppointmentStatus { get; set; }
        public string VideoCallURL { get; set; }
        public string EnquiryStatus { get; set; }
    }
    public class EnquirySendSFA
    {
        public string EnquiryId { get; set; }
        public string CustomerName { get; set; }
        public string MobileNumber { get; set; }
        public string EmailId { get; set; }
        public string CustomerTypeCode { get; set; }
        public string PropertyTypeCode { get; set; }
        public string ConstructionTypeCode { get; set; }
        public string StateCode { get; set; }
        public string CityCode { get; set; }
        public string PINCode { get; set; }
        public string Area { get; set; }
        public string Rooftop { get; set; }
        public string TDS { get; set; }
        public string AssetType { get; set; }
        public string SourceOfCreation { get; set; }
        public string EnquiryGuid { get; set; }
        public List<EnquiryLine> EnquiryLines { get; set; }
    }
    public class EnqueryRequest
    {
        public EnquirySendSFA Data { get; set; }
    }
    public class EnqueryResponse
    {
        public int ResponseStatusCode { get; set; }
        public string Message { get; set; }
        public object ServiceToken { get; set; }
        public List<object> Results { get; set; }
        public bool Result { get; set; }
        public bool IsDeleteInsert { get; set; }
        public object EarlierChannelType { get; set; }
        public object MaxLastSyncDate { get; set; }
        public bool IsSuccess { get; set; }
    }
    public class Attachments
    {
        public string EnquiryId { get; set; }
        public string FileName { get; set; }
        public string Subject { get; set; }
        public string BlobURL { get; set; }
        public string MIMEType { get; set; }
        public string DocType { get; set; }
        public string Size { get; set; }
        public string DocGuid { get; set; }
        public bool IsDelete { get; set; }
    }
    public class AttachmentsList
    {
        public List<Attachments> Data { get; set; }
    }
    public class AttachmentResponse
    {
        public int ResponseStatusCode { get; set; }
        public string Message { get; set; }
        public object ServiceToken { get; set; }
        public List<object> Results { get; set; }
        public bool Result { get; set; }
        public bool IsDeleteInsert { get; set; }
        public object EarlierChannelType { get; set; }
        public object MaxLastSyncDate { get; set; }
        public bool IsSuccess { get; set; }
    }
    public class EnquryType
    {
        public string Result { get; set; }
        public string Success { get; set; }
        public string Message { get; set; }
        public string ErrorCode { get; set; }
        public List<EnquryTypeDate> Results { get; set; }
    }
    public class EnquryTypeDate
    {
        public string EnquiryTypeCode { get; set; }
        public string EnquiryTypeDesc { get; set; }
        public string CreatedDate { get; set; }
        public string CreatedBy { get; set; }
        public string ModifiedDate { get; set; }
        public string ModifiedBy { get; set; }
        public bool IsActive { get; set; }
    }
    public class ProductType
    {
        public string Result { get; set; }
        public string Success { get; set; }
        public string Message { get; set; }
        public string ErrorCode { get; set; }
        public List<ProducttypeResult> Results { get; set; }
    }
    public class ProducttypeResult
    {
        public string AdvisoryProductTypeCode { get; set; }
        public string AdvisoryProductType { get; set; }
        public string CreatedDate { get; set; }
        public string CreatedBy { get; set; }
        public string ModifiedDate { get; set; }
        public string ModifiedBy { get; set; }
        public bool IsActive { get; set; }
    }
    class GetEnquiryProductMapping
    {
        public string Result { get; set; }
        public string Success { get; set; }
        public string Message { get; set; }
        public string ErrorCode { get; set; }
        public List<EnquiryTypeprodmapping> Results { get; set; }
    }
    public class EnquiryTypeprodmapping
    {
        public string EnquiryTypeCode { get; set; }
        public string AdvisoryProductTypeCode { get; set; }
        public string CreatedDate { get; set; }
        public string CreatedBy { get; set; }
        public string ModifiedDate { get; set; }
        public string ModifiedBy { get; set; }
        public bool IsActive { get; set; }
    }
    public class GetPropertyType
    {
        public string Result { get; set; }
        public string Success { get; set; }
        public string Message { get; set; }
        public string ErrorCode { get; set; }
        public List<PropertyType> Results { get; set; }
    }
    public class PropertyType
    {
        public string LookupID { get; set; }
        public string Name { get; set; }
        public string Code { get; set; }
        public string Sequence { get; set; }
        public string MasterTypeID { get; set; }
        public string LookupType { get; set; }
        public string CreatedDate { get; set; }
        public string CreatedBy { get; set; }
        public string ModifiedDate { get; set; }
        public string ModifiedBy { get; set; }
        public bool IsActive { get; set; }
    }
    class WorkflowData
    {
        public string Result { get; set; }
        public string Success { get; set; }
        public string Message { get; set; }
        public string ErrorCode { get; set; }
        public List<AdvisordivisionSetup> Results { get; set; }
    }
    public class AdvisordivisionSetup
    {
        public string EnquiryTypeCode { get; set; }
        public string AdvisoryProductTypeCode { get; set; }
        public string UserCode { get; set; }
        public string SaleOfficeCode { get; set; }
        public string CreatedDate { get; set; }
        public string CreatedBy { get; set; }
        public string ModifiedBy { get; set; }
        public string ModifiedDate { get; set; }
        public string IsActive { get; set; }
        public string WorkFlowDataId { get; set; }
        public string Workflow { get; set; }
        public string Isactive { get; set; }
        public string UserName { get; set; }
        public string MobileNumber { get; set; }
        public string EmailId { get; set; }

    }
}
