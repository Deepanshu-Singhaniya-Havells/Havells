using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HomeAdvisory
{
    public static class D365toSFA
    {
        public static void SaveEnquiryDocument(IOrganizationService service)
        {
            try
            {
                Console.WriteLine("SaveEnquiryDocument....");
                QueryExpression _query = new QueryExpression("hil_attachment");
                _query.ColumnSet = new ColumnSet(true);
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                _query.Criteria.AddCondition("hil_syncwithsfa", ConditionOperator.Equal, false);
                _query.Criteria.AddCondition("hil_documenttype", ConditionOperator.Equal, new Guid("ecee219b-56be-eb11-bacc-6045bd72b765"));
                EntityCollection Found = service.RetrieveMultiple(_query);
                Console.WriteLine("Found .... " + Found.Entities.Count);
                foreach (Entity postImage in Found.Entities)
                {
                    string hil_sourceofdocument = postImage.GetAttributeValue<OptionSetValue>("hil_sourceofdocument").Value.ToString();
                    if (postImage.GetAttributeValue<EntityReference>("regardingobjectid").LogicalName == "hil_homeadvisoryline" && hil_sourceofdocument != "1")
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
                        Integration integration = common.IntegrationConfiguration(service, "SaveEnquiryDocumentToSFA");
                        string _authInfo = integration.Auth;
                        _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));
                        String sUrl = integration.uri;
                        var client = new RestClient(sUrl);
                        client.Timeout = -1;
                        var request = new RestRequest(Method.POST);
                        request.AddHeader("Authorization", "Basic " + _authInfo);
                        request.AddHeader("Content-Type", "application/json");
                        request.AddParameter("application/json", ss, ParameterType.RequestBody);
                        IRestResponse response = client.Execute(request);
                        Console.WriteLine(response.Content);
                        EnqueryResponseReq result = JsonConvert.DeserializeObject<EnqueryResponseReq>(response.Content);
                        Console.WriteLine(result.IsSuccess);
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
            Console.WriteLine("Document pushed Sucessfully....");
        }
        public static void SaveEnquiryDetails(IOrganizationService service)
        {
            Console.WriteLine("Started");
            try
            {
                String _query = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='hil_homeadvisoryline'>
                                        <attribute name='hil_name' />
                                        <attribute name='hil_typeofproduct' />
                                        <attribute name='hil_typeofenquiiry' />
                                        <attribute name='hil_enquirystauts' />
                                        <attribute name='hil_assignedadvisor' />
                                        <attribute name='hil_appointmenttypes' />
                                        <attribute name='hil_appointmentstatus' />
                                        <attribute name='hil_appointmentdate' />
                                        <attribute name='createdon' />
                                        <attribute name='hil_mobilenumber' />
                                        <attribute name='hil_homeadvisorylineid' />
                                        <order attribute='createdon' descending='true' />
                                        <order attribute='hil_mobilenumber' descending='false' />
                                        <filter type='and'>
                                          <condition attribute='hil_enquirystauts' operator='in'>
                                            <value>4</value>
                                            <value>3</value>
                                            <value>2</value>
                                          </condition>
                                          <condition attribute='hil_syncwithsfa' operator='eq' value='0' />
                                        </filter>
                                      </entity>
                                    </fetch>";
                EntityCollection Found = service.RetrieveMultiple(new FetchExpression(_query));
                Console.WriteLine("Total Recored Fetch " + Found.Entities.Count);
                int i = 0;
                //Entity postImage = service.Retrieve("hil_homeadvisoryline", new Guid("f39f3b50-90c3-eb11-bacc-000d3af0c2b5"), new ColumnSet(true));
                foreach (Entity postImage1 in Found.Entities)
                {

                    i++;
                    Console.WriteLine("Count " + i);
                    Entity postImage = service.Retrieve("hil_homeadvisoryline", postImage1.Id, new ColumnSet(true));
                    EnquirySendSFA _enq = new EnquirySendSFA();
                    EnquiryLine _enqLine = new EnquiryLine();
                    EntityReference _EnqueryRef = postImage.GetAttributeValue<EntityReference>("hil_advisoryenquery");
                    Entity _enqueryEntity = service.Retrieve(_EnqueryRef.LogicalName, _EnqueryRef.Id, new ColumnSet(true));

                    _enqLine.EnquiryId = postImage.Contains("hil_name") ? postImage.GetAttributeValue<String>("hil_name") : "";
                    _enqLine.ParentEnquiryId = postImage.Contains("hil_advisoryenquery") ? postImage.GetAttributeValue<EntityReference>("hil_advisoryenquery").Name : "";
                    _enqLine.EnquiryLineGuid = postImage.Id.ToString();
                    if (postImage.Contains("hil_typeofenquiiry"))
                    {
                        Console.WriteLine("hil_typeofenquiiry ");
                        Entity _enqueryTypeEntity = service.Retrieve(postImage.GetAttributeValue<EntityReference>("hil_typeofenquiiry").LogicalName, postImage.GetAttributeValue<EntityReference>("hil_typeofenquiiry").Id, new ColumnSet("hil_enquirytypecode"));
                        _enqLine.EnquiryTypeCode = _enqueryTypeEntity.Contains("hil_enquirytypecode") ? _enqueryTypeEntity.GetAttributeValue<int>("hil_enquirytypecode").ToString() : "";
                    }
                    if (postImage.Contains("hil_typeofproduct"))
                    {
                        Console.WriteLine("hil_typeofproduct ");
                        Entity _ProductTypeEntity = service.Retrieve(postImage.GetAttributeValue<EntityReference>("hil_typeofproduct").LogicalName, postImage.GetAttributeValue<EntityReference>("hil_typeofproduct").Id, new ColumnSet("hil_code"));
                        _enqLine.ProducTypeCode = _ProductTypeEntity.Contains("hil_code") ? _ProductTypeEntity.GetAttributeValue<String>("hil_code") : "";
                    }
                    if (postImage.Contains("hil_assignedadvisor"))
                    {
                        Console.WriteLine("hil_assignedadvisor ");
                        Entity _AdvisorEntity = service.Retrieve(postImage.GetAttributeValue<EntityReference>("hil_assignedadvisor").LogicalName, postImage.GetAttributeValue<EntityReference>("hil_assignedadvisor").Id, new ColumnSet("hil_code"));
                        _enqLine.AdvisorCode = _AdvisorEntity.Contains("hil_code") ? _AdvisorEntity.GetAttributeValue<String>("hil_code") : "";
                    }
                    _enqLine.EnquiryStatus = postImage.Contains("hil_enquirystauts") ? postImage.GetAttributeValue<OptionSetValue>("hil_enquirystauts").Value.ToString() : "";
                    _enqLine.AppointmentId = postImage.Contains("hil_appointmentid") ? postImage.GetAttributeValue<String>("hil_appointmentid") : "";
                    _enqLine.AppointmentType = postImage.Contains("hil_appointmenttypes") ? postImage.GetAttributeValue<OptionSetValue>("hil_appointmenttypes").Value.ToString() : "";
                    if (postImage.Contains("hil_appointmentdate"))
                    {
                        Console.WriteLine("hil_appointmentdate ");
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

                    //String lineCreateon = postImage.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString();
                    DateTime _Createon = postImage.GetAttributeValue<DateTime>("createdon").AddMinutes(330);

                    Console.WriteLine("created on");

                    //String[] dateSplit = lineCreateon.Split(' ');
                    //String[] date = dateSplit[0].Split('-');

                    _enqLine.LineCreatedDate = _Createon.Year.ToString() + "-" + _Createon.Month.ToString().PadLeft(2,'0') + "-" + _Createon.Day.ToString().PadLeft(2, '0') + " " + _Createon.Hour.ToString().PadLeft(2, '0') + ":" + _Createon.Minute.ToString().PadLeft(2, '0') + ":" + _Createon.Second.ToString().PadLeft(2, '0');

                    //_enqLine.LineCreatedDate = lineCreateon;

                    _enqLine.PinCode = postImage.Contains("hil_pincode") ? postImage.GetAttributeValue<EntityReference>("hil_pincode").Name.ToString() : "";
                    List<EnquiryLine> EnqLines = new List<EnquiryLine>();
                    EnqLines.Add(_enqLine);

                    //String HeadCreateon = _enqueryEntity.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString();
                    DateTime _headCreateon = _enqueryEntity.GetAttributeValue<DateTime>("createdon").AddMinutes(330);

                    //Console.WriteLine("HeadCreate on");
                    //String[] dateSplit1 = lineCreateon.Split(' ');
                    //String[] date1 = dateSplit[0].Split('-');

                    _enq.EnquiryCreatedDate = _headCreateon.Year.ToString() + "-" + _headCreateon.Month.ToString().PadLeft(2, '0') + "-" + _headCreateon.Day.ToString().PadLeft(2, '0') + " " + _headCreateon.Hour.ToString().PadLeft(2, '0') + ":" + _headCreateon.Minute.ToString().PadLeft(2, '0') + ":" + _headCreateon.Second.ToString().PadLeft(2, '0');
                    //_enq.EnquiryCreatedDate = HeadCreateon;

                    _enq.EnquiryId = _enqueryEntity.Contains("hil_name") ? _enqueryEntity.GetAttributeValue<String>("hil_name") : "";
                    _enq.Area = _enqueryEntity.Contains("hil_areasqrt") ? _enqueryEntity.GetAttributeValue<String>("hil_areasqrt") : "";
                    if (_enqueryEntity.Contains("hil_typeofcustomer"))
                    {
                        Console.WriteLine("hil_typeofcustomer");
                        Entity _CustomerTypeEntity = service.Retrieve(_enqueryEntity.GetAttributeValue<EntityReference>("hil_typeofcustomer").LogicalName, _enqueryEntity.GetAttributeValue<EntityReference>("hil_typeofcustomer").Id, new ColumnSet("hil_code"));
                        _enq.CustomerTypeCode = _CustomerTypeEntity.Contains("hil_code") ? _CustomerTypeEntity.GetAttributeValue<String>("hil_code") : "";
                    }
                    if (_enqueryEntity.Contains("hil_propertytype"))
                    {
                        Console.WriteLine("hil_propertytype");
                        Entity _propertyTypeEntity = service.Retrieve(_enqueryEntity.GetAttributeValue<EntityReference>("hil_propertytype").LogicalName, _enqueryEntity.GetAttributeValue<EntityReference>("hil_propertytype").Id, new ColumnSet("hil_code"));
                        _enq.PropertyTypeCode = _propertyTypeEntity.Contains("hil_code") ? _propertyTypeEntity.GetAttributeValue<String>("hil_code") : "";
                    }
                    if (_enqueryEntity.Contains("hil_constructiontype"))
                    {
                        Console.WriteLine("hil_constructiontype");
                        Entity _constructionTypeEntity = service.Retrieve(_enqueryEntity.GetAttributeValue<EntityReference>("hil_constructiontype").LogicalName, _enqueryEntity.GetAttributeValue<EntityReference>("hil_constructiontype").Id, new ColumnSet("hil_code"));
                        _enq.ConstructionTypeCode = _constructionTypeEntity.Contains("hil_code") ? _constructionTypeEntity.GetAttributeValue<String>("hil_code") : "";
                    }
                    _enq.Rooftop = _enqueryEntity.GetAttributeValue<bool>("hil_rooftop") ? "true" : "false";
                    _enq.AssetType = _enqueryEntity.Contains("hil_assettype") ? _enqueryEntity.GetAttributeValue<String>("hil_assettype") : "";
                    _enq.CustomerName = _enqueryEntity.Contains("hil_customer") ? _enqueryEntity.GetAttributeValue<EntityReference>("hil_customer").Name : "";
                    _enq.MobileNumber = _enqueryEntity.Contains("hil_mobilenumber") ? _enqueryEntity.GetAttributeValue<String>("hil_mobilenumber") : "";


                    Console.WriteLine("MobileNumber");
                    _enqLine.CustomerRemarks = _enqueryEntity.Contains("hil_customerremarks") ? _enqueryEntity.GetAttributeValue<String>("hil_customerremarks") : "";

                    _enq.EmailId = _enqueryEntity.Contains("hil_emailid") ? _enqueryEntity.GetAttributeValue<String>("hil_emailid") : "";
                    if (_enqueryEntity.Contains("hil_city"))
                    {
                        Console.WriteLine("hil_city");
                        Entity _cityEntity = service.Retrieve(_enqueryEntity.GetAttributeValue<EntityReference>("hil_city").LogicalName, _enqueryEntity.GetAttributeValue<EntityReference>("hil_city").Id, new ColumnSet("hil_citycode"));
                        _enq.CityCode = _cityEntity.Contains("hil_citycode") ? _cityEntity.GetAttributeValue<String>("hil_citycode") : "";
                    }
                    if (_enqueryEntity.Contains("hil_state"))
                    {
                        Console.WriteLine("hil_state");
                        Entity _stateEntity = service.Retrieve(_enqueryEntity.GetAttributeValue<EntityReference>("hil_state").LogicalName, _enqueryEntity.GetAttributeValue<EntityReference>("hil_state").Id, new ColumnSet("hil_statecode"));
                        _enq.StateCode = _stateEntity.Contains("hil_statecode") ? _stateEntity.GetAttributeValue<String>("hil_statecode") : "";
                    }
                    _enq.EnquiryGuid = _enqueryEntity.Id.ToString();
                    _enq.PINCode = _enqueryEntity.Contains("hil_pincode") ? _enqueryEntity.GetAttributeValue<EntityReference>("hil_pincode").Name : "";
                    _enq.TDS = _enqueryEntity.Contains("hil_tds") ? _enqueryEntity.GetAttributeValue<String>("hil_tds") : "";
                    _enq.EnquiryLines = EnqLines;

                    _enq.SourceOfCreation = _enqueryEntity.Contains("hil_sourceofcreation") ? _enqueryEntity.GetAttributeValue<OptionSetValue>("hil_sourceofcreation").Value.ToString() : "";

                    EnqueryRequest enqReq = new EnqueryRequest();

                    Console.WriteLine("Data retrived");
                    enqReq.Data = _enq;
                    try
                    {


                        string data = JsonConvert.SerializeObject(enqReq);
                        Integration integration = common.IntegrationConfiguration(service, "SaveEnquiryDetailsToSFA");
                        string _authInfo = integration.Auth;
                        _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));
                        String sUrl = integration.uri;
                        var client = new RestClient(sUrl);
                        client.Timeout = -1;
                        var request = new RestRequest(Method.POST);
                        request.AddHeader("Authorization", "Basic " + _authInfo);
                        request.AddHeader("Content-Type", "application/json");
                        request.AddParameter("application/json", data, ParameterType.RequestBody);
                        IRestResponse response = client.Execute(request);

                        Console.WriteLine("Response :\n" + response.Content);
                        EnqueryResponseReq result = JsonConvert.DeserializeObject<EnqueryResponseReq>(response.Content);
                        Console.WriteLine("\nResult " + result.IsSuccess);
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
                    catch (Exception ex)
                    {
                        Console.WriteLine("Ex. Error : " + ex.Message);
                        continue;
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error " + ex.Message);
            }
            Console.WriteLine("Advisory pushed Successfully....");
        }
    }


}
