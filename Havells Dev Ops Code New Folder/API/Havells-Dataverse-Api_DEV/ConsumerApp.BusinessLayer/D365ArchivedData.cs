using System;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;
using RestSharp;
using Microsoft.Xrm.Sdk;
using System.Text.RegularExpressions;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using System.IO;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.Linq;
using System.Runtime.Serialization;
using Microsoft.Crm.Sdk.Messages;
using System.Text;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Web.Script.Serialization;

namespace ConsumerApp.BusinessLayer
{
    public class D365ArchivedData
    {
        public ArchivedJobsResponse GetJobData(ArchivedJobRequest inpReq)
        {
            IOrganizationService service = ConnectToCRM.GetOrgServiceProd();
            ArchivedJobsResponse response = new ArchivedJobsResponse();
            string eurl = string.Empty;
            if (inpReq.JobNumber != null && inpReq.JobNumber != string.Empty)
            {
                eurl = "RefId=" + inpReq.JobNumber + "&entityNo=4&ApiType=1";
            }
            else if (inpReq.MobileNumber != null && inpReq.MobileNumber != string.Empty)
            {
                String custRef = string.Empty;
                QueryExpression _query = new QueryExpression("contact");
                _query.ColumnSet = new ColumnSet(false);
                _query.Criteria = new FilterExpression(LogicalOperator.And);
                _query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, inpReq.MobileNumber);
                EntityCollection Found = service.RetrieveMultiple(_query);
                if (Found.Entities.Count == 1)
                {
                    eurl = "RefId=" + Found[0].Id.ToString() + "&entityNo=4&ApiType=3";
                }
                else
                {
                    response.success = false;
                    response.error = "Mobile Number Not Found";
                    return response;
                }
                //if (inpReq.pageNumber != null && inpReq.pageNumber != string.Empty)
                //{
                //    eurl = eurl + "&pageNumber=" + inpReq.pageNumber;
                //}
                //else
                //    eurl = eurl + "&pageNumber=1";
                //if (inpReq.recordsPerPage != null && inpReq.recordsPerPage != string.Empty)
                //{
                //    eurl = eurl + "&recordsPerPage=" + inpReq.recordsPerPage;
                //}
                //else
                //    eurl = eurl + "&recordsPerPage=4";
            }
            else
            {
                response.success = false;
                response.error = "Mobile Number or Job Id is Mandatory";
                return response;
            }
            Integration inconfig = IntegrationConfiguration(service, "ArchivedJobAPI");
            String sUrl = inconfig.uri + eurl;
            String Auth = inconfig.Auth;
            string _authInfo = "Basic " + Convert.ToBase64String(Encoding.ASCII.GetBytes(Auth));

            var client = new RestClient(sUrl);
            client.Timeout = -1;
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", _authInfo);
            IRestResponse responseGet = client.Execute(request);
            Console.WriteLine(responseGet.Content);

            response = JsonConvert.DeserializeObject<ArchivedJobsResponse>(responseGet.Content);
            return response;
        }
        public ResponseData GetArchiveData(RequestData ReqData)
        {
            IOrganizationService service = ConnectToCRM.GetOrgServiceProd();
            ResponseData _objActivity = null;
            Entity ent = new Entity();
            string username = "";
            string password = "";
            string baseUrl = "";

            try
            {
                if (service != null)
                {
                    string[] colset = new string[] { "hil_name", "hil_password", "hil_url", "hil_username" };
                    ent = service.Retrieve("hil_integrationconfiguration", new Guid("961bccd4-72ce-eb11-bacc-6045bd72ec68"), new ColumnSet(colset));
                    if (!ent.Contains("hil_name"))
                    {
                        _objActivity = new ResponseData { StatusCode = "404", StatusDescription = "Not found." };
                    }
                    else
                    {
                        if (ent.Attributes.Contains("hil_username"))
                        {
                            username = ent.GetAttributeValue<string>("hil_username");
                        }
                        if (ent.Attributes.Contains("hil_password"))
                        {
                            password = ent.GetAttributeValue<string>("hil_password");
                        }
                        if (ent.Attributes.Contains("hil_url"))
                        {
                            baseUrl = ent.GetAttributeValue<string>("hil_url");
                        }

                        string reqdata = new JavaScriptSerializer().Serialize(ReqData);
                        HttpClient client = new HttpClient();
                        string apiUrl = baseUrl + "RefId=" + ReqData.RefId + "&entityNo=" + ReqData.entityNo + "&ApiType=" + ReqData.ApiType + "";
                        client.DefaultRequestHeaders.Accept.Clear();
                        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                        var byteArray = Encoding.ASCII.GetBytes(username + ":" + password);
                        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                        HttpResponseMessage response = client.GetAsync(apiUrl).Result;

                        if (response.IsSuccessStatusCode)
                        {
                            //ActivityDetailsResponse Activity = (new JavaScriptSerializer()).Deserialize<ActivityDetailsResponse>(response.Content.ReadAsStringAsync().Result);
                            string Activity = response.Content.ReadAsStringAsync().Result;
                            _objActivity = new ResponseData { ResponseDetails = Activity, StatusCode = "200", StatusDescription = "Success" };
                        }
                        else
                        {
                            _objActivity = new ResponseData { StatusCode = response.StatusCode.ToString(), StatusDescription = response.ReasonPhrase };
                        }
                        return _objActivity;
                    }
                }
                else
                {
                    _objActivity = new ResponseData { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                _objActivity = new ResponseData { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message };
            }
            return _objActivity;
        }
        public Integration IntegrationConfiguration(IOrganizationService service, string Param)
        {
            Integration output = new Integration();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
            return output;
        }
    }

    [DataContract]
    public class ResponseData
    {
        [DataMember]
        public string ResponseDetails { get; set; }

        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
    }

    [DataContract]
    public class RequestData
    {
        [DataMember]
        public string RefId { get; set; }
        [DataMember]
        public string entityNo { get; set; }
        [DataMember]
        public string ApiType { get; set; }
    }

    [DataContract]
    public class ArchivedJobRequest
    {
        [DataMember]
        public String MobileNumber { get; set; }
        [DataMember]
        public String pageNumber { get; set; }
        [DataMember]
        public String recordsPerPage { get; set; }
        [DataMember]
        public String JobNumber { get; set; }
    }
    [DataContract]
    public class ArchivedJobsResponse
    {
        [DataMember]
        public List<dtoWorkOrders> workOrders { get; set; }
        [DataMember]
        public bool success { get; set; }
        [DataMember]
        public object error { get; set; }
    }
    [DataContract]
    public class dtoWorkOrders
    {
        [DataMember]
        public string createdby { get; set; }
        [DataMember]
        public string createdbyname { get; set; }
        [DataMember]
        public string createdbyyominame { get; set; }
        [DataMember]
        public string createdon { get; set; }
        [DataMember]
        public string createdonbehalfby { get; set; }
        [DataMember]
        public string createdonbehalfbyname { get; set; }
        [DataMember]
        public string createdonbehalfbyyominame { get; set; }
        [DataMember]
        public string entityimage { get; set; }
        [DataMember]
        public string entityimage_timestamp { get; set; }
        [DataMember]
        public string entityimage_url { get; set; }
        [DataMember]
        public string entityimageid { get; set; }
        [DataMember]
        public double? exchangerate { get; set; }
        [DataMember]
        public string hil_05 { get; set; }
        [DataMember]
        public double? hil_actualcharges { get; set; }
        [DataMember]
        public double? hil_actualcharges_base { get; set; }
        [DataMember]
        public string hil_address { get; set; }
        [DataMember]
        public string hil_addressdetails { get; set; }
        [DataMember]
        public string hil_addressname { get; set; }
        [DataMember]
        public string hil_agebucket { get; set; }
        [DataMember]
        public string hil_agebucketname { get; set; }
        [DataMember]
        public int? hil_ageing { get; set; }
        [DataMember]
        public string hil_allocatebykpi { get; set; }
        [DataMember]
        public string hil_allocatebykpiname { get; set; }
        [DataMember]
        public string hil_alternate { get; set; }
        [DataMember]
        public int? hil_appointmentcount { get; set; }
        [DataMember]
        public String hil_appointmentcount_date { get; set; }
        [DataMember]
        public int? hil_appointmentcount_state { get; set; }
        [DataMember]
        public bool? hil_appointmentmandatory { get; set; }
        [DataMember]
        public string hil_appointmentmandatoryname { get; set; }
        [DataMember]
        public String hil_appointmentseton { get; set; }
        [DataMember]
        public int? hil_appointmentstatus { get; set; }
        [DataMember]
        public string hil_appointmentstatusname { get; set; }
        [DataMember]
        public String hil_appointmenttime { get; set; }
        [DataMember]
        public string hil_area { get; set; }
        [DataMember]
        public string hil_areaname { get; set; }
        [DataMember]
        public string hil_areatext { get; set; }
        [DataMember]
        public string hil_assignedon { get; set; }
        [DataMember]
        public string hil_assignmentdate { get; set; }
        [DataMember]
        public string hil_assignmentmatrix { get; set; }
        [DataMember]
        public string hil_assignmentmatrixname { get; set; }
        [DataMember]
        public string hil_assigntobranchhead { get; set; }
        [DataMember]
        public string hil_assigntobranchheadname { get; set; }
        [DataMember]
        public bool? hil_assigntome { get; set; }
        [DataMember]
        public string hil_assigntomename { get; set; }
        [DataMember]
        public string hil_associateddealer { get; set; }
        [DataMember]
        public string hil_associateddealername { get; set; }
        [DataMember]
        public string hil_associateddealeryominame { get; set; }
        [DataMember]
        public int? hil_automaticassign { get; set; }
        [DataMember]
        public string hil_automaticassignname { get; set; }
        [DataMember]
        public string hil_branch { get; set; }
        [DataMember]
        public string hil_branchengineercity { get; set; }
        [DataMember]
        public string hil_branchheadapproval { get; set; }
        [DataMember]
        public string hil_branchheadapprovalname { get; set; }
        [DataMember]
        public string hil_branchname { get; set; }
        [DataMember]
        public string hil_branchtext { get; set; }
        [DataMember]
        public int? hil_brand { get; set; }
        [DataMember]
        public string hil_brandname { get; set; }
        [DataMember]
        public string hil_bsh { get; set; }
        [DataMember]
        public string hil_bshname { get; set; }
        [DataMember]
        public string hil_bshyominame { get; set; }
        [DataMember]
        public string hil_bucket { get; set; }
        [DataMember]
        public int? hil_bucket_ageing { get; set; }
        [DataMember]
        public bool? hil_calculatecharges { get; set; }
        [DataMember] public string hil_calculatechargesname { get; set; }
        [DataMember] public int? hil_callertype { get; set; }
        [DataMember] public string hil_callertypename { get; set; }
        [DataMember] public string hil_callingnumber { get; set; }
        [DataMember] public string hil_callsubtype { get; set; }
        [DataMember] public string hil_callsubtypename { get; set; }
        [DataMember] public string hil_calltype { get; set; }
        [DataMember] public string hil_calltypename { get; set; }
        [DataMember] public string hil_cancellationdate { get; set; }
        [DataMember] public string hil_characterstics { get; set; }
        [DataMember] public string hil_charactersticsname { get; set; }
        [DataMember] public string hil_chequedate { get; set; }
        [DataMember] public string hil_chequenumber { get; set; }
        [DataMember] public string hil_city { get; set; }
        [DataMember] public string hil_cityname { get; set; }
        [DataMember] public string hil_citytext { get; set; }
        [DataMember] public string hil_claimheader { get; set; }
        [DataMember] public string hil_claimheadername { get; set; }
        [DataMember] public string hil_claimline { get; set; }
        [DataMember] public string hil_claimlinename { get; set; }
        [DataMember] public bool? hil_closeticket { get; set; }
        [DataMember] public string hil_closeticketname { get; set; }
        [DataMember] public string hil_closureremarks { get; set; }
        [DataMember] public int? hil_countryclassification { get; set; }
        [DataMember] public string hil_countryclassificationname { get; set; }
        [DataMember] public string hil_customercomplaintdescription { get; set; }
        [DataMember] public int? hil_customerfeedback { get; set; }
        [DataMember] public string hil_customerfeedbackname { get; set; }
        [DataMember] public string hil_customername { get; set; }
        [DataMember] public string hil_customerref { get; set; }
        [DataMember] public string hil_customerrefidtype { get; set; }
        [DataMember] public string hil_customerrefname { get; set; }
        [DataMember] public string hil_customerrefyominame { get; set; }
        [DataMember] public bool? hil_delinkpo { get; set; }
        [DataMember] public string hil_delinkponame { get; set; }
        [DataMember] public string hil_district { get; set; }
        [DataMember] public string hil_districtname { get; set; }
        [DataMember] public string hil_districttext { get; set; }
        [DataMember] public string hil_email { get; set; }
        [DataMember] public string hil_emailcustomer { get; set; }
        [DataMember] public string hil_emailcustomername { get; set; }
        [DataMember] public string hil_emailcustomeryominame { get; set; }
        [DataMember] public string hil_emergencycall { get; set; }
        [DataMember] public string hil_emergencycallname { get; set; }
        [DataMember] public string hil_emergencyremarks { get; set; }
        [DataMember] public string hil_emolpyeenamecode { get; set; }
        [DataMember] public string hil_escalationcall { get; set; }
        [DataMember] public string hil_escalationcallname { get; set; }
        [DataMember] public int? hil_escalationcount { get; set; }
        [DataMember] public String hil_escalationcount_date { get; set; }
        [DataMember] public int? hil_escalationcount_state { get; set; }
        [DataMember] public string hil_escallationcountinteger { get; set; }
        [DataMember] public bool? hil_estimatecharges { get; set; }
        [DataMember] public string hil_estimatechargesname { get; set; }
        [DataMember] public string hil_estimatechargestotal { get; set; }
        [DataMember] public string hil_estimatechargestotal_base { get; set; }
        [DataMember] public string hil_estimatedchargedecimal { get; set; }
        [DataMember] public string hil_firstresponsein { get; set; }
        [DataMember] public string hil_firstresponseinname { get; set; }
        [DataMember] public String hil_firstresponseon { get; set; }
        [DataMember] public bool? hil_flagpo { get; set; }
        [DataMember] public string hil_flagponame { get; set; }
        [DataMember] public string hil_fulladdress { get; set; }
        [DataMember] public string hil_ifamceligible { get; set; }
        [DataMember] public string hil_ifamceligiblename { get; set; }
        [DataMember] public bool? hil_ifclosedjob { get; set; }
        [DataMember] public string hil_ifclosedjobname { get; set; }
        [DataMember] public bool? hil_ifparametersadded { get; set; }
        [DataMember] public string hil_ifparametersaddedname { get; set; }
        [DataMember] public int? hil_incidentquantity { get; set; }
        [DataMember] public bool? hil_ischargable { get; set; }
        [DataMember] public string hil_ischargablename { get; set; }
        [DataMember] public string hil_ischargeable { get; set; }
        [DataMember] public string hil_ischargeablename { get; set; }
        [DataMember] public string hil_isclaimable { get; set; }
        [DataMember] public string hil_isclaimablename { get; set; }
        [DataMember] public bool? hil_isocr { get; set; }
        [DataMember] public string hil_isocrname { get; set; }
        [DataMember] public string hil_iswrongjobclosure { get; set; }
        [DataMember] public string hil_iswrongjobclosurename { get; set; }
        [DataMember] public string hil_jobcancelreason { get; set; }
        [DataMember] public string hil_jobcancelreasonname { get; set; }
        [DataMember] public string hil_jobclass { get; set; }
        [DataMember] public string hil_jobclassapproval { get; set; }
        [DataMember] public string hil_jobclassapprovalname { get; set; }
        [DataMember] public string hil_jobclassname { get; set; }
        [DataMember] public bool? hil_jobclosemobile { get; set; }
        [DataMember] public string hil_jobclosemobilename { get; set; }
        [DataMember] public string hil_jobclosureby { get; set; }
        [DataMember] public string hil_jobclosurebyname { get; set; }
        [DataMember] public string hil_jobclosurebyyominame { get; set; }
        [DataMember] public String hil_jobclosuredon { get; set; }
        [DataMember] public String hil_jobclosureon { get; set; }
        [DataMember] public int? hil_jobincidentcount { get; set; }
        [DataMember] public int? hil_jobpriority { get; set; }
        [DataMember] public string hil_jobstatuscode { get; set; }
        [DataMember] public string hil_kkgcode { get; set; }
        [DataMember] public string hil_kkgcode_sms { get; set; }
        [DataMember] public string hil_kkgcode_smsname { get; set; }
        [DataMember] public string hil_kkgotp { get; set; }
        [DataMember] public bool? hil_kkgprovided { get; set; }
        [DataMember] public string hil_kkgprovidedname { get; set; }
        [DataMember] public String hil_lastresponsetime { get; set; }
        [DataMember] public string hil_locationofasset { get; set; }
        [DataMember] public string hil_locationofassetname { get; set; }
        [DataMember] public double? hil_maxquantity { get; set; }
        [DataMember] public double? hil_minquantity { get; set; }
        [DataMember] public string hil_mobilenumber { get; set; }
        [DataMember] public string hil_modelid { get; set; }
        [DataMember] public string hil_modelname { get; set; }
        [DataMember] public string hil_modeofpayment { get; set; }
        [DataMember] public string hil_modeofpaymentname { get; set; }
        [DataMember] public string hil_natureofcomplaint { get; set; }
        [DataMember] public string hil_natureofcomplaintname { get; set; }
        [DataMember] public string hil_newserialnumber { get; set; }
        [DataMember] public string hil_nsh { get; set; }
        [DataMember] public string hil_nshname { get; set; }
        [DataMember] public string hil_nshyominame { get; set; }
        [DataMember] public string hil_observation { get; set; }
        [DataMember] public string hil_observationname { get; set; }
        [DataMember] public string hil_onbehalfofdealer { get; set; }
        [DataMember] public string hil_onbehalfofdealername { get; set; }
        [DataMember] public string hil_onbehalfofdealeryominame { get; set; }
        [DataMember] public string hil_owneraccount { get; set; }
        [DataMember] public string hil_owneraccountname { get; set; }
        [DataMember] public string hil_owneraccountyominame { get; set; }
        [DataMember] public double? hil_payblechargedecimal { get; set; }
        [DataMember] public string hil_pendingquantity { get; set; }
        [DataMember] public int? hil_phonecallcount { get; set; }
        [DataMember] public string hil_pincode { get; set; }
        [DataMember] public string hil_pincodename { get; set; }
        [DataMember] public string hil_pincodetext { get; set; }
        [DataMember] public string hil_pmscount { get; set; }
        [DataMember] public string hil_pmsdate { get; set; }
        [DataMember] public string hil_preferreddate { get; set; }
        [DataMember] public string hil_preferredtime { get; set; }
        [DataMember] public string hil_preferredtimename { get; set; }
        [DataMember] public string hil_preferredtimeofvisit { get; set; }
        [DataMember] public string hil_productcategory { get; set; }
        [DataMember] public string hil_productcategoryname { get; set; }
        [DataMember] public string hil_productcatsubcatmapping { get; set; }
        [DataMember] public string hil_productcatsubcatmappingname { get; set; }
        [DataMember] public string hil_productreplacement { get; set; }
        [DataMember] public string hil_productreplacementname { get; set; }
        [DataMember] public string hil_productsubcategory { get; set; }
        [DataMember] public string hil_productsubcategorycallsubtype { get; set; }
        [DataMember] public string hil_productsubcategoryname { get; set; }
        [DataMember] public String hil_purchasedate { get; set; }
        [DataMember] public string hil_purchasedfrom { get; set; }
        [DataMember] public int? hil_quantity { get; set; }
        [DataMember] public string hil_quantityofunits { get; set; }
        [DataMember] public string hil_receiptamount { get; set; }
        [DataMember] public string hil_receiptamount_base { get; set; }
        [DataMember] public string hil_receiptnumber { get; set; }
        [DataMember] public string hil_regardingemail { get; set; }
        [DataMember] public string hil_regardingemailname { get; set; }
        [DataMember] public string hil_regardingfallback { get; set; }
        [DataMember] public string hil_regardingfallbackname { get; set; }
        [DataMember] public string hil_region { get; set; }
        [DataMember] public string hil_regionbranch { get; set; }
        [DataMember] public string hil_regionname { get; set; }
        [DataMember] public string hil_regiontext { get; set; }
        [DataMember] public int? hil_remaindercount { get; set; }
        [DataMember] public String hil_remaindercount_date { get; set; }
        [DataMember] public int? hil_remaindercount_state { get; set; }
        [DataMember] public string hil_remaindercountinteger { get; set; }
        [DataMember] public string hil_remindercall { get; set; }
        [DataMember] public string hil_remindercallname { get; set; }
        [DataMember] public string hil_repairdone { get; set; }
        [DataMember] public string hil_repairdonename { get; set; }
        [DataMember] public string hil_reportbinary { get; set; }
        [DataMember] public string hil_reporttext { get; set; }
        [DataMember] public int? hil_requesttype { get; set; }
        [DataMember] public string hil_requesttypename { get; set; }
        [DataMember] public bool? hil_resendkkg { get; set; }
        [DataMember] public string hil_resendkkgname { get; set; }
        [DataMember] public string hil_resolvebykpi { get; set; }
        [DataMember] public string hil_resolvebykpiname { get; set; }
        [DataMember] public string hil_returned { get; set; }
        [DataMember] public string hil_returnedname { get; set; }
        [DataMember] public string hil_salesoffice { get; set; }
        [DataMember] public string hil_salesofficename { get; set; }
        [DataMember] public int? hil_salutation { get; set; }
        [DataMember] public string hil_salutationname { get; set; }
        [DataMember] public bool? hil_sendestimate { get; set; }
        [DataMember] public string hil_sendestimatename { get; set; }
        [DataMember] public bool? hil_sendtcr { get; set; }
        [DataMember] public string hil_sendtcrname { get; set; }
        [DataMember] public string hil_serviceaddress { get; set; }
        [DataMember] public string hil_serviceaddressname { get; set; }
        [DataMember] public string hil_slastarton { get; set; }
        [DataMember] public int? hil_slastatus { get; set; }
        [DataMember] public string hil_slastatusname { get; set; }
        [DataMember] public int? hil_sourceofjob { get; set; }
        [DataMember] public string hil_sourceofjobname { get; set; }
        [DataMember] public string hil_state { get; set; }
        [DataMember] public string hil_statename { get; set; }
        [DataMember] public string hil_statetext { get; set; }
        [DataMember] public string hil_systemremarks { get; set; }
        [DataMember] public string hil_technicianname { get; set; }
        [DataMember] public double? hil_timeinappointmentdate { get; set; }
        [DataMember] public double? hil_timeincurrentdate { get; set; }
        [DataMember] public double? hil_timeinfirstresponse { get; set; }
        [DataMember] public double? hil_timeinjobclosure { get; set; }
        [DataMember] public string hil_town { get; set; }
        [DataMember] public string hil_townname { get; set; }
        [DataMember] public string hil_typeofassignee { get; set; }
        [DataMember] public string hil_typeofassigneename { get; set; }
        [DataMember] public int? hil_warrantystatus { get; set; }
        [DataMember] public string hil_warrantystatusname { get; set; }
        [DataMember] public string hil_webclosureremarks { get; set; }
        [DataMember] public string hil_workdoneby { get; set; }
        [DataMember] public string hil_workdonebyname { get; set; }
        [DataMember] public string hil_workdonebyyominame { get; set; }
        [DataMember] public bool? hil_workstarted { get; set; }
        [DataMember] public string hil_workstartedname { get; set; }
        [DataMember] public string importsequencenumber { get; set; }
        [DataMember] public string lastonholdtime { get; set; }
        [DataMember] public string modifiedby { get; set; }
        [DataMember] public string modifiedbyname { get; set; }
        [DataMember] public string modifiedbyyominame { get; set; }
        [DataMember] public String modifiedon { get; set; }
        [DataMember] public string modifiedonbehalfby { get; set; }
        [DataMember] public string modifiedonbehalfbyname { get; set; }
        [DataMember] public string modifiedonbehalfbyyominame { get; set; }
        [DataMember] public string msdyn_address1 { get; set; }
        [DataMember] public string msdyn_address2 { get; set; }
        [DataMember] public string msdyn_address3 { get; set; }
        [DataMember] public string msdyn_addressname { get; set; }
        [DataMember] public string msdyn_agreement { get; set; }
        [DataMember] public string msdyn_agreementname { get; set; }
        [DataMember] public string msdyn_billingaccount { get; set; }
        [DataMember] public string msdyn_billingaccountname { get; set; }
        [DataMember] public string msdyn_billingaccountyominame { get; set; }
        [DataMember] public string msdyn_bookingsummary { get; set; }
        [DataMember] public string msdyn_childindex { get; set; }
        [DataMember] public string msdyn_city { get; set; }
        [DataMember] public string msdyn_closedby { get; set; }
        [DataMember] public string msdyn_closedbyname { get; set; }
        [DataMember] public string msdyn_closedbyyominame { get; set; }
        [DataMember] public string msdyn_country { get; set; }
        [DataMember] public string msdyn_customerasset { get; set; }
        [DataMember] public string msdyn_customerassetname { get; set; }
        [DataMember] public string msdyn_datewindowend { get; set; }
        [DataMember] public string msdyn_datewindowstart { get; set; }
        [DataMember] public double? msdyn_estimatesubtotalamount { get; set; }
        [DataMember] public double? msdyn_estimatesubtotalamount_base { get; set; }
        [DataMember] public string msdyn_followupnote { get; set; }
        [DataMember] public bool? msdyn_followuprequired { get; set; }
        [DataMember] public string msdyn_followuprequiredname { get; set; }
        [DataMember] public string msdyn_instructions { get; set; }
        [DataMember] public string msdyn_internalflags { get; set; }
        [DataMember] public bool? msdyn_isfollowup { get; set; }
        [DataMember] public string msdyn_isfollowupname { get; set; }
        [DataMember] public bool? msdyn_ismobile { get; set; }
        [DataMember] public string msdyn_ismobilename { get; set; }
        [DataMember] public string msdyn_latitude { get; set; }
        [DataMember] public string msdyn_longitude { get; set; }
        [DataMember] public string msdyn_name { get; set; }
        [DataMember] public string msdyn_opportunityid { get; set; }
        [DataMember] public string msdyn_opportunityidname { get; set; }
        [DataMember] public string msdyn_parentworkorder { get; set; }
        [DataMember] public string msdyn_parentworkordername { get; set; }
        [DataMember] public string msdyn_postalcode { get; set; }
        [DataMember] public string msdyn_preferredresource { get; set; }
        [DataMember] public string msdyn_preferredresourcename { get; set; }
        [DataMember] public string msdyn_pricelist { get; set; }
        [DataMember] public string msdyn_pricelistname { get; set; }
        [DataMember] public string msdyn_primaryincidentdescription { get; set; }
        [DataMember] public string msdyn_primaryincidentestimatedduration { get; set; }
        [DataMember] public string msdyn_primaryincidenttype { get; set; }
        [DataMember] public string msdyn_primaryincidenttypename { get; set; }
        [DataMember] public string msdyn_priority { get; set; }
        [DataMember] public string msdyn_priorityname { get; set; }
        [DataMember] public string msdyn_reportedbycontact { get; set; }
        [DataMember] public string msdyn_reportedbycontactname { get; set; }
        [DataMember] public string msdyn_reportedbycontactyominame { get; set; }
        [DataMember] public string msdyn_serviceaccount { get; set; }
        [DataMember] public string msdyn_serviceaccountname { get; set; }
        [DataMember] public string msdyn_serviceaccountyominame { get; set; }
        [DataMember] public string msdyn_servicerequest { get; set; }
        [DataMember] public string msdyn_servicerequestname { get; set; }
        [DataMember] public string msdyn_serviceterritory { get; set; }
        [DataMember] public string msdyn_serviceterritoryname { get; set; }
        [DataMember] public string msdyn_stateorprovince { get; set; }
        [DataMember] public string msdyn_substatus { get; set; }
        [DataMember] public string msdyn_substatusname { get; set; }
        [DataMember] public double? msdyn_subtotalamount { get; set; }
        [DataMember] public double? msdyn_subtotalamount_base { get; set; }
        [DataMember] public string msdyn_supportcontact { get; set; }
        [DataMember] public string msdyn_supportcontactname { get; set; }
        [DataMember] public int? msdyn_systemstatus { get; set; }
        [DataMember] public string msdyn_systemstatusname { get; set; }
        [DataMember] public bool? msdyn_taxable { get; set; }
        [DataMember] public string msdyn_taxablename { get; set; }
        [DataMember] public string msdyn_taxcode { get; set; }
        [DataMember] public string msdyn_taxcodename { get; set; }
        [DataMember] public String msdyn_timeclosed { get; set; }
        [DataMember] public String msdyn_timefrompromised { get; set; }
        [DataMember] public string msdyn_timegroup { get; set; }
        [DataMember] public string msdyn_timegroupdetailselected { get; set; }
        [DataMember] public string msdyn_timegroupdetailselectedname { get; set; }
        [DataMember] public string msdyn_timegroupname { get; set; }
        [DataMember] public string msdyn_timetopromised { get; set; }
        [DataMember] public string msdyn_timewindowend { get; set; }
        [DataMember] public string msdyn_timewindowstart { get; set; }
        [DataMember] public double? msdyn_totalamount { get; set; }
        [DataMember] public double? msdyn_totalamount_base { get; set; }
        [DataMember] public double? msdyn_totalsalestax { get; set; }
        [DataMember] public double? msdyn_totalsalestax_base { get; set; }
        [DataMember] public string msdyn_worklocation { get; set; }
        [DataMember] public string msdyn_worklocationname { get; set; }
        [DataMember] public string msdyn_workorderid { get; set; }
        [DataMember] public string msdyn_workordersummary { get; set; }
        [DataMember] public string msdyn_workordertype { get; set; }
        [DataMember] public string msdyn_workordertypename { get; set; }
        [DataMember] public string new_owner_temp { get; set; }
        [DataMember] public string new_owner_tempname { get; set; }
        [DataMember] public string new_owner_tempyominame { get; set; }
        [DataMember] public string onholdtime { get; set; }
        [DataMember] public string overriddencreatedon { get; set; }
        [DataMember] public string ownerid { get; set; }
        [DataMember] public string owneridname { get; set; }
        [DataMember] public string owneridtype { get; set; }
        [DataMember] public string owneridyominame { get; set; }
        [DataMember] public string owningbusinessunit { get; set; }
        [DataMember] public string owningteam { get; set; }
        [DataMember] public string owninguser { get; set; }
        [DataMember] public string processid { get; set; }
        [DataMember] public string slaid { get; set; }
        [DataMember] public string slaidname { get; set; }
        [DataMember] public string slainvokedid { get; set; }
        [DataMember] public string slainvokedidname { get; set; }
        [DataMember] public string stageid { get; set; }
        [DataMember] public int? statecode { get; set; }
        [DataMember] public string statecodename { get; set; }
        [DataMember] public int? statuscode { get; set; }
        [DataMember] public string statuscodename { get; set; }
        [DataMember] public int? timezoneruleversionnumber { get; set; }
        [DataMember] public string transactioncurrencyid { get; set; }
        [DataMember] public string transactioncurrencyidname { get; set; }
        [DataMember] public string traversedpath { get; set; }
        [DataMember] public string utcconversiontimezonecode { get; set; }
        [DataMember] public long? versionnumber { get; set; }
        [DataMember] public String CreatedDateWH { get; set; }
    }
}
