using HavellsSync_Data.IManager;
using HavellsSync_ModelData.AMC;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.EasyReward;
using HavellsSync_ModelData.Epos;
using HavellsSync_ModelData.Product;
using Microsoft.Extensions.Configuration;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.ComponentModel.DataAnnotations;
using System.Net;
using System.Net.NetworkInformation;
using System.Text;
using System.Text.RegularExpressions;

namespace HavellsSync_Data.Manager
{
    public class EposManager : IEposManager
    {
        private IConfiguration configuration;
        private ICrmService _CrmService;
        public EposManager(ICrmService crmService, IConfiguration configuration)
        {
            Check.Argument.IsNotNull(nameof(crmService), crmService);
            _CrmService = crmService;
            this.configuration = configuration;
        }


        public async Task<EposUserinfoDetails> GetConsumerInfo(string MobileNumber)
        {
            EposUserinfoDetails details = null;
            QueryExpression query;
            query = new QueryExpression("contact");
            query.ColumnSet = new ColumnSet("fullname", "hil_gender", "hil_dateofbirth", "emailaddress1", "hil_consent");
            query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, MobileNumber);
            EntityCollection Info = _CrmService.RetrieveMultiple(query);
            if (Info.Entities.Count > 0)
            {
                details = new EposUserinfoDetails()
                {
                    ConsumerName = Info.Entities[0].Contains("fullname") ? Info.Entities[0].GetAttributeValue<string>("fullname") : "",
                    EmailId = Info.Entities[0].Contains("emailaddress1") ? Info.Entities[0].GetAttributeValue<string>("emailaddress1") : "",
                    Gender = Info.Entities[0].Contains("hil_gender") ? (Info.Entities[0].GetAttributeValue<bool>("hil_gender") ? "False" : "True") : "",
                    DOB = Info.Entities[0].Contains("hil_dateofbirth") ? Info.Entities[0].GetAttributeValue<DateTime>("hil_dateofbirth").AddMinutes(330).ToString("yyyy-MM-dd") : "",
                    Consent = Info.Entities[0].Contains("hil_consent") ? (Info.Entities[0].GetAttributeValue<bool>("hil_consent") ? "True" : "False") : "",
                    Status = true,
                    Response = "Success",

                };

                string fetchXml = $@"<fetch top='1' >
                    <entity name='msdyn_workorder'>
                        <attribute name='hil_address' />
                        <attribute name='hil_fulladdress' />
                        <attribute name='msdyn_name' />
                        <attribute name='hil_addressdetails' />
                        <attribute name='hil_serviceaddress' />
                        <attribute name='msdyn_addressname' />
                        <attribute name='msdyn_displayaddress' />
                        <filter>
                            <condition attribute='hil_customerref' operator='eq' value='{Info.Entities[0].Id}' uitype='contact' />
                        </filter>
                        <order attribute='createdon' descending='true' />
                        <link-entity name='hil_address' from='hil_addressid' to='hil_address' link-type='inner' alias='addrs' >
                            <attribute name='hil_street1' />
                            <attribute name='hil_street2' />
                            <attribute name='hil_businessgeo' />    
                            <link-entity name='hil_businessmapping' from='hil_businessmappingid' to='hil_businessgeo' link-type='inner' alias='mapping' >
                            <attribute name='hil_pincode' />                         
                        </link-entity>
                        </link-entity>                       
                    </entity>
                </fetch>";
                EntityCollection Info1 = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));
                if (Info1.Entities.Count > 0)
                {
                    details.AddressLine1 = Info1.Entities[0].Contains("addrs.hil_street1") ? Info1.Entities[0].GetAttributeValue<AliasedValue>("addrs.hil_street1").Value.ToString() : "";
                    details.AddressLine2 = Info1.Entities[0].Contains("addrs.hil_street2") ? Info1.Entities[0].GetAttributeValue<AliasedValue>("addrs.hil_street2").Value.ToString() : "";
                    details.PINCode = Info1.Entities[0].Contains("mapping.hil_pincode") ? ((EntityReference)Info1.Entities[0].GetAttributeValue<AliasedValue>("mapping.hil_pincode").Value).Name : null;
                }
                else
                {
                    string fetchXml1 = $@"<fetch top='1'>
                    <entity name='hil_address'>
                    <attribute name='hil_street2'/>
                    <attribute name='hil_street1'/>
                     <order attribute='createdon' descending='true' />
                    <filter type='and'>
                    <condition attribute='hil_customer' operator='eq' uitype='contact' value='{Info.Entities[0].Id}'/>
                    </filter>
                    <link-entity name='hil_businessmapping' from='hil_businessmappingid' to='hil_businessgeo' visible='false' link-type='outer' alias='mapping'>
                    <attribute name='hil_pincode'/>
                    </link-entity>
                    </entity>
                    </fetch>";
                    EntityCollection Info2 = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml1));
                    if (Info2.Entities.Count > 0)
                    {
                        details.AddressLine1 = Info2.Entities[0].Contains("hil_street1") ? Info2.Entities[0].Attributes["hil_street1"].ToString() : null;
                        details.AddressLine2 = Info2.Entities[0].Contains("hil_street2") ? Info2.Entities[0].Attributes["hil_street2"].ToString() : null;
                        details.PINCode = Info2.Entities[0].Contains("mapping.hil_pincode") ? ((EntityReference)Info2.Entities[0].GetAttributeValue<AliasedValue>("mapping.hil_pincode").Value).Name : null;
                    }
                }
            }
            return details;
        }

        public async Task<EposJobStatus> GetServiceCallStatus(string JobId)
        {
            EposJobStatus JobStatus = null;
            string fetchXml = $@"<fetch>
                        <entity name='msdyn_workorder'>
                        <attribute name='msdyn_name'/>
                        <attribute name='msdyn_substatus'/>
                        <order attribute='msdyn_name' descending='false'/>
                        <filter type='and'>
                        <condition attribute='msdyn_name' operator='eq' value='{JobId}'/>
                        </filter>
                        </entity>
                        </fetch>";
            EntityCollection Info = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));
            if (Info.Entities.Count > 0)
            {
                JobStatus = new EposJobStatus()
                {
                    JobStatus = Info.Entities[0].Contains("msdyn_substatus") ? Info.Entities[0].GetAttributeValue<EntityReference>("msdyn_substatus").Name : null,
                    Status = true,
                    Response = "Success",

                };
            }
            return JobStatus;
        }

        public async Task<ServiceCallRequestData> SyncEPOSSalesData(ServiceCallRequestData param)
        {
            ServiceCallRequestData Servicecall = new ServiceCallRequestData();
            StringBuilder ErrorMessage = new StringBuilder();
            Regex regexDate = new Regex(@"^\d{4}\-(0[1-9]|1[012])\-(0[1-9]|[12][0-9]|3[01])$");
            string[] PreferredTimeofService = new string[] { "1", "2", "3" };
            bool flag = true;
            bool Exceptionflag = true;
            try
            {
                Entity ServiceCallRequest = new Entity("hil_servicecallrequest");
                ServiceCallRequest["hil_customername"] = param.ConsumerName;
                ServiceCallRequest["hil_customeremailid"] = param.EmailId;
                ServiceCallRequest["hil_customermobilenumber"] = param.MobileNumber;
                ServiceCallRequest["hil_fulladdress"] = param.AddressLine1;
                ServiceCallRequest["hil_pincode"] = param.PINCode;
                ServiceCallRequest["hil_addressline2"] = param.AddressLine2;
                ServiceCallRequest["hil_consent"] = Convert.ToBoolean(param.Consent);
                ServiceCallRequest["hil_dateofbirth"] = param.DOB;
                ServiceCallRequest["hil_gender"] = Convert.ToBoolean(param.Gender);
                ServiceCallRequest["hil_landmark"] = param.Landmark;
                ServiceCallRequest["hil_source"] = new OptionSetValue(23);
                Servicecall.GuidID = _CrmService.Create(ServiceCallRequest);
                int i = 1;
                foreach (var item in param.ServiceCallLineItem)
                {
                    flag = true;
                    try
                    {
                        Entity ServiceCallRequestDetail = new Entity("hil_servicecallrequestdetail");
                        ServiceCallRequestDetail["hil_serialnumber"] = item.SerialNumber;
                        if (!string.IsNullOrEmpty(item.SKUCode))
                        {
                            if (CommonMethods.IsValidModelNumber(_CrmService, item.SKUCode))
                            {
                                ServiceCallRequestDetail["hil_name"] = item.SKUCode;
                            }
                            else
                            {
                                ErrorMessage.AppendLine(string.Format("Invalid SKU Code {0} for line item {1}.", item.SKUCode, i));
                                flag = false;
                            }
                        }
                        if (regexDate.IsMatch(item.PreferredDateofService))
                        {
                            ServiceCallRequestDetail["hil_servicedate"] = Convert.ToDateTime(item.PreferredDateofService);
                        }
                        else
                        {
                            ErrorMessage.AppendLine(string.Format("Preferred Date of Service {0} Format should be (yyyy-MM-dd) for line item {1}.", item.PreferredDateofService, i));
                            flag = false;
                        }
                        ServiceCallRequestDetail["hil_callsubtype"] = new EntityReference("hil_callsubtype", new Guid("E3129D79-3C0B-E911-A94E-000D3AF06CD4"));
                        if (PreferredTimeofService.Contains(item.PreferredTimeofService))
                        {
                            ServiceCallRequestDetail["hil_preferredtime"] = new OptionSetValue(Convert.ToInt32(item.PreferredTimeofService));
                        }
                        else
                        {
                            ErrorMessage.AppendLine(string.Format("Preferred Time of Service {0} should be 1|2|3 for line item {1}.", item.PreferredTimeofService, i));
                            flag = false;
                        }
                        ServiceCallRequestDetail["hil_servicecallrequest"] = new EntityReference("hil_servicecallrequest", Servicecall.GuidID);
                        if (flag)
                        {
                            _CrmService.Create(ServiceCallRequestDetail);
                            ErrorMessage.AppendLine(string.Format("Success line item {0}.", i));
                        }
                    }
                    catch (Exception ex)
                    {
                        Exceptionflag = false;
                        ErrorMessage.AppendLine(ex.Message);
                        continue;
                    }
                    i++;
                }
                return new ServiceCallRequestData { Response = (flag ? "Success" : ErrorMessage.ToString()), Status = Exceptionflag };
            }
            catch (Exception ex)
            {
                ErrorMessage.AppendLine(ex.Message);
            }
            finally
            {
                #region log
                if (_CrmService != null)
                {
                    Entity intigrationTrace = new Entity("hil_integrationtrace");
                    intigrationTrace["hil_entityname"] = "hil_servicecallrequest";
                    intigrationTrace["hil_entityid"] = Servicecall.GuidID.ToString();
                    intigrationTrace["hil_request"] = JsonConvert.SerializeObject(param);
                    intigrationTrace["hil_name"] = "ePOS";
                    intigrationTrace["hil_response"] = ErrorMessage.ToString();
                    _CrmService.Create(intigrationTrace);
                }
                #endregion log
            }
            return Servicecall;
        }
    }
}
