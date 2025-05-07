using HavellsSync_Data.IManager;
using HavellsSync_ModelData.AMC;
using HavellsSync_ModelData.Common;
using HavellsSync_ModelData.EasyReward;
using HavellsSync_ModelData.ICommon;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Net;
using System.Text.RegularExpressions;

namespace HavellsSync_Data.Manager
{
    public class EasyRewardManager : IEasyRewardManager
    {
        private IConfiguration configuration;
        private ICrmService _CrmService;
        private readonly ICustomLog _logger;
        public EasyRewardManager(ICrmService crmService, IConfiguration configuration, ICustomLog logger)
        {
            Check.Argument.IsNotNull(nameof(crmService), crmService);
            _CrmService = crmService;
            this.configuration = configuration;
            _logger = logger;
        }
        public async Task<UserinfoDetails> GetUserInfo(string MobileNumber)
        {
            _logger.LogToFile(new Exception(string.Format("GetUserInfo|Manager|2|{0}|Start GetUserInfo Manager", MobileNumber)));
            UserinfoDetails details = null;
            try
            {
                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='contact'>
                                    <attribute name='fullname' />
                                    <attribute name='contactid' />
                                    <attribute name='hil_preferredlanguageforcommunication' />
                                    <attribute name='hil_loyaltyprogramenabled' />
                                    <attribute name='lastname' />
                                    <attribute name='hil_gender' />
                                    <attribute name='firstname' />
                                    <attribute name='emailaddress1' />
                                    <attribute name='hil_dateofbirth' />
                                    <order attribute='contactid' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='mobilephone' operator='eq' value='{MobileNumber}' />
                                    </filter>
                                  </entity>
                                </fetch>";
                EntityCollection Info = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));
                _logger.LogToFile(new Exception(string.Format("GetUserInfo|Manager|3|{0}|Searching for Consumer Mobile", MobileNumber)));

                if (Info.Entities.Count > 0)
                {
                    _logger.LogToFile(new Exception(string.Format("GetUserInfo|Manager|4|{0}|Consumer Mobile found", MobileNumber)));
                    string UserPreferredLanguage = "";
                    Guid preferredlanguageforcommunicationId = Info.Entities[0].Contains("hil_preferredlanguageforcommunication") ? Info.Entities[0].GetAttributeValue<EntityReference>("hil_preferredlanguageforcommunication").Id : Guid.Empty;
                    if (preferredlanguageforcommunicationId != Guid.Empty)
                    {
                        UserPreferredLanguage = _CrmService.Retrieve("hil_preferredlanguageforcommunication", preferredlanguageforcommunicationId, new ColumnSet("hil_code")).GetAttributeValue<string>("hil_code");
                    }
                    details = new UserinfoDetails()
                    {
                        FirstName = Info.Entities[0].Contains("firstname") ? Info.Entities[0].GetAttributeValue<string>("firstname") : "",
                        LastName = Info.Entities[0].Contains("lastname") ? Info.Entities[0].GetAttributeValue<string>("lastname") : "",
                        Gender = Info.Entities[0].Contains("hil_gender") ? (Info.Entities[0].GetAttributeValue<bool>("hil_gender") ? "Female" : "Male") : "",
                        DOB = Info.Entities[0].Contains("hil_dateofbirth") ? Info.Entities[0].GetAttributeValue<DateTime>("hil_dateofbirth").ToShortDateString() : "",
                        EmailId = Info.Entities[0].Contains("emailaddress1") ? Info.Entities[0].GetAttributeValue<string>("emailaddress1") : "",
                        LoyaltyMember = Info.Entities[0].Contains("hil_loyaltyprogramenabled") ? (Info.Entities[0].GetAttributeValue<bool>("hil_loyaltyprogramenabled") ? 'Y' : 'N') : 'N',
                        PreferredLanguage = UserPreferredLanguage,
                        Guid = Info.Entities[0].Id,
                        StatusCode = (int)HttpStatusCode.OK
                    };
                    _logger.LogToFile(new Exception(string.Format("GetUserInfo|Manager|5|{0}|Fetch Starts Consumer Assets for LR", MobileNumber)));
                    fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' top='1'>
                        <entity name='msdyn_customerasset'>
                        <attribute name='msdyn_customerassetid' />
                        <filter type='and'>
                            <condition attribute='hil_customer' operator='eq' value='{Info.Entities[0].Id}' />
                        </filter>
                        <link-entity name='product' from='productid' to='hil_productcategory' link-type='inner' alias='ak'>
                            <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='al'>
                                <filter type='and'>
                                    <condition attribute='hil_eligibleforloyaltyprograms' operator='eq' value='1' />
                                </filter>
                            </link-entity>
                        </link-entity>
                        </entity>
                        </fetch>";

                    EntityCollection Info1 = _CrmService.RetrieveMultiple(new FetchExpression(fetchXml));
                    _logger.LogToFile(new Exception(string.Format("GetUserInfo|Manager|6|{0}|Fetch Ends Consumer Assets for LR", MobileNumber)));
                    if (Info1.Entities.Count > 0)
                    {
                        _logger.LogToFile(new Exception(string.Format("GetUserInfo|Manager|7|{0}|Eligible Consumer Assets found", MobileNumber)));
                        details.Eligible = 'Y';
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogToFile(new Exception(string.Format("GetUserInfo|Manager|8|{0}|Manager Exception: {1}", MobileNumber, JsonConvert.SerializeObject(ex))));
            }
            _logger.LogToFile(new Exception(string.Format("GetUserInfo|Manager|9|{0}|End GetUserInfo Manager", MobileNumber)));
            return details;
        }

        public async Task<EasyRewardResponse> UpdateLoyaltyStatus(string MobileNumber, string SourceType)
        {
            EasyRewardResponse ERResponse = new EasyRewardResponse();
            _logger.LogToFile(new Exception(string.Format("UpdateLoyaltyStatus|Manager|2|{0}|Start UpdateLoyaltyStatus Manager", MobileNumber)));
            try
            {
                QueryExpression query;
                query = new QueryExpression("contact");
                query.ColumnSet = new ColumnSet("firstname");
                query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, MobileNumber);
                EntityCollection Info = _CrmService.RetrieveMultiple(query);
                _logger.LogToFile(new Exception(string.Format("UpdateLoyaltyStatus|Manager|3|{0}|Retrieved Customer Details", MobileNumber)));

                if (Info.Entities.Count > 0)
                {
                    //var CustomerGuid = CommonMethods.getCustomerGuid(_CrmService, MobileNumber);
                    Entity ContactUpdate = new Entity("contact", Info.Entities[0].Id);

                    ContactUpdate["hil_loyaltyprogramtier"] = new OptionSetValue(1);
                    ContactUpdate["hil_loyaltyprogramenabled"] = true;
                    _CrmService.Update(ContactUpdate);
                    _logger.LogToFile(new Exception(string.Format("UpdateLoyaltyStatus|Manager|4|{0}|Update loyalty details for Customer", MobileNumber)));
                    ERResponse.Response = "Success";
                }
                else
                {
                    ERResponse.Response = null;
                }
            }
            catch (Exception ex)
            {
                _logger.LogToFile(new Exception(string.Format("UpdateLoyaltyStatus|Manager|5|{0}|Manager Exception: {1}", MobileNumber, JsonConvert.SerializeObject(ex))));
            }
            _logger.LogToFile(new Exception(string.Format("UpdateLoyaltyStatus|Manager|6|{0}|End UpdateLoyaltyStatus Manager", MobileNumber)));
            return ERResponse;
        }
    }
}
