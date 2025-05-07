using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Net;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.EasyRewardLoyalty
{
    public class EasyRewardGetUserInfo : IPlugin
    {
        private static readonly string[] SourceTypeAr = new string[] { "1", "3", "5", "6", "7", "9", "12" };
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            if (context == null)
            {
                var sourceTypeResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = "Invalid Application Context." });
                context.OutputParameters["data"] = sourceTypeResponse;
                return;
            }

            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            string LoginUserId = Convert.ToString(context.InputParameters["LoginUserId"]);
            string UserToken = Convert.ToString(context.InputParameters["UserToken"]);
            string jsonString = Convert.ToString(context.InputParameters["reqdata"]);
            var data = JsonSerializer.Deserialize<AMCOrdersParam>(jsonString);
            string SourceType = data.SourceType;


            if (!APValidate.IsValidMobileNumber(LoginUserId))
            {
                string msg = string.IsNullOrWhiteSpace(LoginUserId) ? "Mobile number is required." : "Invalid mobile number.";
                var loginResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 400, Message = msg });
                context.OutputParameters["data"] = loginResponse;
                return;
            }
            if (!SourceTypeAr.Contains(SourceType))
            {
                string msg = string.IsNullOrWhiteSpace(SourceType) ? "Source type is required." : "Invalid Source type.";
                var sourceTypeResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = msg });
                context.OutputParameters["data"] = sourceTypeResponse;
                return;
            }
            UserinfoDetails details = new UserinfoDetails();
            try
            {
                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                       <entity name='contact'>
                         <attribute name='fullname' />
                         <attribute name='contactid' />
                         <attribute name='hil_preferredlanguageforcommunication' />
                         <attribute name='hil_isloyaltyprogramenabled' />
                         <attribute name='lastname' />
                         <attribute name='hil_gender' />
                         <attribute name='firstname' />
                         <attribute name='emailaddress1' />
                         <attribute name='hil_dateofbirth' />
                         <order attribute='contactid' descending='false' />
                         <filter type='and'>
                           <condition attribute='mobilephone' operator='eq' value='{LoginUserId}' />
                         </filter>
                       </entity>
                     </fetch>";
                EntityCollection Info = service.RetrieveMultiple(new FetchExpression(fetchXml));

                if (Info.Entities.Count > 0)
                {
                    string UserPreferredLanguage = "";
                    Guid preferredlanguageforcommunicationId = Info.Entities[0].Contains("hil_preferredlanguageforcommunication") ? Info.Entities[0].GetAttributeValue<EntityReference>("hil_preferredlanguageforcommunication").Id : Guid.Empty;
                    if (preferredlanguageforcommunicationId != Guid.Empty)
                    {
                        UserPreferredLanguage = service.Retrieve("hil_preferredlanguageforcommunication", preferredlanguageforcommunicationId, new ColumnSet("hil_code")).GetAttributeValue<string>("hil_code");
                    }
                    //details = new UserinfoDetails()
                    //{
                    details.FirstName = Info.Entities[0].Contains("firstname") ? Info.Entities[0].GetAttributeValue<string>("firstname") : "";
                    details.LastName = Info.Entities[0].Contains("lastname") ? Info.Entities[0].GetAttributeValue<string>("lastname") : "";
                    details.Gender = Info.Entities[0].Contains("hil_gender") ? (Info.Entities[0].GetAttributeValue<bool>("hil_gender") ? "Female" : "Male") : "";
                    details.DOB = Info.Entities[0].Contains("hil_dateofbirth") ? Info.Entities[0].GetAttributeValue<DateTime>("hil_dateofbirth").ToShortDateString() : "";
                    details.EmailId = Info.Entities[0].Contains("emailaddress1") ? Info.Entities[0].GetAttributeValue<string>("emailaddress1") : "";
                    details.LoyaltyMember = Info.Entities[0].Contains("hil_isloyaltyprogramenabled") ? Info.Entities[0].GetAttributeValue<OptionSetValue>("hil_isloyaltyprogramenabled").Value.ToString() : null;
                    details.PreferredLanguage = UserPreferredLanguage;
                    details.Guid = Info.Entities[0].Id;
                    details.StatusCode = (int)HttpStatusCode.OK;

                    if (details.LoyaltyMember == "1")
                    {
                        details.LoyaltyMember = "Y";
                    }
                    else if (details.LoyaltyMember == "2")
                    {
                        details.LoyaltyMember = "B";
                    }
                    else
                    {
                        details.LoyaltyMember = "N";
                    }
                    //};

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

                    EntityCollection Info1 = service.RetrieveMultiple(new FetchExpression(fetchXml));

                    if (Info1.Entities.Count > 0)
                    {
                        details.Eligible = 'Y';
                    }

                    context.OutputParameters["data"] = JsonSerializer.Serialize(details);
                }

                else
                {
                    var sourceTypeResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = "Record Not Found." });
                    context.OutputParameters["data"] = sourceTypeResponse;
                    return;
                }
            }
            catch (Exception ex)
            {
                var sourceTypeResponse = JsonSerializer.Serialize(new RequestStatus { StatusCode = 204, Message = ex.Message });
                context.OutputParameters["data"] = sourceTypeResponse;
                return;
            }
        }
    }
    public class UserinfoDetails
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Gender { get; set; }
        public string EmailId { get; set; }
        public string DOB { get; set; }
        public string LoyaltyMember { get; set; }
        public char Eligible { get; set; } = 'N';
        public Guid Guid { get; set; }
        public string PreferredLanguage { get; set; }
        public int StatusCode { get; set; }

    }
    public class AMCOrdersParam
    {
        public string SourceType { get; set; }
    }
    public class RequestStatus
    {
        public int StatusCode { get; set; }
        public string Message { get; set; }
    }
}