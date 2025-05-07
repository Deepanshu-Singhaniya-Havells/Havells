using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;

namespace Havells.Dataverse.CustomConnector.EasyRewardLoyalty
{
    public class EasyRewardUpdateLoyaltyStatus : IPlugin
    {
        private static readonly string[] SourceTypeAr = new string[] { "1", "3", "5", "6", "7", "9", "12" };
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
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

            EasyRewardResponse response = UpdateLoyaltyStatus(service, LoginUserId, SourceType).Result;
            if (!string.IsNullOrWhiteSpace(response.Response))
            {
                var data1 = JsonSerializer.Serialize(new { StatusCode = 200, Message = response.Response });
                context.OutputParameters["data"] = data1;
            }
            else
            {
                var data1 = JsonSerializer.Serialize(new { StatusCode = 400, Message = "Customer does not exist." });
                context.OutputParameters["data"] = data1;
            }
        }
        public async Task<EasyRewardResponse> UpdateLoyaltyStatus(IOrganizationService _CrmService, string MobileNumber, string SourceType)
        {
            EasyRewardResponse ERResponse = new EasyRewardResponse();

            try
            {
                QueryExpression query;
                query = new QueryExpression("contact");
                query.ColumnSet = new ColumnSet("firstname");
                query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, MobileNumber);
                EntityCollection Info = _CrmService.RetrieveMultiple(query);

                if (Info.Entities.Count > 0)
                {
                    Entity ContactUpdate = new Entity("contact", Info.Entities[0].Id);
                    ContactUpdate["hil_loyaltyprogramtier"] = new OptionSetValue(1);
                    ContactUpdate["hil_loyaltyprogramenabled"] = true;
                    ContactUpdate["hil_isloyaltyprogramenabled"] = new OptionSetValue(1);
                    _CrmService.Update(ContactUpdate);
                    ERResponse.Response = "Success";
                }
                else
                {
                    ERResponse.Response = null;
                }
            }
            catch (Exception ex)
            {

            }
            return ERResponse;
        }
    }
    public class EasyRewardResponse
    {
        public string Response { get; set; }

    }
}
