using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.Actions
{
    public class ValidateSessions : IPlugin

    {

        private ITracingService tracingService;

        private IOrganizationService service;

        public void Execute(IServiceProvider serviceProvider)
        {

            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion

            try
            {
                if (context.InputParameters.Contains("AccessToken") && context.InputParameters["AccessToken"] is string
                    && context.InputParameters.Contains("MobileNumber") && context.InputParameters["MobileNumber"] is string
                    && context.Depth == 1)
                {

                    string mobileNumber = (string)context.InputParameters["MobileNumber"];
                    string accessToken = (string)context.InputParameters["AccessToken"];


                    if (string.IsNullOrEmpty(mobileNumber))
                    {
                        context.OutputParameters["Status"] = "False";
                        context.OutputParameters["Message"] = "Mobile Number is required";
                        return;
                    }

                    if (string.IsNullOrEmpty(accessToken))
                    {
                        context.OutputParameters["Status"] = "False";
                        context.OutputParameters["Message"] = "Access Token is required";
                        return;
                    }


                    Request testReq = new Request();
                    testReq.AccessToken = accessToken;
                    testReq.MobileNumber = mobileNumber;

                    Response response = ValidateToken(testReq);

                    context.OutputParameters["Status"] = response.Status;
                    context.OutputParameters["Message"] = response.Message;


                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["Status"] = "Error !";
                context.OutputParameters["Message"] = "D365 Internal Error " + ex.Message;
            }

        }

        private Response ValidateToken(Request testReq)
        {

            //Request testReq = new Request();
            //testReq.AccessToken = "Temp Data";
            //testReq.MobileNumber = "123245"; 

            Response response = new Response();

            QueryExpression sessions = new QueryExpression("hil_customconnectorauthsession");
            sessions.ColumnSet = new ColumnSet(true);
            sessions.Criteria.AddCondition("hil_accesstoken", ConditionOperator.Equal, testReq.AccessToken);

            EntityCollection sessionsColl = service.RetrieveMultiple(sessions);
            if (sessionsColl.Entities.Count == 1)
            {
                string mobileNumber = sessionsColl.Entities[0].GetAttributeValue<string>("hil_mobilenumber");

                if (mobileNumber == testReq.MobileNumber)
                {
                    response.Status = "True";
                    response.Message = "Token is valid and the mobile number is correct";


                }
                else
                {
                    response.Status = "False";
                    response.Message = "Token is valid but the mobile number is incorrect";

                }

            }
            else if (sessionsColl.Entities.Count > 1)
            {
                response.Status = "False";
                response.Message = "There are multiple records that exist with the same sessions";
            }
            else
            {

                Entity newSession = new Entity("hil_customconnectorauthsession");
                newSession["hil_mobilenumber"] = testReq.MobileNumber;
                newSession["hil_accesstoken"] = testReq.AccessToken;

                service.Create(newSession);
                response.Status = "True";
                response.Message = "Created new session";
            }

            return response;

        }

        class Request
        {
            public string AccessToken { get; set; }
            public string MobileNumber { get; set; }
        }

        class Response
        {
            public string Status { get; set; }

            public string Message { get; set; }
        }
    }
}
