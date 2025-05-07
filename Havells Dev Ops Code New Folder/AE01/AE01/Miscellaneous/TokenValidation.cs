
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace AE01.Miscellaneous
{
    internal class TokenValidation
    {
        private IOrganizationService service;

        public TokenValidation(IOrganizationService _service)
        {
            this.service = _service; 
        }
     
        public void Main()
        {
            Request testReq = new Request();
            testReq.AccessToken = "Temp Data";
            testReq.MobileNumber = "123245";
            Response response = ValidateToken(testReq);

            Console.WriteLine("Status: " + response.Status);
            Console.WriteLine("Message: " + response.Message); 
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
                    response.Status= true;
                    response.Message = "Token is valid and the mobile number is correct";

                }
                else
                {
                    response.Status = false;
                    response.Message = "Token is valid but the mobile number is incorrect";

                }

            }
            else if (sessionsColl.Entities.Count > 1)
            {
                response.Status = false;
                response.Message = "There are multiple records that exist with the same sessions";
            }
            else
            {   

                Entity newSession = new Entity("hil_customconnectorauthsession");
                newSession["hil_mobilenumber"] = testReq.MobileNumber;
                newSession["hil_accesstoken"] = testReq.AccessToken;

                service.Create(newSession);
                response.Status = true;
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
            public bool Status { get; set; }

            public string Message { get; set; }
        }
    }
}
