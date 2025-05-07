using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.Text.Json.Nodes;
using Newtonsoft.Json;

namespace ConsumerApp.BusinessLayer
{
    public class LitmusFeedback
    {
        public LtimusFeedBackResponse UpdateCustFeedBack(LtimusCustomerFeedBack ParamCustFeedBack)
        {
            LtimusFeedBackResponse obj_Result = new LtimusFeedBackResponse();
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    string xml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                             <entity name = 'hil_jobsextension' >
                             <attribute name = 'hil_jobsextensionid' />
                             <attribute name = 'hil_name' />
                             <filter type = 'and' >
                               <condition attribute='hil_name' operator='eq' value='{ParamCustFeedBack.JobId}'/>
                             </filter>
                             </entity>
                             </fetch> ";
                    EntityCollection entColExt = service.RetrieveMultiple(new FetchExpression(xml));
                    if (entColExt.Entities.Count > 0)
                    {
                        Entity _jobExt = new Entity("hil_jobsextension", entColExt.Entities[0].Id);
                        _jobExt["hil_category"] = ParamCustFeedBack.Category;
                        _jobExt["hil_score"] = ParamCustFeedBack.Score;
						_jobExt["hil_technicianrating"] = ParamCustFeedBack.TechnicianScore;
						_jobExt["hil_closuretag"] = ParamCustFeedBack.ClosureTag;

						service.Update(_jobExt);
                        obj_Result.Status = true;
                        obj_Result.Message = "Success";
                    }
                    else
                    {
                        obj_Result.Status = false;
                        obj_Result.Message = "Job Id does not exist.";
                    }
                }
                else
                {
                    obj_Result.Status = false;
                    obj_Result.Message = "D365 Service Unavailable.";
                }
            }
            catch (Exception ex)
            {
                obj_Result.Status = false;
                obj_Result.Message = "D365 Internal Server Error : " + ex.Message;
            }
            return obj_Result;
        }

    }
    [DataContract]
    public class LtimusCustomerFeedBack
    {
        [DataMember]
        public string JobId { get; set; }
        [DataMember]
        public int Score { get; set; }
        [DataMember]
        public string Category { get; set; }
		[DataMember]
		public int TechnicianScore { get; set; }
		[DataMember]
		public string ClosureTag { get; set; }
    }
    [DataContract]
    public class LtimusFeedBackResponse
    {
        [DataMember]
        public Boolean Status { get; set; }
        [DataMember]
        public string Message { get; set; }

    }
}

