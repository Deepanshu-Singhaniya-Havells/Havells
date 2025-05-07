using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Globalization;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.NPS

{
    public class LitmusFeedback : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion
            try
            {
                string jobId = Convert.ToString(context.InputParameters["JobId"]);
                string category = Convert.ToString(context.InputParameters["Category"]);
                int technicianScore = 0;
                string _technicianScore = Convert.ToString(context.InputParameters["TechnicianScore"]);

                if (int.TryParse(_technicianScore, out technicianScore))
                {
                    technicianScore = Convert.ToInt32(_technicianScore);
                }

                int score = 0;
                string _Score = Convert.ToString(context.InputParameters["Score"]);

                if (int.TryParse(_Score, out score))
                {
                    score = Convert.ToInt32(_Score);
                }

                string closureTag = Convert.ToString(context.InputParameters["ClosureTag"]);

                if (string.IsNullOrWhiteSpace(jobId))
                {
                    string msg = "Please enter jobid.";
                    var jobIdResponse = JsonSerializer.Serialize(new LitmusFeedBackResponse { Status = false, Message = msg });
                    context.OutputParameters["data"] = jobIdResponse;
                    return;
                }
                if (!APValidate.NumericValue(jobId))
                {
                    var jobIdResponse = JsonSerializer.Serialize(new LitmusFeedBackResponse { Status = false, Message = "Invalid Job Id format." });
                    context.OutputParameters["data"] = jobIdResponse;
                    return;
                }
                if (string.IsNullOrWhiteSpace(category))
                {
                    string msg = "Please enter category.";
                    var categoryResponse = JsonSerializer.Serialize(new LitmusFeedBackResponse { Status = false, Message = msg });
                    context.OutputParameters["data"] = categoryResponse;
                    return;
                }
                if (!string.IsNullOrWhiteSpace(category))
                {
                    if (!APValidate.IsValidString(category))
                    {
                        string msg = "Invalid Category";
                        var categoryValidResponse = JsonSerializer.Serialize(new LitmusFeedBackResponse { Status = false, Message = msg });
                        context.OutputParameters["data"] = categoryValidResponse;
                        return;
                    }
                    if (category.Length > 100)
                    {
                        string msg = "Category should not be more than 100 characters long.";
                        var categoryValidResponse = JsonSerializer.Serialize(new LitmusFeedBackResponse { Status = false, Message = msg });
                        context.OutputParameters["data"] = categoryValidResponse;
                        return;
                    }
                }
                if (decimal.TryParse(_Score, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal ScoreValue))
                {
                    if (ScoreValue < 0 || ScoreValue > 10)
                    {
                        string msg = "Please enter a valid score number between 0 to 10.";
                        var TechnicianScoreValidResponse = JsonSerializer.Serialize(new LitmusFeedBackResponse { Status = false, Message = msg });
                        context.OutputParameters["data"] = TechnicianScoreValidResponse;
                        return;
                    }

                }
                if (decimal.TryParse(_technicianScore, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal numericValue))
                {
                    if (numericValue < 0 || numericValue > 5)
                    {
                        string msg = "Please enter a valid Technician Score number between 0 to 5.";
                        var TechnicianScoreValidResponse = JsonSerializer.Serialize(new LitmusFeedBackResponse { Status = false, Message = msg });
                        context.OutputParameters["data"] = TechnicianScoreValidResponse;
                        return;
                    }

                }
                if (!string.IsNullOrWhiteSpace(closureTag))
                {
                    if (closureTag.Length > 999)
                    {
                        string msg = "Closure Tag should not be more than 999 characters long.";
                        var TechnicianScoreValidResponse = JsonSerializer.Serialize(new LitmusFeedBackResponse { Status = false, Message = msg });
                        context.OutputParameters["data"] = TechnicianScoreValidResponse;
                        return;
                    }
                }
                LitmusCustomerFeedBack litmusCustomerFeedBack = new LitmusCustomerFeedBack()
                {
                    JobId = jobId,
                    Category = category,
                    ClosureTag = closureTag,
                    Score = int.TryParse(_Score, out score) ? score.ToString() : null,
                    TechnicianScore = int.TryParse(_technicianScore, out technicianScore) ? technicianScore.ToString() : null
                };

                LitmusFeedBackResponse litmusFeedBackResponse = UpdateCustFeedBack(service, litmusCustomerFeedBack);
                var serializedResponse = JsonSerializer.Serialize(litmusFeedBackResponse);
                context.OutputParameters["data"] = serializedResponse;
            }
            catch (JsonException ex)
            {
                var TechnicianScoreValidResponse = JsonSerializer.Serialize(new LitmusFeedBackResponse { Status = false, Message = ex.Message });
                context.OutputParameters["data"] = TechnicianScoreValidResponse;
                return;
            }
        }
        public LitmusFeedBackResponse UpdateCustFeedBack(IOrganizationService service, LitmusCustomerFeedBack ParamCustFeedBack)
        {
            LitmusFeedBackResponse obj_Result = new LitmusFeedBackResponse();
            try
            {
                string jobsextensionFetchXml = string.Empty;
                if (service != null)
                {
                    jobsextensionFetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name = 'hil_jobsextension' >
                        <attribute name = 'hil_jobsextensionid' />
                        <attribute name = 'hil_name' />
                        <filter type = 'and' >
                        <condition attribute='hil_name' operator='eq' value='{ParamCustFeedBack.JobId}'/>
                        </filter>
                        </entity>
                        </fetch> ";
                    EntityCollection entColExt = service.RetrieveMultiple(new FetchExpression(jobsextensionFetchXml));
                    if (entColExt.Entities.Count > 0)
                    {
                        Entity _jobExt = new Entity("hil_jobsextension", entColExt.Entities[0].Id);
                        _jobExt["hil_category"] = ParamCustFeedBack.Category;
                        if (ParamCustFeedBack.Score != null)
                            _jobExt["hil_score"] = Convert.ToInt32(ParamCustFeedBack.Score);
                        else
                            _jobExt["hil_score"] = null;
                        if (ParamCustFeedBack.TechnicianScore != null)
                            _jobExt["hil_technicianrating"] = Convert.ToInt32(ParamCustFeedBack.TechnicianScore);
                        else
                            _jobExt["hil_technicianrating"] = null;

                        _jobExt["hil_closuretag"] = ParamCustFeedBack.ClosureTag;
                        DateTime _today = DateTime.Now.AddMinutes(330);
                        _jobExt["hil_npsmodifiedon"] = _today;

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
    public class LitmusCustomerFeedBack
    {
        public string JobId { get; set; }
        public string Score { get; set; }
        public string Category { get; set; }
        public string TechnicianScore { get; set; }
        public string ClosureTag { get; set; }
    }
    public class LitmusFeedBackResponse
    {
        public bool Status { get; set; }
        public string Message { get; set; }
    }
}