using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Text.Json;

namespace Havells.Dataverse.CustomConnector.WhatsApp
{

    public class UpdateWhatsAppCtaResponse : IPlugin
    {
        private IOrganizationService service;
        private ITracingService tracingService;
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
                GetWAResponse _retObj = new GetWAResponse();
                string AppointmentDate = string.Empty;
                string AppointmentTime = string.Empty;

                string JobID = Convert.ToString(context.InputParameters["JobId"]);
                string CtAResponse = Convert.ToString(context.InputParameters["CTAResponse"]);
                string CustomerGuid = Convert.ToString(context.InputParameters["CustomerGuid"]);
                string TriggerName = Convert.ToString(context.InputParameters["CampaignId"]);

                if (context.InputParameters["AppointmentDate"] != null)
                    AppointmentDate = Convert.ToString(context.InputParameters["AppointmentDate"]);
                if (context.InputParameters["AppointmentTime"] != null)
                    AppointmentTime = Convert.ToString(context.InputParameters["AppointmentTime"]);

                if (context.Depth == 1)
                {
                    GetWACTAData data = new GetWACTAData();
                    data.JobId = JobID;
                    data.CustomerGuid = CustomerGuid;
                    data.CTAResponse = CtAResponse;
                    data.CampaignId = TriggerName;
                    data.AppointmentDate = AppointmentDate;
                    data.AppointmentTime = AppointmentTime;
                    context.OutputParameters["data"] = JsonSerializer.Serialize(UpdateWAResponse(data));
                }
                else
                {
                    context.OutputParameters["data"] = JsonSerializer.Serialize(new GetWAResponse() { Status = false, StatusMessage = "Failed" });
                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["data"] = JsonSerializer.Serialize(new GetWAResponse() { Status = false, StatusMessage = "Failed to Update " + ex.Message });
            }
        }
        public GetWAResponse UpdateWAResponse(GetWACTAData _reqData)
        {
            GetWAResponse _retObj = new GetWAResponse();
            string _fetchXML = string.Empty;

            if (string.IsNullOrWhiteSpace(_reqData.JobId))
            {
                _retObj = new GetWAResponse() { Status = false, StatusMessage = "Job Id is mandatory" };
                return _retObj;
            }
            if (string.IsNullOrWhiteSpace(_reqData.CustomerGuid))
            {
                _retObj = new GetWAResponse() { Status = false, StatusMessage = "Customer Guid is mandatory" };
                return _retObj;
            }
            if (string.IsNullOrWhiteSpace(_reqData.CTAResponse))
            {
                _retObj = new GetWAResponse() { Status = false, StatusMessage = "CTA Response is mandatory" };
                return _retObj;
            }
            if (string.IsNullOrWhiteSpace(_reqData.CampaignId))
            {
                _retObj = new GetWAResponse() { Status = false, StatusMessage = "Campaign Id is mandatory" };
                return _retObj;
            }

            _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='msdyn_workorder'>
                <attribute name='msdyn_workorderid'/>
                <filter type='and'>
                    <condition attribute='msdyn_name' operator='eq' value='{_reqData.JobId}'/>
                </filter>
                </entity>
                </fetch>";

            EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (entCol.Entities.Count > 0)
            {
                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_wacta'>
                    <attribute name='hil_wactaid'/>
                    <attribute name='hil_watriggername'/>
                    <attribute name='createdon'/>
                    <attribute name='hil_whatsappctaresponse'/>
                    <order attribute='hil_watriggername' descending='false'/>
                    <filter type='and'>
                        <condition attribute='hil_jobid' operator='eq' value='{entCol.Entities[0].Id}'/>
                        <condition attribute='hil_watriggername' operator='eq' value='{_reqData.CampaignId}'/>
                    </filter>
                    </entity>
                    </fetch>";

                EntityCollection entColl = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (entColl.Entities.Count > 0)
                {
                    OptionSetValue _whatsappctaresponse = null;
                    if (entColl[0].Contains("hil_whatsappctaresponse"))
                    {
                        _whatsappctaresponse = entColl[0].GetAttributeValue<OptionSetValue>("hil_whatsappctaresponse");
                    }
                    if (_whatsappctaresponse != null)
                    {
                        _retObj = new GetWAResponse() { Status = false, StatusMessage = "Campaign Response is already submitted." };
                    }
                    else
                    {
                        Entity wcta = new Entity("hil_wacta", entColl[0].Id);
                        wcta["hil_whatsappctaresponse"] = new OptionSetValue(Convert.ToInt32(_reqData.CTAResponse));
                        wcta["hil_wastatusreason"] = new OptionSetValue(2);//Close the WhatsApp CTA Activity
                        if (_reqData.AppointmentDate != null && _reqData.AppointmentDate.Length > 0)
                            wcta["hil_appointmentdate"] = Convert.ToDateTime(_reqData.AppointmentDate);
                        if (_reqData.AppointmentTime != null && _reqData.AppointmentTime.Length > 0)
                            wcta["hil_appointmenttime"] = new OptionSetValue(Convert.ToInt32(_reqData.AppointmentTime));

                        try
                        {
                            service.Update(wcta);
                            _retObj = new GetWAResponse() { Status = true, StatusMessage = "Response updated successfully" };
                        }
                        catch (Exception ex)
                        {
                            _retObj = new GetWAResponse() { Status = false, StatusMessage = "ERROR:" + ex.Message };
                        }
                    }
                }
                else
                {
                    _retObj = new GetWAResponse() { Status = false, StatusMessage = "No WhatsApp CTA found." };
                }
                return _retObj;
            }
            else
            {
                _retObj = new GetWAResponse() { Status = false, StatusMessage = "Customer is not link with JOB" };
                return _retObj;

            }
        }
        public class GetWACTAData
        {
            public string JobId { get; set; }
            public string CTAResponse { get; set; }
            public string CustomerGuid { get; set; }
            public string CampaignId { get; set; }
            public string AppointmentDate { get; set; }
            public string AppointmentTime { get; set; }
        }

        public class GetWAResponse
        {
            public string StatusMessage { get; set; }
            public bool Status { get; set; }
        }

    }

}

