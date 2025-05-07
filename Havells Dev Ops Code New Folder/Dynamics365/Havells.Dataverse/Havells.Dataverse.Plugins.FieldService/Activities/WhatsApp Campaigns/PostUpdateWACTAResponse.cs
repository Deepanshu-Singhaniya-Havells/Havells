using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace Havells.Dataverse.Plugins.FieldService.Activities.Notes
{
    public class PostUpdateWACTAResponse : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && (context.MessageName.ToUpper() == "UPDATE") && (context.PrimaryEntityName.ToLower() == "hil_wacta"))
            {
                Entity wacta = (Entity)context.InputParameters["Target"];
                try
                {
                    if (wacta.Contains("hil_whatsappctaresponse"))
                    {
                        try
                        {
                            OptionSetValue _whatsappCTAResponse = wacta.Contains("hil_whatsappctaresponse") ? wacta.GetAttributeValue<OptionSetValue>("hil_whatsappctaresponse") : null;
                            Entity _entWhatsappCTA = service.Retrieve(wacta.LogicalName, wacta.Id, new ColumnSet("hil_jobid", "hil_watriggername", "hil_appointmentdate", "hil_appointmenttime"));
                            DateTime? _appointmentdate = _entWhatsappCTA.Contains("hil_appointmentdate") ? _entWhatsappCTA.GetAttributeValue<DateTime?>("hil_appointmentdate") : null;
                            OptionSetValue _appointmenttime = _entWhatsappCTA.Contains("hil_appointmenttime") ? _entWhatsappCTA.GetAttributeValue<OptionSetValue>("hil_appointmenttime") : null;
                            string _whatsappCampaign = _entWhatsappCTA.Contains("hil_watriggername") ? _entWhatsappCTA.GetAttributeValue<string>("hil_watriggername") : null;
                            if (_whatsappCTAResponse != null && !string.IsNullOrWhiteSpace(_whatsappCampaign))
                                ChangeWorkOrderSubstatus(_entWhatsappCTA.GetAttributeValue<EntityReference>("hil_jobid"), _whatsappCTAResponse, service, _whatsappCampaign, _appointmentdate, _appointmenttime);

                        }
                        catch (Exception ex)
                        {
                            throw new InvalidPluginExecutionException(ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);

                }
            }
        }
        private void ChangeWorkOrderSubstatus(EntityReference _entRefworkOrder, OptionSetValue _whatsappCTAResponse, IOrganizationService service, string _whatsappCampaign, DateTime? _appointmentdate, OptionSetValue _appointmenttime)
        {
            string _jobsubstatus = string.Empty;
            string _jobClosureRemarks = string.Empty;

            Entity entUpdate = new Entity(_entRefworkOrder.LogicalName, _entRefworkOrder.Id);

            if (_whatsappCampaign == "cc_b2c_kkg")
            {
                if (_whatsappCTAResponse.Value == 1) //YES
                {
                    _jobsubstatus = "Closed";
                    entUpdate["hil_closeticket"] = true;
                    entUpdate["hil_kkgcode_sms"] = new OptionSetValue(910590008); // Work Completed (Confirmation received from Customer through WhatsApp)
                }
                else if (_whatsappCTAResponse.Value == 2) //NO
                {
                    _jobsubstatus = string.Empty;
                    entUpdate["hil_createchildjob"] = true; // True
                    entUpdate["hil_kkgcode_sms"] = new OptionSetValue(100000001); // Work not done(Re-open- new complain no)
                }
                _jobClosureRemarks = "Job closed via Whatsapp over KKG Audit Campaign.";
            }
            else if (_whatsappCampaign == "cancel_2cta_h8")
            {
                if (_whatsappCTAResponse.Value == 3)
                {
                    _jobsubstatus = "Canceled";
                    _jobClosureRemarks = "Job Canceled via Whatsapp over Cancellation Campaign.";

                }
                else if (_whatsappCTAResponse.Value == 4)
                {
                    entUpdate["hil_reopenbyconsumer"] = true;
                }
            }
            else if (_whatsappCampaign == "cancel_3cta_jo")
            {
                if (_whatsappCTAResponse.Value == 3)
                {
                    _jobsubstatus = "Canceled";
                    _jobClosureRemarks = "Job Canceled via Whatsapp over Cancellation Campaign.";

                }
                else if (_whatsappCTAResponse.Value == 4)
                {
                    entUpdate["hil_reopenbyconsumer"] = true;
                }
                else if (_whatsappCTAResponse.Value == 5)
                {
                    entUpdate["hil_reopenbyconsumer"] = true;
                    if (_appointmentdate != null)
                        entUpdate["hil_preferreddate"] = Convert.ToDateTime(_appointmentdate);
                    if (_appointmenttime != null)
                        entUpdate["hil_preferredtime"] = _appointmenttime;
                }
            }
            else if (_whatsappCampaign == "invoice_2cta_68")
            {
                if (_whatsappCTAResponse.Value == 3) //consumer don't have invoice bill copy
                {
                    _jobsubstatus = "Canceled";
                    _jobClosureRemarks = "Job Canceled via Whatsapp over Cancellation Campaign.";
                }
                else if (_whatsappCTAResponse.Value == 6) //User Uploaded Invoice Copy
                {
                    entUpdate["hil_reopenbyconsumer"] = true;
                }
            }
            else if (_whatsappCampaign == "est_na_2cta")
            {
                if (_whatsappCTAResponse.Value == 1)
                {
                    entUpdate["hil_reopenbyconsumer"] = true;
                }
                else if (_whatsappCTAResponse.Value == 2)
                {
                    _jobsubstatus = "Canceled";
                    _jobClosureRemarks = "Job Canceled via Whatsapp over Cancellation Campaign.";
                }
            }
            else
            {
                throw new Exception($"Invalid WhatsApp Campaign");
            }

            EntityCollection entCol = null;
            if (!string.IsNullOrWhiteSpace(_jobsubstatus))
            {
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='msdyn_workordersubstatus'>
                <attribute name='msdyn_systemstatus' />
                <attribute name='msdyn_workordersubstatusid' />
                <filter type='and'>
                    <condition attribute='msdyn_name' operator='eq' value='{_jobsubstatus}' />
                </filter>
                </entity>
                </fetch>";
                entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
            }
            if (entCol != null)
            {
                if (entCol.Entities.Count > 0)
                {
                    entUpdate["msdyn_substatus"] = entCol.Entities[0].ToEntityReference();
                    entUpdate["msdyn_systemstatus"] = entCol.Entities[0].GetAttributeValue<OptionSetValue>("msdyn_systemstatus");
                }
            }

            if (!string.IsNullOrWhiteSpace(_jobClosureRemarks))
            {
                entUpdate["hil_webclosureremarks"] = _jobClosureRemarks;
                entUpdate["hil_closureremarks"] = _jobClosureRemarks;
            }
            try
            {
                service.Update(entUpdate);
            }
            catch (Exception ex)
            {
                throw new Exception($"{ex.Message} {ex.StackTrace}");
            }
        }
    }
}
