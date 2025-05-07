using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;

namespace Havells_Plugin.SMS
{
    public class SMSPostCreate : IPlugin
    {
        private static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                if (context.InputParameters.Contains("Target")
                         && context.InputParameters["Target"] is Entity
                         && context.PrimaryEntityName.ToLower() == "hil_smsconfiguration"
                         && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity entity1 = (Entity)context.InputParameters["Target"];
                    

                    Entity entity = service.Retrieve(entity1.LogicalName, entity1.Id, new ColumnSet(true));
                    hil_smsconfiguration SMS = entity.ToEntity<hil_smsconfiguration>();
                    Guid _regardingObjectId = Guid.Empty;

                    if (entity.Contains("hil_smstemplate"))
                    {
                        if (entity.Attributes.Contains("regardingobjectid"))
                        {
                            _regardingObjectId = entity.GetAttributeValue<EntityReference>("regardingobjectid").Id;
                        }

                        EntityReference _entSMSTemplate = entity.GetAttributeValue<EntityReference>("hil_smstemplate");
                        Entity _smsTemplate = service.Retrieve("hil_smstemplates", _entSMSTemplate.Id, new ColumnSet("hil_templateid", "hil_encryptsms"));
                        string _smsTemplateId = _smsTemplate.GetAttributeValue<string>("hil_templateid");
                        string message = SMS.hil_Message;
                        message = message.Replace("#", "%23");
                        message = message.Replace("&", "%26");
                        message = message.Replace("+", "%2B");

                        Entity entTemp = new Entity(entity.LogicalName, entity.Id);
                        if (_smsTemplateId != "1107161191448698079" && _smsTemplateId != "1107161191438154934")
                        {
                            SMS = HelperShootSMS.OnDemandSMSShootFunctionDLT(service, message, SMS.hil_MobileNumber, SMS, _smsTemplateId, SMS.ActivityAdditionalParams, _regardingObjectId);

                            entTemp["hil_responsefromserver"] = SMS["hil_responsefromserver"];
                            entTemp["hil_message"] = SMS["hil_message"];
                            //entTemp["statuscode"] = new OptionSetValue(2); // Sent|Pending
                            //entTemp["statecode"] = new OptionSetValue(1); // Mark as completed
                        }
                        if (_smsTemplate.GetAttributeValue<bool>("hil_encryptsms"))
                        {
                            entTemp["hil_message"] = HelperShootSMS.EncryptString(HelperShootSMS.EncriptionKey, (string)entity["hil_message"]);
                            entTemp["hil_encrypted"] = true;
                        }
                        service.Update(entTemp);

                        //SetStateRequest req = new SetStateRequest();
                        //req.State = new OptionSetValue(1);
                        //req.Status = new OptionSetValue(2);
                        //req.EntityMoniker = SMS.ToEntityReference();
                        //SetStateResponse res = (SetStateResponse)service.Execute(req);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.SMS_PostOperation_Create.ResolveMessage" + ex.Message);
            }
        }
    }
}
