using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Text;
using System.IO;
using System.Security.Cryptography;

namespace HavellsNewPlugin.SMS
{
    public class ResendSMS : IPlugin
    {
        public static ITracingService tracingService = null;
        public static string EncriptionKey = "12s2s121sasfdasdf45346fwrt3w56fw";
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                var entityId = context.InputParameters["EntityID"].ToString();
                if (entityId == null)
                {
                    context.OutputParameters["Status"] = "Failed !!";
                    context.OutputParameters["Message"] = "Invalid SMS GUID";
                }
                else
                {
                    Entity _oldSms = service.Retrieve("hil_smsconfiguration", new Guid(entityId), new ColumnSet(true));
                    EntityReference _smsTemplateID = _oldSms.GetAttributeValue<EntityReference>("hil_smstemplate");
                    //Entity _smsTemplate = service.Retrieve(_smsTemplateID.LogicalName, _smsTemplateID.Id, new ColumnSet("hil_encryptsms"));
                    Entity _newSMS = new Entity(_oldSms.LogicalName);
                    if (_oldSms.GetAttributeValue<bool>("hil_encrypted"))
                    {
                        _newSMS["hil_message"] = DecryptString(EncriptionKey, (string)_oldSms["hil_message"]);
                    }
                    else
                        _newSMS["hil_message"] = _oldSms["hil_message"];
                    var attributes = _oldSms.Attributes.Keys;

                    foreach (string name in attributes)
                    {
                        // if (name != "adx_partnervisible")
                        if (name != "modifiedby" && name != "createdby" && name != "ownerid" && name != "owninguser" && name != "statecode"
                            && name != "statuscode" && name != "hil_responsefromserver" && name != "hil_message" && name != "activityid" && name != "hil_encrypted"
                            && name != "createdon" && name != "modifiedon" && name != "owningbusinessunit" && name != "timezoneruleversionnumber"
                            && name != "activitytypecode" && name != "instancetypecode" && name != "isworkflowcreated" && name != "processid")
                        {
                            _newSMS[name] = _oldSms[name];
                        }
                    }
                    Guid _newSMSID = service.Create(_newSMS);
                    context.OutputParameters["Status"] = "Successful !!";
                    context.OutputParameters["Message"] = "SMS Resend Successfully";
                }

            }
            catch (Exception ex)
            {
                context.OutputParameters["Status"] = "Failed !!";
                context.OutputParameters["Message"] = ex.Message;
               // throw new InvalidPluginExecutionException("Error " + ex.Message);
            }
        }

        public static string DecryptString(string key, string cipherText)
        {
            byte[] iv = new byte[16];
            byte[] buffer = Convert.FromBase64String(cipherText);

            using (Aes aes = Aes.Create())
            {
                aes.Key = Encoding.UTF8.GetBytes(key);
                aes.IV = iv;
                ICryptoTransform decryptor = aes.CreateDecryptor(aes.Key, aes.IV);

                using (MemoryStream memoryStream = new MemoryStream(buffer))
                {
                    using (CryptoStream cryptoStream = new CryptoStream((Stream)memoryStream, decryptor, CryptoStreamMode.Read))
                    {
                        using (StreamReader streamReader = new StreamReader((Stream)cryptoStream))
                        {
                            return streamReader.ReadToEnd();
                        }
                    }
                }
            }
        }
    }
}
