using HavellsNewPlugin.TenderModule.OrderCheckList;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.IO;
using System.Security.Cryptography;

namespace HavellsNewPlugin.Actions
{
    public class GetSamparkURL : IPlugin
    {
        public static ITracingService tracingService = null;
        public static Integration IntegrationConfiguration(IOrganizationService service, string Param)
        {
            Integration output = new Integration();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                if (integrationConfiguration.Contains("hil_username") && integrationConfiguration.Contains("hil_password"))
                {
                    output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error:- " + ex.Message);
            }
            return output;
        }
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
                if (context.InputParameters.Contains("UserId") && context.InputParameters["UserId"] is string && context.Depth == 1)
                {
                    try
                    {
                        string UserId = (string)context.InputParameters["UserId"];
                        Entity entTemp = service.Retrieve("systemuser", new Guid(UserId), new ColumnSet("hil_employeecode"));
                        if (entTemp != null)
                        {
                            //_retObj.user_id = entTemp.GetAttributeValue<string>("hil_employeecode");
                            string access_token = entTemp.GetAttributeValue<string>("hil_employeecode");
                            if (!string.IsNullOrEmpty(access_token))
                            {
                                access_token = EncryptAES256(access_token);
                            }
                            string baseURL = IntegrationConfiguration(service, "SamparkAppUrl").uri;
                            context.OutputParameters["Uri"] = baseURL + access_token;
                            context.OutputParameters["status_code"] = "204";
                            context.OutputParameters["status_description"] = "User does not exist";
                        }
                        else
                        {
                            context.OutputParameters["status_code"] = "204";
                            context.OutputParameters["status_description"] = "User does not exist";
                        }
                    }
                    catch (Exception ex)
                    {
                        context.OutputParameters["status_code"] = "400";
                        context.OutputParameters["status_description"] = "D365 Internal Server Error : " + ex.Message;
                    }
                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["status_code"] = "400";
                context.OutputParameters["status_description"] = "D365 Internal Server Error : " + ex.Message;
            }
        }
        private static string EncryptAES256(string plainText)
        {
            string Key = "DklsdvkfsDlkslsdsdnv234djSDAjkd1";
            byte[] key32 = Encoding.UTF8.GetBytes(Key);
            byte[] IV16 = Encoding.UTF8.GetBytes(Key.Substring(0, 16)); if (plainText == null || plainText.Length <= 0)
            {
                throw new ArgumentNullException("plainText");
            }
            byte[] encrypted;
            using (AesManaged aesAlg = new AesManaged())
            {
                aesAlg.KeySize = 256;
                aesAlg.Padding = PaddingMode.PKCS7;
                aesAlg.IV = IV16;
                aesAlg.Key = key32;
                ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
                using (MemoryStream msEncrypt = new MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                        {
                            swEncrypt.Write(plainText);
                        }
                        encrypted = msEncrypt.ToArray();
                    }
                }
            }
            return Convert.ToBase64String(encrypted);
        }
    }
}
