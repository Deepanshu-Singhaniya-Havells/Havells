using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Security.Cryptography;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Customer_APIs
{
    public class ValidateKKGCodeAction : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion

            string StatusCode = "";
            bool? IsVarified = false;
            string StatusDescription = "";
            try
            {
                if (context.InputParameters.Contains("JobId") && context.InputParameters["JobId"] is string
                        && context.InputParameters.Contains("KKGCode") && context.InputParameters["KKGCode"] is string)
                {
                    try
                    {
                        string _jobId = (string)context.InputParameters["JobId"];
                        string _kkgCode = (string)context.InputParameters["KKGCode"];

                        QueryExpression query = new QueryExpression("hil_jobsauth");
                        query.ColumnSet = new ColumnSet("hil_checksum", "hil_hash", "hil_salt");
                        query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, _jobId));
                        query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1));
                        EntityCollection entColl = service.RetrieveMultiple(query);
                        if (entColl.Entities.Count == 1)
                        {
                            string Checksum = entColl[0].GetAttributeValue<string>("hil_checksum");
                            string _HashOld = entColl[0].GetAttributeValue<string>("hil_hash");
                            string _salt = entColl[0].GetAttributeValue<string>("hil_salt");
                            string _NewHash = getHash(_salt + _kkgCode);
                            if (_NewHash == _HashOld)
                            {
                                IsVarified = true;
                                StatusDescription = "Success";
                            }
                            else
                            {
                                IsVarified = false;
                                StatusDescription = "Invalid KKG Code";
                            }
                        }
                        else
                        {
                            query = new QueryExpression("msdyn_workorder");
                            query.ColumnSet = new ColumnSet("hil_kkgotp");
                            query.Criteria.AddCondition(new ConditionExpression("msdyn_name", ConditionOperator.Equal, _jobId));
                            entColl = service.RetrieveMultiple(query);
                            if (entColl.Entities.Count == 1)
                            {
                                string kkgOtp = entColl[0].GetAttributeValue<string>("hil_kkgotp");
                                var base64Bytes = System.Convert.FromBase64String(kkgOtp);
                                kkgOtp = Encoding.UTF8.GetString(base64Bytes);
                                if (kkgOtp == _kkgCode)
                                {
                                    IsVarified = true;
                                    StatusCode = "200";
                                    StatusDescription = "Success";
                                }
                                else
                                {
                                    IsVarified = false;
                                    StatusCode = "204";
                                    StatusDescription = "Invalid KKG Code";
                                }
                            }
                            else
                            {
                                IsVarified = false;
                                StatusCode = "503";
                                StatusDescription = "Invalid Job Id.";
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        IsVarified = false;
                        StatusCode = "500";
                        StatusDescription = "D365 Internal Server Error : " + ex.Message;
                    }
                }
            }
            catch (Exception ex)
            {
                IsVarified = false;
                StatusCode = "500";
                StatusDescription = "D365 Internal Server Error : " + ex.Message;
            }
            finally
            {
                context.OutputParameters["IsVarified"] = IsVarified;
                context.OutputParameters["StatusCode"] = StatusCode;
                context.OutputParameters["StatusDescription"] = StatusDescription;
            }
        }

        private string getHash(string text)
        {
            using (var sha512 = SHA512.Create())
            {
                var hashedBytes = sha512.ComputeHash(Encoding.UTF8.GetBytes(text));
                return BitConverter.ToString(hashedBytes).Replace("-", "").ToLower();
            }
        }
    }
}
