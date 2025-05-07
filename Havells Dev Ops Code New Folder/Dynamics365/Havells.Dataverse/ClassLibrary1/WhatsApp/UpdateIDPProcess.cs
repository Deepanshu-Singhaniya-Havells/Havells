using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;

namespace Havells.Dataverse.CustomConnector.WhatsApp_CTA
{
    public class UpdateIDPProcess : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            List<IDPProcessResponse> UpdateIDPProcess = new List<IDPProcessResponse>();
            string JsonResponse = "";
            try
            {
                string Guid = Convert.ToString(context.InputParameters["Guid"]);
                string DocumentActualName = Convert.ToString(context.InputParameters["DocumentActualName"]);
                string DocumentName = Convert.ToString(context.InputParameters["DocumentName"]);
                int IDPProcessStep = Convert.ToInt32(context.InputParameters["IDPProcessStep"]);
                string filesArray = Convert.ToString(context.InputParameters["IDPGeneratedFiles"]);
                string[] IDPGeneratedFiles = filesArray.Split(',');

                if (string.IsNullOrWhiteSpace(Guid))
                {
                    UpdateIDPProcess.Add(new IDPProcessResponse()
                    {
                        status = false,
                        message = "No Data found."
                    });
                    JsonResponse = JsonConvert.SerializeObject(UpdateIDPProcess);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                if (string.IsNullOrWhiteSpace(DocumentActualName))
                {
                    UpdateIDPProcess.Add(new IDPProcessResponse()
                    {
                        status = false,
                        message = "Document actual name required."
                    });
                    JsonResponse = JsonConvert.SerializeObject(UpdateIDPProcess);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                if (string.IsNullOrWhiteSpace(DocumentName))
                {
                    UpdateIDPProcess.Add(new IDPProcessResponse()
                    {
                        status = false,
                        message = "Document name required."
                    });
                    JsonResponse = JsonConvert.SerializeObject(UpdateIDPProcess);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                int[] IDPProcessStage = { 1, 2, 3, 4 };
                if (!IDPProcessStage.Contains(IDPProcessStep))
                {
                    UpdateIDPProcess.Add(new IDPProcessResponse()
                    {
                        status = false,
                        message = "IDP process step must be 2-Error, 3-Processing or 4-Generated."
                    });
                    JsonResponse = JsonConvert.SerializeObject(UpdateIDPProcess);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                else
                {
                    if (IDPProcessStep == (int)IDPProcessStepEnum.Generated)
                    {
                        if (IDPGeneratedFiles != null)
                        {
                            if (IDPGeneratedFiles.Length > 4 || IDPGeneratedFiles.Length < 3)
                            {
                                UpdateIDPProcess.Add(new IDPProcessResponse()
                                {
                                    status = false,
                                    message = "IDP generated files must be 3 or 4."
                                });
                                JsonResponse = JsonConvert.SerializeObject(UpdateIDPProcess);
                                context.OutputParameters["data"] = JsonResponse;
                                return;
                            }
                        }
                        else
                        {
                            UpdateIDPProcess.Add(new IDPProcessResponse()
                            {
                                status = false,
                                message = "IDP generated files required"
                            });
                            JsonResponse = JsonConvert.SerializeObject(UpdateIDPProcess);
                            context.OutputParameters["data"] = JsonResponse;
                            return;
                        }
                    }
                }
                if (service != null)
                {
                    if (IDPProcessStep == (int)IDPProcessStepEnum.Error
                        || IDPProcessStep == (int)IDPProcessStepEnum.Initiated)
                    {
                        string ProcessStatus = "";
                        string EmailId = "";
                        string TenderNo = DocumentActualName.Split('_')[0].Trim();
                        QueryExpression query = new QueryExpression("hil_tender");
                        query.ColumnSet = new ColumnSet("ownerid");
                        query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, TenderNo);//TEND00051251
                        Entity entityTender = service.RetrieveMultiple(query).Entities[0];
                        Guid OwnerId = entityTender.Contains("ownerid") ? entityTender.GetAttributeValue<EntityReference>("ownerid").Id : System.Guid.Empty;
                        if (OwnerId != System.Guid.Empty)
                        {
                            Entity entUser = service.Retrieve("systemuser", OwnerId, new ColumnSet("internalemailaddress"));
                            EmailId = entUser.Contains("internalemailaddress") ? entUser.GetAttributeValue<string>("internalemailaddress") : null;
                            if (EmailId == null)
                            {
                                UpdateIDPProcess.Add(new IDPProcessResponse()
                                {
                                    status = false,
                                    message = "Email Id is not set for owner."
                                });
                                JsonResponse = JsonConvert.SerializeObject(UpdateIDPProcess);
                                context.OutputParameters["data"] = JsonResponse;
                                return;
                            }
                        }
                        else
                        {
                            UpdateIDPProcess.Add(new IDPProcessResponse()
                            {
                                status = false,
                                message = "owner is not found."
                            });
                            JsonResponse = JsonConvert.SerializeObject(UpdateIDPProcess);
                            context.OutputParameters["data"] = JsonResponse;
                            return;
                        }
                        IDPProcessResponse result = IDPProcessInitiate(new IDPInitiateProcess
                        {
                            Empmail = EmailId,
                            n_GUID = Guid,//AttachmentId
                            s_DocumentActualName = DocumentActualName,//"testing~TEND00051251_GTP_202402150538248658.pdf", 
                            s_DocumentName = DocumentName, //subject(Description)
                            source = "EMS"
                        }, service);
                        if (!result.status || (result.status && IDPProcessStep == (int)IDPProcessStepEnum.Error))
                        {
                            Entity entTender = new Entity("hil_attachment", new Guid(Guid));
                            entTender["hil_idpprocesssteps"] = new OptionSetValue(2); // Error
                            service.Update(entTender);
                            ProcessStatus = "Something Occur wrong, Please try again after sometime.";
                        }
                        else
                        {
                            Entity entTender = new Entity("hil_attachment", new Guid(Guid));
                            entTender["hil_idpprocesssteps"] = new OptionSetValue(1); // Process Initiated
                            service.Update(entTender);
                            ProcessStatus = "IDP Process Initiated";
                        }
                        UpdateIDPProcess.Add(new IDPProcessResponse()
                        {
                            status = true,
                            message = ProcessStatus
                        });
                        JsonResponse = JsonConvert.SerializeObject(UpdateIDPProcess);
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                    else if (IDPProcessStep == (int)IDPProcessStepEnum.Generated)
                    {
                        Entity entityAttch = new Entity("hil_attachment", new Guid(Guid));
                        entityAttch["hil_idpprocesssteps"] = new OptionSetValue(IDPProcessStep);
                        foreach (var item in IDPGeneratedFiles)
                        {
                            string fileType = item.Split('_').Last().Split('.')[0];
                            if (fileType.ToLower() == "summary")
                            {
                                entityAttch["hil_summaryfileurl"] = item;
                            }
                            else if (fileType.ToLower() == "tables")
                            {
                                entityAttch["hil_csvfileurl"] = item;
                            }
                            else if (fileType.ToLower() == "visual")
                            {
                                entityAttch["hil_htmlfileurl"] = item;
                            }
                            else if (fileType.ToLower() == "ocr")
                            {
                                entityAttch["hil_ocrfileurl"] = item;
                            }
                        }
                        service.Update(entityAttch);
                        UpdateIDPProcess.Add(new IDPProcessResponse()
                        {
                            status = true,
                            message = "IDP Process Completed."
                        });
                        JsonResponse = JsonConvert.SerializeObject(UpdateIDPProcess);
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                    else if (IDPProcessStep == (int)IDPProcessStepEnum.Processing)
                    {
                        Entity entTender = new Entity("hil_attachment", new Guid(Guid));
                        entTender["hil_idpprocesssteps"] = new OptionSetValue(3); // Default Processing
                        service.Update(entTender);
                        UpdateIDPProcess.Add(new IDPProcessResponse()
                        {
                            status = true,
                            message = "IDP is under Processing."
                        });
                        JsonResponse = JsonConvert.SerializeObject(UpdateIDPProcess);
                        context.OutputParameters["data"] = JsonResponse;
                        return;
                    }
                    UpdateIDPProcess.Add(new IDPProcessResponse()
                    {
                        status = true,
                        message = "Success."
                    });
                    JsonResponse = JsonConvert.SerializeObject(UpdateIDPProcess);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
                else
                {
                    UpdateIDPProcess.Add(new IDPProcessResponse()
                    {
                        status = false,
                        message = "D365 service unavailable..."
                    });
                    JsonResponse = JsonConvert.SerializeObject(UpdateIDPProcess);
                    context.OutputParameters["data"] = JsonResponse;
                    return;
                }
            }
            catch (Exception ex)
            {
                UpdateIDPProcess.Add(new IDPProcessResponse()
                {
                    status = false,
                    message = ex.Message
                });
                JsonResponse = JsonConvert.SerializeObject(UpdateIDPProcess);
                context.OutputParameters["data"] = JsonResponse;
                return;
            }
        }
        public static IDPProcessResponse IDPProcessInitiate(IDPInitiateProcess requestParm, IOrganizationService service)
        {
            try
            {
                IDPProcessResponse response = new IDPProcessResponse();

                QueryExpression qe = new QueryExpression("hil_integrationconfiguration");
                qe.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qe.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "IDPProcess");
                Entity enColl = service.RetrieveMultiple(qe)[0];
                String URL = enColl.GetAttributeValue<string>("hil_url");
                String Auth = enColl.GetAttributeValue<string>("hil_username") + ":" + enColl.GetAttributeValue<string>("hil_password");

                var data = new StringContent(JsonConvert.SerializeObject(requestParm), System.Text.Encoding.UTF8, "application/json");
                HttpClient client = new HttpClient();
                var byteArray = Encoding.ASCII.GetBytes(Auth);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));

                Console.WriteLine("Push customer Asset Data to ER API for Earn loyalty points");
                HttpResponseMessage response1 = client.PostAsync(URL, data).Result;
                if (response1.IsSuccessStatusCode)
                {
                    Console.WriteLine("API response code");
                    var result = response1.Content.ReadAsStringAsync().Result;
                    response = JsonConvert.DeserializeObject<IDPProcessResponse>(result);
                }
                return response;
            }
            catch (Exception ex)
            {
                return new IDPProcessResponse() { status = false, message = ex.Message };
            }

        }
        public static IntegrationConfig IntegrationConfiguration(IOrganizationService service, string Param)
        {
            IntegrationConfig output = new IntegrationConfig();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
            return output;
        }
        public class IDPInitiateProcess
        {
            public string s_DocumentName { get; set; }
            public string s_DocumentActualName { get; set; }
            public string source { get; set; }
            public string n_GUID { get; set; }
            public string Empmail { get; set; }
        }
        public class IDPProcessResponse
        {
            public string message { get; set; }
            public bool status { get; set; }
        }
        public class IDPProcessDetails
        {
            public string Guid { get; set; }
            public string DocumentActualName { get; set; }
            public string DocumentName { get; set; }
            public int IDPProcessStep { get; set; }
            public string[] IDPGeneratedFiles { get; set; }
        }
        public class IntegrationConfig
        {
            public string uri { get; set; }
            public string Auth { get; set; }
        }
        public enum IDPProcessStepEnum
        {
            Initiated = 1,
            Error = 2,
            Processing = 3,
            Generated = 4
        }
    }

}

