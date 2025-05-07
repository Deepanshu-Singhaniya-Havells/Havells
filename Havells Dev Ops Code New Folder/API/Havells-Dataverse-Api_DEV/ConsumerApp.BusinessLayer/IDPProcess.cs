using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RestSharp;
using System;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class IDPProcess
    {
        public IDPProcessResponse UpdateIDPProcess(IDPProcessDetails iDPProcessDetails)
        {
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (string.IsNullOrWhiteSpace(iDPProcessDetails.Guid))
                {
                    return new IDPProcessResponse { status = false, message = "Guid is required." };
                }
                if (string.IsNullOrWhiteSpace(iDPProcessDetails.DocumentActualName))
                {
                    return new IDPProcessResponse { status = false, message = "Document actual name required." };
                }
                if (string.IsNullOrWhiteSpace(iDPProcessDetails.DocumentName))
                {
                    return new IDPProcessResponse { status = false, message = "Document name required." };
                }
                int[] IDPProcessStage = { 1, 2, 3, 4 };
                if (!IDPProcessStage.Contains(iDPProcessDetails.IDPProcessStep))
                {
                    return new IDPProcessResponse { status = false, message = "IDP process step must be 2-Error, 3-Processing or 4-Generated." };
                }
                else
                {
                    if (iDPProcessDetails.IDPProcessStep == (int)IDPProcessStepEnum.Generated)
                    {
                        if (iDPProcessDetails.IDPGeneratedFiles != null)
                        {
                            if (iDPProcessDetails.IDPGeneratedFiles.Length > 4 || iDPProcessDetails.IDPGeneratedFiles.Length < 3)
                            {
                                return new IDPProcessResponse { status = false, message = "IDP generated files must be 3 or 4." };
                            }
                        }
                        else
                        {
                            return new IDPProcessResponse { status = false, message = "IDP generated files required" };
                        }
                    }
                }
                if (service != null)
                {
                    if (iDPProcessDetails.IDPProcessStep == (int)IDPProcessStepEnum.Error
                        || iDPProcessDetails.IDPProcessStep == (int)IDPProcessStepEnum.Initiated)
                    {
                        string ProcessStatus = "";
                        string EmailId = "";
                        string TenderNo = iDPProcessDetails.DocumentActualName.Split('_')[0].Trim();
                        QueryExpression query = new QueryExpression("hil_tender");
                        query.ColumnSet = new ColumnSet("ownerid");
                        query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, TenderNo);//TEND00051251
                        Entity entityTender = service.RetrieveMultiple(query).Entities[0];
                        Guid OwnerId = entityTender.Contains("ownerid") ? entityTender.GetAttributeValue<EntityReference>("ownerid").Id : Guid.Empty;
                        if (OwnerId != Guid.Empty)
                        {
                            Entity entUser = service.Retrieve("systemuser", OwnerId, new ColumnSet("internalemailaddress"));
                            EmailId = entUser.Contains("internalemailaddress") ? entUser.GetAttributeValue<string>("internalemailaddress") : null;
                            if (EmailId == null)
                            {
                                return new IDPProcessResponse { status = false, message = "Email Id is not set for owner." };
                            }
                        }
                        else
                        {
                            return new IDPProcessResponse { status = false, message = "owner is not found." };
                        }
                        IDPProcessResponse result = IDPProcessInitiate(new IDPInitiateProcess
                        {
                            Empmail = EmailId,
                            n_GUID = iDPProcessDetails.Guid,//AttachmentId
                            s_DocumentActualName = iDPProcessDetails.DocumentActualName,//"testing~TEND00051251_GTP_202402150538248658.pdf", 
                            s_DocumentName = iDPProcessDetails.DocumentName, //subject(Description)
                            source = "EMS"
                        }, service);
                        if (!result.status || (result.status && iDPProcessDetails.IDPProcessStep == (int)IDPProcessStepEnum.Error))
                        {
                            Entity entTender = new Entity("hil_attachment", new Guid(iDPProcessDetails.Guid));
                            entTender["hil_idpprocesssteps"] = new OptionSetValue(2); // Error
                            service.Update(entTender);
                            ProcessStatus = "Something Occur wrong, Please try again after sometime.";
                        }
                        else
                        {
                            Entity entTender = new Entity("hil_attachment", new Guid(iDPProcessDetails.Guid));
                            entTender["hil_idpprocesssteps"] = new OptionSetValue(1); // Process Initiated
                            service.Update(entTender);
                            ProcessStatus = "IDP Process Initiated";
                        }
                        return new IDPProcessResponse { status = true, message = ProcessStatus };
                    }
                    else if (iDPProcessDetails.IDPProcessStep == (int)IDPProcessStepEnum.Generated)
                    {
                        Entity entityAttch = new Entity("hil_attachment", new Guid(iDPProcessDetails.Guid));
                        entityAttch["hil_idpprocesssteps"] = new OptionSetValue(iDPProcessDetails.IDPProcessStep);
                        foreach (var item in iDPProcessDetails.IDPGeneratedFiles)
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
                        return new IDPProcessResponse { status = true, message = "IDP Process Completed." };
                    }
                    else if (iDPProcessDetails.IDPProcessStep == (int)IDPProcessStepEnum.Processing)
                    {
                        Entity entTender = new Entity("hil_attachment", new Guid(iDPProcessDetails.Guid));
                        entTender["hil_idpprocesssteps"] = new OptionSetValue(3); // Default Processing
                        service.Update(entTender);
                        return new IDPProcessResponse { status = true, message = "IDP is under Processing." };
                    }
                    return new IDPProcessResponse { status = true, message = "Success." };
                }
                else
                {
                    return new IDPProcessResponse { status = false, message = "D365 service unavailable..." };
                }
            }
            catch (Exception ex)
            {
                return new IDPProcessResponse { status = false, message = ex.Message };
            }
        }
        public IDPProcessResponse IDPProcessInitiate(IDPInitiateProcess requestParm, IOrganizationService service)
        {
            try
            {
                Integration intConFig = IntegrationConfiguration(service, "IDPProcess");
                string _authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConFig.Auth));
                var client = new RestClient(intConFig.uri);
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);
                request.AddHeader("Authorization", _authInfo);
                request.AddHeader("Content-Type", "application/json");
                String body = Newtonsoft.Json.JsonConvert.SerializeObject(requestParm);
                request.AddParameter("application/json", body, ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                Console.WriteLine(response.Content);
                IDPProcessResponse rootObject = Newtonsoft.Json.JsonConvert.DeserializeObject<IDPProcessResponse>(response.Content);
                return rootObject;
            }
            catch (Exception ex)
            {
                return new IDPProcessResponse() { status = false, message = ex.Message };
            }
        }
        public Integration IntegrationConfiguration(IOrganizationService service, string Param)
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
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
            return output;
        }

    }
    public class IDPInitiateProcess
    {
        [DataMember]
        public string s_DocumentName { get; set; }
        [DataMember]
        public string s_DocumentActualName { get; set; }
        [DataMember]
        public string source { get; set; }
        [DataMember]
        public string n_GUID { get; set; }
        public string Empmail { get; set; }
    }
    public class IDPProcessResponse
    {
        [DataMember]
        public string message { get; set; }
        [DataMember]
        public bool status { get; set; }
    }
    public class IDPProcessDetails
    {
        [DataMember]
        public string Guid { get; set; }
        [DataMember]
        public string DocumentActualName { get; set; }
        [DataMember]
        public string DocumentName { get; set; }
        [DataMember]
        public int IDPProcessStep { get; set; }
        [DataMember]
        public string[] IDPGeneratedFiles { get; set; }
    }
    public enum IDPProcessStepEnum
    {
        Initiated = 1,
        Error = 2,
        Processing = 3,
        Generated = 4
    }
}
