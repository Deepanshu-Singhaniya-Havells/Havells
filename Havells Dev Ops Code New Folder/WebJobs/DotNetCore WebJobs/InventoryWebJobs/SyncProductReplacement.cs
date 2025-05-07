using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Net;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace InventoryWebJobs.ProductReplacement
{
    public class SyncProductReplacement
    {
        private IOrganizationService service;

        public SyncProductReplacement(IOrganizationService _service)
        {
            service = _service;
        }

        internal void GetProductRequisition()
        {
            Entity iUpdateHeader = new Entity("hil_productrequestheader");
            string PO_Packet = string.Empty;

            string sUserName = "D365_Havells";
            string sPassword = "PRDD365@1234";

            //string sUserName = "D365_Havells";
            //string sPassword = "QAD365@1234";

            int j = 0;
            int i = 0;
            string SAP_CODE = string.Empty;
            string Date = string.Empty;
            Guid HeaderId = new Guid();

            QueryExpression Query = new QueryExpression("hil_productrequestheader");
            Query.ColumnSet = new ColumnSet(true);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("hil_customersapcode", ConditionOperator.NotNull));
            Query.Criteria.AddCondition(new ConditionExpression("hil_syncstatus", ConditionOperator.Equal, 1)); //Pending for Submission
            Query.Criteria.AddCondition(new ConditionExpression("hil_ivusage", ConditionOperator.NotNull));
            Query.Criteria.AddCondition(new ConditionExpression("hil_prtype", ConditionOperator.Equal, 910590002));
            Query.AddOrder("modifiedon", OrderType.Descending);
            //Query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.In, new string[] {
            //"PH20-1769183"}));

            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                HeaderId = Found.Entities[0].Id;
                PO_SAP_Request SAPrequest = new PO_SAP_Request();
                IntegrationLog il = new IntegrationLog();

                //Guid integrationHeaderId = IntegrationLogUtils.CreateLogHeader(service, il);

                foreach (Entity PrHeader in Found.Entities)
                {
                    try
                    {

                        int syncStatus = PrHeader.GetAttributeValue<OptionSetValue>("hil_syncstatus").Value;
                        DateTime iDate = PrHeader.GetAttributeValue<DateTime>("createdon");
                        // if(true)
                        if (syncStatus == 1 || syncStatus == 3 || syncStatus == 2) // 1- Pending submission,   2 - Sync success,   3- Sync failed
                        {
                            i = Found.Entities.Count;
                            if (iDate != null)
                            {
                                Date = iDate.AddMinutes(330).ToString("yyyy-MM-dd");
                            }

                            Parent iParent = new Parent();
                            iParent.SPART = PrHeader.GetAttributeValue<string>("hil_divisioncode");
                            iParent.BSTKD = PrHeader.GetAttributeValue<string>("hil_name");

                            iParent.BSTDK = Date;
                            iParent.KUNNR = PrHeader.GetAttributeValue<string>("hil_customersapcode");
                            iParent.ABRVW = (string)PrHeader["hil_ivusage"];
                            iParent.VKORG = "HIL";

                            if (PrHeader.Contains("hil_warrantystatus") && PrHeader.GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value == 1 || PrHeader.GetAttributeValue<string>("hil_customersapcode").StartsWith("E"))
                            {
                                iParent.AUART = "ZRS4";
                            }
                            else if (PrHeader.Contains("hil_warrantystatus") && PrHeader.GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value == 2 || PrHeader.GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value == 3 || PrHeader.GetAttributeValue<OptionSetValue>("hil_warrantystatus").Value == 2)
                            {
                                iParent.AUART = "ZRS4";
                            }
                            // OutWarranty C Code ZSW6

                            if (PrHeader.GetAttributeValue<string>("hil_customersapcode").StartsWith("E"))
                            {
                                iParent.VTWEG = getDistributionChannel(PrHeader.GetAttributeValue<OptionSetValue>("hil_warrantystatus"), new OptionSetValue(9));// PRLine.hil_DistributionChannel;
                            }
                            else
                            {
                                if (PrHeader.GetAttributeValue<string>("hil_customersapcode").StartsWith("C"))
                                {
                                    iParent.VTWEG = getDistributionChannel(new OptionSetValue(2), new OptionSetValue(6));// PRLine.hil_DistributionChannel;
                                }
                                else if (PrHeader.GetAttributeValue<string>("hil_customersapcode").StartsWith("F"))
                                {
                                    iParent.VTWEG = getDistributionChannel(new OptionSetValue(1), new OptionSetValue(6));// PRLine.hil_DistributionChannel;
                                }
                                //In Warranty - 56 {FCode}
                                //Out Warranty - 55 {CCode}
                            }

                            IMPROJECT project = new IMPROJECT();
                            project.IM_PROJECT = "D365";

                            QueryExpression Query1 = new QueryExpression("hil_productrequest");
                            Query1.ColumnSet = new ColumnSet(true);
                            Query1.Criteria = new FilterExpression(LogicalOperator.And);
                            Query1.Criteria.AddCondition(new ConditionExpression("hil_prheader", ConditionOperator.Equal, PrHeader.Id));
                            Query1.Criteria.AddCondition(new ConditionExpression("hil_quantity", ConditionOperator.NotNull));
                            Query1.Criteria.AddCondition(new ConditionExpression("hil_partcode", ConditionOperator.NotNull));
                            Query1.Criteria.AddCondition(new ConditionExpression("hil_potypecode", ConditionOperator.NotNull));
                            EntityCollection Found1 = service.RetrieveMultiple(Query1);
                            if (Found1.Entities.Count > 0)
                            {
                                List<IT_DATA> childs = new List<IT_DATA>();
                                foreach (Entity PRLine in Found1.Entities)
                                {
                                    IT_DATA iChild = new IT_DATA();
                                    iChild.MATNR = PRLine.GetAttributeValue<EntityReference>("hil_partcode").Name;
                                    iChild.SPART = PRLine.GetAttributeValue<string>("hil_divisionsapcode");
                                    iChild.DZMENG = PRLine.GetAttributeValue<int>("hil_quantity");
                                    childs.Add(iChild);
                                }
                                SAPrequest.IM_PROJECT = "D365";
                                SAPrequest.IM_HEADER = iParent;
                                SAPrequest.LT_LINE_ITEM = childs;
                                var Json = JsonConvert.SerializeObject(SAPrequest);
                                //Production URL
                                WebRequest request = WebRequest.Create("https://p90ci.havells.com:50001/RESTAdapter/Common/CreateSalesOrder");
                                //WebRequest request = WebRequest.Create("https://middleware.havells.com:50001/RESTAdapter/Common/CreateSalesOrder");
                                //UAT URL
                                //WebRequest request = WebRequest.Create("https://middlewareqa.havells.com:50001/RESTAdapter/Common/CreateSalesOrder");
                                string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
                                request.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
                                // Set the Method property of the request to POST.  
                                request.Method = "POST";
                                byte[] byteArray = Encoding.UTF8.GetBytes(Json);
                                // Set the ContentType property of the WebRequest.  
                                request.ContentType = "application/x-www-form-urlencoded";
                                // Set the ContentLength property of the WebRequest.  
                                request.ContentLength = byteArray.Length;
                                // Get the request stream.  
                                Stream dataStream = request.GetRequestStream();
                                // Write the data to the request stream.  
                                dataStream.Write(byteArray, 0, byteArray.Length);
                                // Close the Stream object.  
                                dataStream.Close();
                                // Get the response.  
                                WebResponse response = request.GetResponse();
                                // Display the status.  
                                Console.WriteLine(((HttpWebResponse)response).StatusDescription);
                                // Get the stream containing content returned by the server.  
                                dataStream = response.GetResponseStream();
                                // Open the stream using a StreamReader for easy access.  
                                StreamReader reader = new StreamReader(dataStream);
                                // Read the content.  
                                string responseFromServer = reader.ReadToEnd();
                                // responseFromServer = responseFromServer.Replace("ZBAPI_CREATE_SALES_ORDER.Response", "ZBAPI_CREATE_SALES_ORDER"); // removed for new response
                                Response resp = JsonConvert.DeserializeObject<Response>(responseFromServer);
                                Console.WriteLine(resp.EX_SALESDOC_NO + " - " + resp.RETURN + " - " + PrHeader.GetAttributeValue<string>("hil_name"));
                                if (resp.EX_SALESDOC_NO != "")
                                {
                                    iUpdateHeader.Id = PrHeader.Id;
                                    iUpdateHeader["hil_message"] = resp.RETURN;
                                    iUpdateHeader["hil_salesorderno"] = resp.EX_SALESDOC_NO.ToString();
                                    iUpdateHeader["hil_packet"] = Json.ToString();
                                    iUpdateHeader["hil_syncstatus"] = new OptionSetValue(2);    //Sync success
                                    iUpdateHeader["hil_retrycount"] = Convert.ToInt32(0);
                                    service.Update(iUpdateHeader);
                                }
                                else
                                {
                                    iUpdateHeader.Id = PrHeader.Id;
                                    iUpdateHeader["hil_message"] = resp.RETURN;
                                    iUpdateHeader["hil_salesorderno"] = resp.EX_SALESDOC_NO.ToString();
                                    iUpdateHeader["hil_syncstatus"] = new OptionSetValue(3);            //Sync failed
                                    iUpdateHeader["hil_packet"] = Json.ToString();
                                    iUpdateHeader["hil_retrycount"] = Convert.ToInt32(0);
                                    service.Update(iUpdateHeader);
                                }
                                QueryExpression Query2 = new QueryExpression("hil_productrequest");
                                Query2.ColumnSet = new ColumnSet(false);
                                Query2.Criteria = new FilterExpression(LogicalOperator.And);
                                Query2.Criteria.AddCondition(new ConditionExpression("hil_prheader", ConditionOperator.Equal, PrHeader.Id));
                                EntityCollection Found2 = service.RetrieveMultiple(Query2);
                                if (Found2.Entities.Count > 0)
                                {
                                    Entity ent;
                                    foreach (Entity PRLine in Found2.Entities)
                                    {
                                        PRLine["hil_salesorderno"] = resp.EX_SALESDOC_NO.ToString();
                                        service.Update(PRLine);

                                        //Update Message on PR Line
                                        ent = new Entity("hil_productrequest", PRLine.Id);
                                        ent["hil_message"] = resp.RETURN;
                                        service.Update(ent);
                                    }
                                }
                                Console.WriteLine("DONE : " + ++j + ", TOTAL : " + Found.Entities.Count);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        continue;
                    }
                }
            }
        }


        private String getDistributionChannel(OptionSetValue opWarrantyStatus, OptionSetValue opAccountType)
        {
            String result = String.Empty;
            try
            {
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {


                    QueryExpression query = new QueryExpression("hil_integrationconfiguration");
                    query.ColumnSet = new ColumnSet("new_distributionchannel");
                    query.Criteria.AddCondition("new_warrantystatus", ConditionOperator.Equal, opWarrantyStatus.Value);
                    query.Criteria.AddCondition("new_customertype", ConditionOperator.Equal, opAccountType.Value);

                    EntityCollection confColl = service.RetrieveMultiple(query);


                    if (confColl.Entities.Count > 0)
                    {
                        foreach (Entity conf in confColl.Entities)
                        {
                            if (conf.Contains("new_distributionchannel"))
                            {
                                result = conf.GetAttributeValue<string>("new_distributionchannel");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Havells_Plugin.HelpserClasses.HelperPO.getDistributionChannel" + ex.Message);
            }
            return result;
        }


    }
}