using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Net;
using System.Text;

namespace InventoryWebJobs
{
    internal class POtoSAP
    {
        private ServiceClient service;

        public POtoSAP(ServiceClient _service)
        {
            this.service = _service;
        }
        internal void SyncPoToSAP()
        {
            IntegrationConfiguration creds = HelperClass.GetIntegrationConfiguration(service, "CreateSalesOrder");

            string Date = string.Empty;

            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                <entity name='hil_inventorypurchaseorder'>
                <all-attributes />
                <order attribute='createdon' descending='true' />
                <filter type='and'>
                    <condition attribute='statecode' operator='eq' value='0' />
                    <condition attribute='hil_postatus' operator='eq' value='3' />
                    <condition attribute='createdon' operator='on-or-after' value='2024-11-01' />
                    <condition attribute='hil_franchise' operator='eq' value='{{f55ac377-4e90-ef11-8a6a-7c1e52327bb2}}' />
                </filter>
                <link-entity name='hil_inventorypurchaseorderline' from='hil_ponumber' to='hil_inventorypurchaseorderid' link-type='inner' alias='ab' />
                </entity>
                </fetch>";

            //QueryExpression Query = new QueryExpression("hil_inventorypurchaseorder");
            //Query.ColumnSet = new ColumnSet(true);
            //Query.Criteria = new FilterExpression(LogicalOperator.And);
            //Query.Criteria.AddCondition(new ConditionExpression("hil_postatus", ConditionOperator.Equal, 3)); //Approved 
            //Query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 1);
            //Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, "2024-11-01");
            ////Query.AddOrder("modifiedon", OrderType.Descending);
            EntityCollection Found = service.RetrieveMultiple(new FetchExpression(_fetchXML));

            //HeaderId = Found.Entities[0].Id;
            PO_SAP_Request SAPrequest = new PO_SAP_Request();
            IntegrationLog il = new IntegrationLog();
            int _rowCount = 1, _totalRows = Found.Entities.Count;

            if (Found.Entities.Count > 0)
            {
                foreach (Entity purchaseOrder in Found.Entities)
                {
                    Console.WriteLine($"Processing Row# {_rowCount++}/{_totalRows} PO# {purchaseOrder.GetAttributeValue<string>("hil_name")}");
                    try
                    {
                        if (purchaseOrder.Contains("createdon"))
                        {
                            DateTime iDate = purchaseOrder.GetAttributeValue<DateTime>("createdon");
                            Date = iDate.AddMinutes(330).ToString("yyyy-MM-dd");
                        }

                        EntityReference productRef = purchaseOrder.GetAttributeValue<EntityReference>("hil_productdivision");
                        Entity product = service.Retrieve(productRef.LogicalName, productRef.Id, new ColumnSet("productnumber", "name"));

                        EntityReference franchiseRef = purchaseOrder.GetAttributeValue<EntityReference>("hil_franchise");
                        Entity franchise = service.Retrieve("account", franchiseRef.Id, new ColumnSet("accountnumber", "customertypecode"));

                        EntityReference divRef = purchaseOrder.GetAttributeValue<EntityReference>("hil_productdivision");
                        Entity div = service.Retrieve(divRef.LogicalName, divRef.Id, new ColumnSet("hil_sapcode"));
                        string SAPCode = div.GetAttributeValue<string>("hil_sapcode");


                        Parent iParent = new Parent();
                        iParent.SPART = SAPCode;//product.Contains("productnumber") ? product.GetAttributeValue<string>("productnumber") : string.Empty;
                        iParent.BSTKD = purchaseOrder.GetAttributeValue<string>("hil_name");


                        iParent.BSTDK = Date;
                        int CustomerType = franchise.GetAttributeValue<OptionSetValue>("customertypecode").Value;
                        string AccountNumber = franchise.GetAttributeValue<string>("accountnumber");
                        if (CustomerType == 6)//Franchisee
                        {
                            iParent.KUNNR = "F" + AccountNumber;
                            iParent.VTWEG = "56";
                        }
                        else
                        {
                            iParent.KUNNR = AccountNumber;
                            iParent.VTWEG = "66";
                        }

                        iParent.VKORG = "HIL";
                        int orderType = purchaseOrder.GetAttributeValue<OptionSetValue>("hil_ordertype").Value;
                        if (orderType == 1)
                        {
                            iParent.ABRVW = "EM";
                        }
                        else
                        {
                            iParent.ABRVW = "S";
                        }
                        iParent.AUART = "ZRS4";


                        IMPROJECT project = new IMPROJECT();
                        project.IM_PROJECT = "D365";

                        QueryExpression Query1 = new QueryExpression("hil_inventorypurchaseorderline");
                        Query1.ColumnSet = new ColumnSet(true);
                        Query1.Criteria = new FilterExpression(LogicalOperator.And);
                        Query1.Criteria.AddCondition(new ConditionExpression("hil_ponumber", ConditionOperator.Equal, purchaseOrder.Id));
                        Query1.Criteria.AddCondition(new ConditionExpression("hil_orderquantity", ConditionOperator.GreaterEqual,0));
                        Query1.Criteria.AddCondition(new ConditionExpression("hil_partcode", ConditionOperator.NotNull));
                        Query1.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); // Active

                        //<condition attribute='statecode' operator='eq' value='0' />
                        EntityCollection Found1 = service.RetrieveMultiple(Query1);
                        if (Found1.Entities.Count > 0)
                        {
                            List<IT_DATA> childs = new List<IT_DATA>();
                            foreach (Entity orderLine in Found1.Entities)
                            {
                                IT_DATA iChild = new IT_DATA();
                                iChild.MATNR = orderLine.GetAttributeValue<EntityReference>("hil_partcode").Name;
                                iChild.SPART = SAPCode;
                                iChild.DZMENG = orderLine.GetAttributeValue<int>("hil_orderquantity");
                                childs.Add(iChild);
                            }
                            SAPrequest.IM_PROJECT = "D365";
                            SAPrequest.IM_HEADER = iParent;
                            SAPrequest.LT_LINE_ITEM = childs;
                            var Json = JsonConvert.SerializeObject(SAPrequest);
                            //Production URL
                            WebRequest request = WebRequest.Create("https://p90ci.havells.com:50001/RESTAdapter/Common/CreateSalesOrder");
                            //WebRequest request = WebRequest.Create(creds.url);
                            //UAT URL
                            //WebRequest request = WebRequest.Create("https://middlewareqa.havells.com:50001/RESTAdapter/Common/CreateSalesOrder");
                            string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(creds.userName + ":" + creds.password));
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
                            Response_SAP resp = JsonConvert.DeserializeObject<Response_SAP>(responseFromServer);
                            Console.WriteLine(resp.EX_SALESDOC_NO + " - " + resp.RETURN + " - " + purchaseOrder.GetAttributeValue<string>("hil_name"));
                            if (resp.EX_SALESDOC_NO != "")
                            {
                                Console.WriteLine($"Processed PO# {purchaseOrder.GetAttributeValue<string>("hil_name")} with SAP SO# {resp.EX_SALESDOC_NO}");

                                purchaseOrder["hil_syncedwithsap"] = true;
                                purchaseOrder["hil_responsefromsap"] = resp.RETURN;
                                purchaseOrder["hil_salesordernumber"] = resp.EX_SALESDOC_NO.ToString();
                                purchaseOrder["hil_postatus"] = new OptionSetValue(4);
                                service.Update(purchaseOrder);
                            }
                            else
                            {
                                Console.WriteLine($"Processed PO# {purchaseOrder.GetAttributeValue<string>("hil_name")} with ERROR# {resp.RETURN}");
                                purchaseOrder["hil_syncedwithsap"] = true;
                                purchaseOrder["hil_responsefromsap"] = resp.RETURN;
                                purchaseOrder["hil_salesordernumber"] = resp.EX_SALESDOC_NO.ToString();

                                service.Update(purchaseOrder);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            else
            {
                Console.WriteLine("No Order found to Sync.");
            }
        }
        private class PO_SAP_Request
        {
            public string IM_PROJECT { get; set; }
            public List<IT_DATA> LT_LINE_ITEM { get; set; }
            public Parent IM_HEADER { get; set; }
        }


        private class Parent
        {
            public string? AUART { get; set; }
            public string? VTWEG { get; set; }
            public string? SPART { get; set; }
            public string? BSTKD { get; set; }
            public string? BSTDK { get; set; }
            public string? KUNNR { get; set; }
            public string? ABRVW { get; set; }
            public string? VKORG { get; set; }
        }


        private class Response_SAP
        {
            public string EX_SALESDOC_NO { get; set; }
            public string RETURN { get; set; }
            public List<ET_Details> ET_SO_DETAILS;
        }
        private class IMPROJECT
        {
            public string IM_PROJECT { get; set; }
        }
        private class IT_DATA
        {


            public string MATNR { get; set; }
            public int DZMENG { get; set; }

            public string SPART { get; set; }

        }
        private class IntegrationLog
        {
            public Guid hil_integrationlogid { get; set; }
            public string hil_name { get; set; }
            public int hil_totalrecords { get; set; }
            public int hil_recordsaffected { get; set; }
            public int hil_totalerrorrecords { get; set; }
            public bool hil_states { get; set; }
            public string hil_syncdate { get; set; }
        }

        private class ET_Details
        {
            public string VBELN { get; set; }
            public string POSNR { get; set; }
            public string MATNR { get; set; }
            public string MAKTX { get; set; }
            public string ERDAT { get; set; }
            public string KWMENG { get; set; }
            public string VRKME { get; set; }
            public string NETWR { get; set; }
            public string WAERK { get; set; }
            public string NETPR { get; set; }
            public string MWSBP { get; set; }
            public string H_LFSTK { get; set; }
            public string H_STATUS { get; set; }
            public string L_LFSTA { get; set; }
            public string L_STATUS { get; set; }
            public string KUNNR { get; set; }
            public string H_DEL_DATE { get; set; }
            public string L_DEL_DATE { get; set; }
            public string BSTNK { get; set; }
            public string IHREZ { get; set; }
            public string H_NETWR { get; set; }
            public string H_WAERK { get; set; }
            public string ST_KUNNR { get; set; }
            public string NAME1 { get; set; }
            public string NAME2 { get; set; }
            public string LAND1 { get; set; }
            public string ORT01 { get; set; }
            public string PSTLZ { get; set; }
            public string REGIO { get; set; }
            public string STRAS { get; set; }
            public string TELF1 { get; set; }
            public string TELFX { get; set; }
            public string REGSMTP_ADDRIO { get; set; }
        }
    }
}
