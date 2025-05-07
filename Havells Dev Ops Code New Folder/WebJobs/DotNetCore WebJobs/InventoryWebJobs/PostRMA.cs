using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Text;

namespace InventoryWebJobs
{
    internal class PostRMA
    {
        private readonly ServiceClient service;

        public PostRMA(ServiceClient _service)
        {
            this.service = _service;
        }

        private class Request
        {
            public string ZREF_D365 { get; set; }
            public string KUNNR { get; set; }
            public string RETURN_TYPE { get; set; }
            public string REF_DATE { get; set; }
            public List<RmaLine> LT_TABLE;
        }

        private class RmaLine
        {
            public string MATNR { get; set; }
            public int RNQTY { get; set; }
        }

        private class Response
        {
            public string ZREF_D365 { get; set; }
            public string INSPNO { get; set; }
            public string ERROR { get; set; }
            public string STATUS { get; set; }
        }
        internal async Task PostRMAs()
        {
            QueryExpression queryExpression = new QueryExpression("hil_inventoryrma");
            queryExpression.ColumnSet = new ColumnSet("hil_name", "hil_franchise", "hil_returntype", "hil_responsefromsap", "hil_syncedwithsap", "hil_inspectionnumber", "hil_rmastatus");
            queryExpression.Criteria.AddCondition("hil_rmastatus", ConditionOperator.Equal, 7); // Submit for inspection
            queryExpression.Criteria.AddCondition("hil_syncedwithsap", ConditionOperator.Equal, false);
            //queryExpression.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "DC2024005048");
            //queryExpression.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "DC2024001684");
             EntityCollection rmaCollection = service.RetrieveMultiple(queryExpression);

            foreach (Entity inventoryRMA in rmaCollection.Entities)
            {
                IntegrationConfiguration creds = HelperClass.GetIntegrationConfiguration(service, "Inventory_RMA");
                var planeBytes = Encoding.Default.GetBytes(creds.userName + ":" + creds.password);
                string auth = Convert.ToBase64String(planeBytes);

                Request request = new Request();
                request.REF_DATE = DateTime.Now.Date.ToString("yyyy-MM-dd").Replace("-", "");
                request.ZREF_D365 = inventoryRMA.Contains("hil_name") ? inventoryRMA.GetAttributeValue<string>("hil_name") : string.Empty;

                EntityReference franchiseRef = inventoryRMA.GetAttributeValue<EntityReference>("hil_franchise");
                Entity franchise = service.Retrieve("account", franchiseRef.Id, new ColumnSet("accountnumber", "customertypecode"));

                int CustomerType = franchise.GetAttributeValue<OptionSetValue>("customertypecode").Value;

                if (CustomerType == 6)//Franchisee
                    request.KUNNR = "F" + franchise.GetAttributeValue<string>("accountnumber");
                else
                    request.KUNNR = franchise.GetAttributeValue<string>("accountnumber");

                Console.WriteLine("Franchise Code: " + request.KUNNR);

                EntityReference rmaTypeRef = inventoryRMA.GetAttributeValue<EntityReference>("hil_returntype");
                Entity rmaType = service.Retrieve(rmaTypeRef.LogicalName, rmaTypeRef.Id, new ColumnSet("hil_sapcode"));
                request.RETURN_TYPE = rmaType.GetAttributeValue<string>("hil_sapcode");

                string fetchXml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' aggregate='true'>
                <entity name='hil_inventoryproductjournal'>
                <attribute name='hil_quantity' alias='Quantity' aggregate='sum'/>
                <attribute name='hil_partcode' alias='PartCode' groupby='true'  />
                <filter type='and'>
                    <condition attribute='hil_rma' operator='eq' value='{inventoryRMA.Id}' />
                    <condition attribute='hil_isused' operator='eq' value='1' />
                    <condition attribute='statecode' operator='eq' value='0' />
                </filter>
                </entity>
                </fetch>";

                EntityCollection FetchXmlCollection = service.RetrieveMultiple(new FetchExpression(fetchXml));

                List<RmaLine> tempList = new List<RmaLine>();
                for (int i = 0; i < FetchXmlCollection.Entities.Count; i++)
                {
                    RmaLine tempLine = new RmaLine();
                    EntityReference productRef = (EntityReference)FetchXmlCollection.Entities[i].GetAttributeValue<AliasedValue>("PartCode").Value;
                    Entity product = service.Retrieve(productRef.LogicalName, productRef.Id, new ColumnSet("productnumber"));
                    tempLine.MATNR = product.Contains("productnumber") ? product.GetAttributeValue<string>("productnumber") : string.Empty;
                    tempLine.RNQTY = (int)FetchXmlCollection.Entities[i].GetAttributeValue<AliasedValue>("Quantity").Value;
                    tempList.Add(tempLine);
                }

                request.LT_TABLE = tempList;

                var client = new HttpClient();
                var apiRequest = new HttpRequestMessage(HttpMethod.Post, creds.url);

                apiRequest.Headers.Add("Authorization", "Basic " + auth);
                HttpContent content = new StringContent(JsonConvert.SerializeObject(request));
                apiRequest.Content = content;
                var response = await client.SendAsync(apiRequest);
                response.EnsureSuccessStatusCode();
                Response res = JsonConvert.DeserializeObject<Response>(await response.Content.ReadAsStringAsync());

                if (res != null)
                {
                    inventoryRMA["hil_responsefromsap"] = res.ERROR;
                    if (res.STATUS == "S")
                    {
                        inventoryRMA["hil_syncedwithsap"] = true;
                        inventoryRMA["hil_responsefromsap"] = "Success";
                    }
                    if (!string.IsNullOrEmpty(res.INSPNO)) inventoryRMA["hil_inspectionnumber"] = res.INSPNO;
                    service.Update(inventoryRMA);
                    UpdateInventoryJournal(inventoryRMA.Id);
                }
            }
        }
        internal bool UpdateInventoryJournal(Guid inventoryRMAId)
        {
            bool _statusCode = true;
            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                       <entity name='hil_inventoryproductjournal'>
                         <attribute name='hil_inventoryproductjournalid' />
                         <attribute name='hil_name' />
                         <attribute name='createdon' />
                         <order attribute='hil_name' descending='false' />
                         <filter type='and'>
                           <condition attribute='statecode' operator='eq' value='0' />
                           <condition attribute='hil_transactiontype' operator='eq' value='3' />
                           <condition attribute='hil_rma' operator='eq' value='{inventoryRMAId}' />
                          <condition attribute='hil_isused' operator='ne' value='1' />
                        </filter>
                       </entity>
                     </fetch>";
            EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (entCol.Entities.Count > 0)
            {
                ExecuteMultipleRequest requestWithResults = new ExecuteMultipleRequest()
                {
                    // Assign settings that define execution behavior: continue on error, return responses. 
                    Settings = new ExecuteMultipleSettings()
                    {
                        ContinueOnError = false,
                        ReturnResponses = true
                    },
                    // Create an empty organization request collection.
                    Requests = new OrganizationRequestCollection()
                };
                Entity InventoryJournal = null;
                foreach (var entity in entCol.Entities)
                {
                    InventoryJournal = new Entity("hil_inventoryproductjournal", entity.Id);
                    InventoryJournal["hil_rma"] = null;
                    UpdateRequest updateRequest = new UpdateRequest() { Target = InventoryJournal };
                    requestWithResults.Requests.Add(updateRequest);
                }
                try
                {
                    ExecuteMultipleResponse responseWithResults = (ExecuteMultipleResponse)service.Execute(requestWithResults);
                }
                catch (Exception ex)
                {
                    _statusCode = false;
                }
            }
            return _statusCode;
        }
        //internal async Task PostRMAs_OLD()
        //{
        //    QueryExpression queryExpression = new QueryExpression("hil_inventoryrma");
        //    queryExpression.ColumnSet = new ColumnSet("hil_name", "hil_franchise", "hil_returntype", "hil_responsefromsap", "hil_syncedwithsap", "hil_inspectionnumber", "hil_rmastatus");
        //    queryExpression.Criteria.AddCondition("hil_rmastatus", ConditionOperator.Equal, 7); // Submit for inspection
        //    queryExpression.Criteria.AddCondition("hil_syncedwithsap", ConditionOperator.Equal, false);
        //    EntityCollection rmaCollection = service.RetrieveMultiple(queryExpression);


        //    foreach (Entity inventoryRMA in rmaCollection.Entities)
        //    {

        //        IntegrationConfiguration creds = HelperClass.GetIntegrationConfiguration(service, "Inventory_RMA");
        //        var planeBytes = Encoding.Default.GetBytes(creds.userName + ":" + creds.password);
        //        string auth = Convert.ToBase64String(planeBytes);


        //        Request request = new Request();
        //        request.REF_DATE = DateTime.Now.Date.ToString("yyyy-MM-dd").Replace("-", "");
        //        request.ZREF_D365 = inventoryRMA.Contains("hil_name") ? inventoryRMA.GetAttributeValue<string>("hil_name") : string.Empty;


        //        EntityReference franchiseRef = inventoryRMA.GetAttributeValue<EntityReference>("hil_franchise");
        //        Entity franchise = service.Retrieve("account", franchiseRef.Id, new ColumnSet("accountnumber"));
        //        request.KUNNR = "F" + franchise.GetAttributeValue<string>("accountnumber");
        //        Console.WriteLine("Franchise Code: " + request.KUNNR);

        //        EntityReference rmaTypeRef = inventoryRMA.GetAttributeValue<EntityReference>("hil_returntype");
        //        Entity rmaType = service.Retrieve(rmaTypeRef.LogicalName, rmaTypeRef.Id, new ColumnSet("hil_sapcode"));
        //        request.RETURN_TYPE = rmaType.GetAttributeValue<string>("hil_sapcode");

        //        QueryExpression query = new QueryExpression("hil_inventoryrmaline");
        //        query.ColumnSet = new ColumnSet("hil_product", "hil_quantity");
        //        query.Criteria.AddCondition("hil_rma", ConditionOperator.Equal, inventoryRMA.Id);
        //        query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 1);
        //        EntityCollection linesCollection = service.RetrieveMultiple(query);
        //        List<RmaLine> tempList = new List<RmaLine>();

        //        for (int i = 0; i < linesCollection.Entities.Count; i++)
        //        {
        //            RmaLine tempLine = new RmaLine();
        //            EntityReference productRef = linesCollection.Entities[i].GetAttributeValue<EntityReference>("hil_product");
        //            Entity product = service.Retrieve(productRef.LogicalName, productRef.Id, new ColumnSet("productnumber"));
        //            tempLine.MATNR = product.Contains("productnumber") ? product.GetAttributeValue<string>("productnumber") : string.Empty;
        //            tempLine.RNQTY = linesCollection.Entities[i].Contains("hil_quantity") ? linesCollection.Entities[i].GetAttributeValue<int>("hil_quantity") : 0;
        //            tempList.Add(tempLine);
        //        }
        //        request.LT_TABLE = tempList;

        //        var client = new HttpClient();
        //        var apiRequest = new HttpRequestMessage(HttpMethod.Post, creds.url);

        //        apiRequest.Headers.Add("Authorization", "Basic " + auth);
        //        HttpContent content = new StringContent(JsonConvert.SerializeObject(request));
        //        apiRequest.Content = content;
        //        var response = await client.SendAsync(apiRequest);
        //        response.EnsureSuccessStatusCode();
        //        Response res = JsonConvert.DeserializeObject<Response>(await response.Content.ReadAsStringAsync());

        //        if (res != null)
        //        {
        //            inventoryRMA["hil_responsefromsap"] = res.ERROR;
        //            if (res.STATUS == "S")
        //            {
        //                inventoryRMA["hil_syncedwithsap"] = true;
        //                inventoryRMA["hil_responsefromsap"] = "Success";
        //            }
        //            if (!string.IsNullOrEmpty(res.INSPNO)) inventoryRMA["hil_inspectionnumber"] = res.INSPNO;
        //            service.Update(inventoryRMA);
        //        }
        //    }

        //}

    }
}
