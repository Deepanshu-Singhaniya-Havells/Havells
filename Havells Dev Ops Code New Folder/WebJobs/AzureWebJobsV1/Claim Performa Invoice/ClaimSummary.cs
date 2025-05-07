using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OMSMail
{
    public class ClaimSummary
    {
        public class ClaimSummaryData
        {
            public string ClaimId { get; set; }
            public string ZMONTH { get; set; }
            public string SALES_OFFICE { get; set; }
            public string VENDOR_CODE { get; set; }
            public string ACTIVITY_NUMBER { get; set; }
            public string ACTIVITY_QTY { get; set; }
            public string Amount { get; set; }
            public string DIVISION { get; set; }
        }
        public class RootClaimSummary
        {
            public List<ClaimSummaryData> LT_TABLE { get; set; }
        }
        public void sendClaimSummaryDataAdhoc(IOrganizationService service)
        {
            try
            {
                Guid _performaInvoiceId = Guid.Empty;
                string _fetchXML = string.Empty;
                EntityReference _erFiscalMonth = null;

                QueryExpression queryExp = new QueryExpression("hil_claimperiod");
                queryExp.ColumnSet = new ColumnSet("hil_fromdate", "hil_todate");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition("hil_isapplicable", ConditionOperator.Equal, true);
                queryExp.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                EntityCollection entcoll = service.RetrieveMultiple(queryExp);
                if (entcoll.Entities.Count > 0)
                {
                    _erFiscalMonth = entcoll.Entities[0].ToEntityReference();
                }
                if (_erFiscalMonth == null)
                {
                    Console.WriteLine("Fiscal Month is not defined." + DateTime.Now.ToString());
                    return;
                }

                var ClaimId = new string[] { "CL-00063982" };
                for (int j = 0; j < ClaimId.Length; j++)
                {
                    QueryExpression query = new QueryExpression("hil_claimheader");
                    query.ColumnSet = new ColumnSet("hil_name");
                    //query.Distinct = true;
                    query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                    //query.Criteria.AddCondition("hil_performastatus", ConditionOperator.Equal, 3); // Approved by BSH
                    //query.Criteria.AddCondition("hil_syncstatus", ConditionOperator.Equal, false); // Pending to Sync
                    //query.Criteria.AddCondition("hil_fiscalmonth", ConditionOperator.Equal, _erFiscalMonth.Id);
                    query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, ClaimId[j]);
                    EntityCollection claimSummaryCol1 = service.RetrieveMultiple(query);

                    if (claimSummaryCol1.Entities.Count > 0)
                    {
                        int i = 1;
                        string message = string.Empty;
                        Entity cSummary;
                        foreach (Entity entLine in claimSummaryCol1.Entities)
                        {
                            message = string.Empty;

                            _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_claimsummary'>
                            <attribute name='hil_claimsummaryid' />
                            <attribute name='hil_name' />
                            <attribute name='createdon' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                                <filter type='or'>
                                    <condition attribute='hil_activitycode' operator='null' />
                                    <condition attribute='hil_claimamount' operator='eq' value='0' />
                                </filter>
                                <condition attribute='hil_performainvoiceid' operator='eq' value='{entLine.Id}' />
                            </filter>
                            </entity>
                            </fetch>";

                            EntityCollection claimSummaryCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (claimSummaryCol.Entities.Count > 0)
                            {
                                message = "Activity Code/Claim Amount is null : " + entLine.GetAttributeValue<string>("hil_name");
                                Console.WriteLine(message);
                                cSummary = new Entity("hil_claimheader");
                                cSummary.Id = entLine.Id;
                                cSummary["hil_syncstatus"] = false;
                                cSummary.Attributes["hil_syncresponse"] = message;
                                cSummary.Attributes["hil_syncdoneon"] = DateTime.Now;
                                service.Update(cSummary);
                                continue;
                            }
                            Decimal _claimAmount = 0, _claimSumAmount = 0;
                            _fetchXML = $@"<fetch distinct='false' mapping='logical' aggregate='true'>
                            <entity name='hil_claimline'>
                            <attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/> 
                            <filter type='and'>
                                <condition attribute='hil_claimheader' operator='eq' value='{entLine.Id}' />
                                <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            </entity>
                            </fetch>";
                            claimSummaryCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (claimSummaryCol.Entities.Count > 0)
                            {
                                if (claimSummaryCol.Entities[0].Contains("claimamount") && claimSummaryCol.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value != null)
                                    _claimAmount = ((Money)claimSummaryCol.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value).Value;
                            }
                            _fetchXML = $@"<fetch distinct='false' mapping='logical' aggregate='true'>
                            <entity name='hil_claimsummary'>
                            <attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/> 
                            <filter type='and'>
                                <condition attribute='hil_performainvoiceid' operator='eq' value='{entLine.Id}' />
                                <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            </entity>
                            </fetch>";
                            claimSummaryCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (claimSummaryCol.Entities.Count > 0)
                            {
                                if (claimSummaryCol.Entities[0].Contains("claimamount") && claimSummaryCol.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value != null)
                                    _claimSumAmount = ((Money)claimSummaryCol.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value).Value;
                            }
                            if (_claimAmount != _claimSumAmount)
                            {
                                message = "Claim Line and Summary amount is mismatch" + entLine.GetAttributeValue<string>("hil_name");
                                Console.WriteLine(message);
                                cSummary = new Entity("hil_claimheader");
                                cSummary.Id = entLine.Id;
                                cSummary["hil_syncstatus"] = false;
                                cSummary.Attributes["hil_syncresponse"] = message;
                                cSummary.Attributes["hil_syncdoneon"] = DateTime.Now;
                                service.Update(cSummary);
                                continue;
                            }
                            query = new QueryExpression("hil_claimsummary");
                            query.ColumnSet = new ColumnSet("hil_claimsummaryid", "hil_productdivision", "hil_name", "hil_performainvoiceid", "hil_franchise", "hil_franchisecode", "hil_salesoffice", "hil_salesofficecode", "hil_claimmonth", "hil_claimmonthcode", "hil_activitycode", "hil_qty", "hil_claimamount");
                            query.Distinct = true;
                            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            query.Criteria.AddCondition("hil_performainvoiceid", ConditionOperator.Equal, entLine.Id);
                            claimSummaryCol = service.RetrieveMultiple(query);

                            if (claimSummaryCol.Entities.Count > 0)
                            {
                                Entity lstpo2 = claimSummaryCol.Entities[0];
                                List<ClaimSummaryData> ClaimSummaryDataLst = new List<ClaimSummaryData>();
                                foreach (Entity entity in claimSummaryCol.Entities)
                                {
                                    ClaimSummaryData claimSummaryData = new ClaimSummaryData();
                                    claimSummaryData.ClaimId = entity.GetAttributeValue<EntityReference>("hil_performainvoiceid").Name;
                                    claimSummaryData.ZMONTH = entity.GetAttributeValue<string>("hil_claimmonthcode");
                                    claimSummaryData.SALES_OFFICE = entity.GetAttributeValue<string>("hil_salesofficecode");
                                    claimSummaryData.VENDOR_CODE = entity.GetAttributeValue<string>("hil_franchisecode");
                                    claimSummaryData.ACTIVITY_NUMBER = entity.GetAttributeValue<string>("hil_activitycode");
                                    claimSummaryData.ACTIVITY_QTY = Convert.ToString(entity.GetAttributeValue<Int32>("hil_qty"));
                                    claimSummaryData.Amount = String.Format("{0:0.00}", entity.GetAttributeValue<Money>("hil_claimamount").Value);
                                    claimSummaryData.DIVISION = entity.GetAttributeValue<string>("hil_productdivision");
                                    ClaimSummaryDataLst.Add(claimSummaryData);
                                }
                                RootClaimSummary rootClaimSummary = new RootClaimSummary();
                                rootClaimSummary.LT_TABLE = ClaimSummaryDataLst;
                                IntegrationConfig integrationConfig = new IntegrationConfig();
                                //integrationConfig.uri = "https://middlewaredev.havells.com:50001/RESTAdapter/dynamics/ClaimData?IM_PROJECT=D365";
                                //integrationConfig.uri = "https://middlewareqa.havells.com:50001/RESTAdapter/dynamics/ClaimData?IM_PROJECT=D365";
                                integrationConfig.uri = "https://p90ci.havells.com:50001/RESTAdapter/dynamics/ClaimData?IM_PROJECT=D365";

                                //integrationConfig.Auth = "D365_HAVELLS:DEVD365@1234";
                                //integrationConfig.Auth = "D365_Havells:QAD365@1234";
                                integrationConfig.Auth = "D365_Havells:PRDD365@1234";

                                string url = integrationConfig.uri;
                                string authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(integrationConfig.Auth));
                                string data = JsonConvert.SerializeObject(rootClaimSummary);
                                var client = new RestClient(url);
                                client.Timeout = -1;
                                var request = new RestRequest(Method.POST);

                                request.AddHeader("Authorization", "Basic " + authInfo);
                                request.AddHeader("Content-Type", "application/json");

                                Entity intigrationTrace = new Entity("hil_integrationtrace");
                                intigrationTrace["hil_entityname"] = lstpo2.LogicalName;
                                intigrationTrace["hil_entityid"] = lstpo2.Id.ToString();
                                intigrationTrace["hil_request"] = data;
                                intigrationTrace["hil_name"] = lstpo2.GetAttributeValue<string>("hil_name");
                                Guid intigrationTraceID = service.Create(intigrationTrace);

                                request.AddParameter("application/json", data, ParameterType.RequestBody);
                                IRestResponse response = client.Execute(request);
                                Console.WriteLine("ClaimId: " + lstpo2.GetAttributeValue<string>("hil_name") + " Response: " + response.Content);

                                //Entity intigrationTraceUp = new Entity("hil_integrationtrace");
                                //intigrationTraceUp["hil_response"] = response.Content == "" ? response.ErrorMessage : response.Content;
                                //intigrationTraceUp.Id = intigrationTraceID;
                                //service.Update(intigrationTraceUp);

                                try
                                {
                                    Console.WriteLine("Data Fetched");
                                    dynamic returndatas = JsonConvert.DeserializeObject<dynamic>(response.Content);
                                    string cln = returndatas["CLAIM_ID"].ToString();
                                    string EBELNO = returndatas["EBELN"].ToString();
                                    List<MSG> lstmsg = new List<MSG>();
                                    dynamic allMsg = returndatas["MSG"];

                                    bool syncstatus = false;
                                    if (allMsg.Count == null)
                                    {
                                        syncstatus = true;
                                    }
                                    else if (allMsg.Count > 0)
                                    {
                                        foreach (var item in allMsg)
                                        {
                                            MSG msg = new MSG();
                                            msg.ID = item.ID;
                                            msg.TYPE = item.TYPE;
                                            msg.NUMBER = item.NUMBER;
                                            msg.MESSAGE = item.MESSAGE;
                                            lstmsg.Add(msg);
                                            message = message + " " + item.MESSAGE;
                                        }
                                    }
                                    cSummary = new Entity("hil_claimheader");
                                    cSummary.Id = lstpo2.GetAttributeValue<EntityReference>("hil_performainvoiceid").Id;
                                    cSummary["hil_syncstatus"] = syncstatus;
                                    if (syncstatus)
                                    {
                                        cSummary["hil_performastatus"] = new OptionSetValue(4);//Posted
                                        cSummary.Attributes["hil_sapponumber"] = EBELNO; // SAP PO Number
                                    }
                                    cSummary.Attributes["hil_syncresponse"] = message;
                                    cSummary.Attributes["hil_syncdoneon"] = DateTime.Now;
                                    service.Update(cSummary);
                                }
                                catch
                                {
                                }
                            }
                            Console.WriteLine("Processing... " + i++.ToString() + "/" + claimSummaryCol1.Entities.Count);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public void sendClaimSummaryData(IOrganizationService service)
        {
            try
            {
                Guid _performaInvoiceId = Guid.Empty;
                string _fetchXML = string.Empty;

                QueryExpression query = new QueryExpression("hil_claimheader");
                query.ColumnSet = new ColumnSet("hil_name", "hil_claimheaderid");
                query.Distinct = true;
                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                query.Criteria.AddCondition("hil_performastatus", ConditionOperator.Equal, 3); // Approved by BSH
                query.Criteria.AddCondition("hil_syncstatus", ConditionOperator.Equal, false); // Pending to Sync
                query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "CL-00069563"); // Claim Performa Invoices
                EntityCollection claimSummaryCol1 = service.RetrieveMultiple(query);

                if (claimSummaryCol1.Entities.Count > 0)
                {
                    int i = 1;
                    string message = string.Empty;
                    Entity cSummary;
                    foreach (Entity entLine in claimSummaryCol1.Entities)
                    {
                        message = string.Empty;

                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_claimsummary'>
                            <attribute name='hil_claimsummaryid' />
                            <attribute name='hil_name' />
                            <attribute name='createdon' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                                <filter type='or'>
                                    <condition attribute='hil_activitycode' operator='null' />
                                    <condition attribute='hil_claimamount' operator='eq' value='0' />
                                </filter>
                                <condition attribute='hil_performainvoiceid' operator='eq' value='{entLine.Id}' />
                            </filter>
                            </entity>
                            </fetch>";

                        EntityCollection claimSummaryCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (claimSummaryCol.Entities.Count > 0)
                        {
                            message = "Activity Code/Claim Amount is null : " + entLine.GetAttributeValue<string>("hil_name");
                            Console.WriteLine(message);
                            cSummary = new Entity("hil_claimheader");
                            cSummary.Id = entLine.Id;
                            cSummary["hil_syncstatus"] = false;
                            cSummary.Attributes["hil_syncresponse"] = message;
                            cSummary.Attributes["hil_syncdoneon"] = DateTime.Now;
                            service.Update(cSummary);
                            continue;
                        }
                        Decimal _claimAmount = 0, _claimSumAmount = 0;
                        _fetchXML = $@"<fetch distinct='false' mapping='logical' aggregate='true'>
                            <entity name='hil_claimline'>
                            <attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/> 
                            <filter type='and'>
                                <condition attribute='hil_claimheader' operator='eq' value='{entLine.Id}' />
                                <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            </entity>
                            </fetch>";
                        claimSummaryCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (claimSummaryCol.Entities.Count > 0)
                        {
                            if (claimSummaryCol.Entities[0].Contains("claimamount") && claimSummaryCol.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value != null)
                                _claimAmount = ((Money)claimSummaryCol.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value).Value;
                        }
                        _fetchXML = $@"<fetch distinct='false' mapping='logical' aggregate='true'>
                            <entity name='hil_claimsummary'>
                            <attribute name='hil_claimamount' alias='claimamount' aggregate='sum'/> 
                            <filter type='and'>
                                <condition attribute='hil_performainvoiceid' operator='eq' value='{entLine.Id}' />
                                <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            </entity>
                            </fetch>";
                        claimSummaryCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (claimSummaryCol.Entities.Count > 0)
                        {
                            if (claimSummaryCol.Entities[0].Contains("claimamount") && claimSummaryCol.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value != null)
                                _claimSumAmount = ((Money)claimSummaryCol.Entities[0].GetAttributeValue<AliasedValue>("claimamount").Value).Value;
                        }
                        if (_claimAmount != _claimSumAmount)
                        {
                            message = "Claim Line and Summary amount is mismatch" + entLine.GetAttributeValue<string>("hil_name");
                            Console.WriteLine(message);
                            cSummary = new Entity("hil_claimheader");
                            cSummary.Id = entLine.Id;
                            cSummary["hil_syncstatus"] = false;
                            cSummary.Attributes["hil_syncresponse"] = message;
                            cSummary.Attributes["hil_syncdoneon"] = DateTime.Now;
                            service.Update(cSummary);
                            continue;
                        }
                        query = new QueryExpression("hil_claimsummary");
                        query.ColumnSet = new ColumnSet("hil_claimsummaryid", "hil_productdivision", "hil_name", "hil_performainvoiceid", "hil_franchise", "hil_franchisecode", "hil_salesoffice", "hil_salesofficecode", "hil_claimmonth", "hil_claimmonthcode", "hil_activitycode", "hil_qty", "hil_claimamount");
                        query.Distinct = true;
                        query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                        query.Criteria.AddCondition("hil_performainvoiceid", ConditionOperator.Equal, entLine.Id);
                        claimSummaryCol = service.RetrieveMultiple(query);

                        if (claimSummaryCol.Entities.Count > 0)
                        {
                            Entity lstpo2 = claimSummaryCol.Entities[0];
                            List<ClaimSummaryData> ClaimSummaryDataLst = new List<ClaimSummaryData>();
                            foreach (Entity entity in claimSummaryCol.Entities)
                            {
                                ClaimSummaryData claimSummaryData = new ClaimSummaryData();
                                claimSummaryData.ClaimId = entity.GetAttributeValue<EntityReference>("hil_performainvoiceid").Name;
                                claimSummaryData.ZMONTH = entity.GetAttributeValue<string>("hil_claimmonthcode");
                                claimSummaryData.SALES_OFFICE = entity.GetAttributeValue<string>("hil_salesofficecode");
                                claimSummaryData.VENDOR_CODE = entity.GetAttributeValue<string>("hil_franchisecode");
                                claimSummaryData.ACTIVITY_NUMBER = entity.GetAttributeValue<string>("hil_activitycode");
                                claimSummaryData.ACTIVITY_QTY = Convert.ToString(entity.GetAttributeValue<Int32>("hil_qty"));
                                claimSummaryData.Amount = String.Format("{0:0.00}", entity.GetAttributeValue<Money>("hil_claimamount").Value);
                                claimSummaryData.DIVISION = entity.GetAttributeValue<string>("hil_productdivision");
                                ClaimSummaryDataLst.Add(claimSummaryData);
                            }
                            RootClaimSummary rootClaimSummary = new RootClaimSummary();
                            rootClaimSummary.LT_TABLE = ClaimSummaryDataLst;
                            IntegrationConfig integrationConfig = new IntegrationConfig();
                            //integrationConfig.uri = "https://middlewaredev.havells.com:50001/RESTAdapter/dynamics/ClaimData?IM_PROJECT=D365";
                            //integrationConfig.uri = "https://middlewareqa.havells.com:50001/RESTAdapter/dynamics/ClaimData?IM_PROJECT=D365";
                            integrationConfig.uri = "https://p90ci.havells.com:50001/RESTAdapter/dynamics/ClaimData?IM_PROJECT=D365";
                            //integrationConfig.uri = "https://middleware.havells.com:50001/RESTAdapter/dynamics/ClaimData?IM_PROJECT=D365";

                            //integrationConfig.Auth = "D365_HAVELLS:DEVD365@1234";
                            //integrationConfig.Auth = "D365_Havells:QAD365@1234";
                            integrationConfig.Auth = "D365_Havells:PRDD365@1234";

                            string url = integrationConfig.uri;
                            string authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(integrationConfig.Auth));
                            string data = JsonConvert.SerializeObject(rootClaimSummary);
                            var client = new RestClient(url);
                            client.Timeout = -1;
                            var request = new RestRequest(Method.POST);

                            request.AddHeader("Authorization", "Basic " + authInfo);
                            request.AddHeader("Content-Type", "application/json");

                            Entity intigrationTrace = new Entity("hil_integrationtrace");
                            intigrationTrace["hil_entityname"] = lstpo2.LogicalName;
                            intigrationTrace["hil_entityid"] = lstpo2.Id.ToString();
                            intigrationTrace["hil_request"] = data;
                            intigrationTrace["hil_name"] = lstpo2.GetAttributeValue<string>("hil_name");
                            Guid intigrationTraceID = service.Create(intigrationTrace);

                            request.AddParameter("application/json", data, ParameterType.RequestBody);
                            IRestResponse response = client.Execute(request);
                            Console.WriteLine("ClaimId: " + lstpo2.GetAttributeValue<string>("hil_name") + " Response: " + response.Content);

                            //Entity intigrationTraceUp = new Entity("hil_integrationtrace");
                            //intigrationTraceUp["hil_response"] = response.Content == "" ? response.ErrorMessage : response.Content;
                            //intigrationTraceUp.Id = intigrationTraceID;
                            //service.Update(intigrationTraceUp);

                            try
                            {
                                Console.WriteLine("Data Fetched");
                                dynamic returndatas = JsonConvert.DeserializeObject<dynamic>(response.Content);
                                string cln = returndatas["CLAIM_ID"].ToString();
                                string EBELNO = returndatas["EBELN"].ToString();
                                List<MSG> lstmsg = new List<MSG>();
                                dynamic allMsg = returndatas["MSG"];

                                bool syncstatus = false;
                                if (allMsg.Count == null)
                                {
                                    syncstatus = true;
                                }
                                else if (allMsg.Count > 0)
                                {
                                    foreach (var item in allMsg)
                                    {
                                        MSG msg = new MSG();
                                        msg.ID = item.ID;
                                        msg.TYPE = item.TYPE;
                                        msg.NUMBER = item.NUMBER;
                                        msg.MESSAGE = item.MESSAGE;
                                        lstmsg.Add(msg);
                                        message = message + " " + item.MESSAGE;
                                    }
                                }
                                cSummary = new Entity("hil_claimheader");
                                cSummary.Id = lstpo2.GetAttributeValue<EntityReference>("hil_performainvoiceid").Id;
                                cSummary["hil_syncstatus"] = syncstatus;
                                if (syncstatus)
                                {
                                    cSummary["hil_performastatus"] = new OptionSetValue(4);//Posted
                                    cSummary.Attributes["hil_sapponumber"] = EBELNO; // SAP PO Number
                                }
                                cSummary.Attributes["hil_syncresponse"] = message;
                                cSummary.Attributes["hil_syncdoneon"] = DateTime.Now;
                                service.Update(cSummary);
                            }
                            catch
                            {
                            }
                        }
                        Console.WriteLine("Processing... " + i++.ToString() + "/" + claimSummaryCol1.Entities.Count);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public void getCliamIdResponsrByCliamId(IOrganizationService service)
        {
            QueryExpression query = new QueryExpression("hil_claimsummary");
            query.ColumnSet = new ColumnSet("hil_productdivision", "hil_name", "hil_performainvoiceid", "hil_franchise", "hil_franchisecode", "hil_salesoffice", "hil_salesofficecode", "hil_claimmonth", "hil_claimmonthcode", "hil_activitycode", "hil_qty", "hil_claimamount");
            query.Distinct = true;
            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            EntityCollection claimSummaryCol = service.RetrieveMultiple(query);
        }
    }
    public class IntegrationConfig
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
    public class Info
    {
        public string CLAIM_ID { get; set; }
        public string EBELNO { get; set; }
        public List<MSG> msgss { get; set; }

    }
    public class MSG
    {
        public string ID { get; set; }
        public string MESSAGE { get; set; }
        public int NUMBER { get; set; }
        public string TYPE { get; set; }
    }

}

