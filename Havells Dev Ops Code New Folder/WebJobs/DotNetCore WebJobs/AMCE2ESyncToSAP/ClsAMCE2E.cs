using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Configuration;
using System.Net;
using System.Text;

namespace AMCE2ESyncToSAP
{
    public class ClsAMCE2E
    {
        private static Guid PriceLevelForFGsale = new Guid("c05e7837-8fc4-ed11-b597-6045bdac5e78");//AMC Ominichannnel
        private readonly ServiceClient _service;
        public ClsAMCE2E(ServiceClient service)
        {
            _service = service;
        }
        public void getPaymentStatusFromPayU()
        {
            try
            {
                List<TransactionDetails> lstTransactionDetails = new List<TransactionDetails>();

                //List<TransactionDetails> lstJobTransactionDetails = getJobsForPaymentStatus();
                //lstTransactionDetails.AddRange(lstJobTransactionDetails);

                List<TransactionDetails> lstInvoiceTransactionDetails = getInvoiceForPaymentStatus();
                lstTransactionDetails.AddRange(lstInvoiceTransactionDetails);

                string payStatus = string.Empty;
                IntegrationConfig intConfig = IntegrationConfiguration("Send Payment Link");
                string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));
                int i = 1;
                foreach (TransactionDetails transactionDetails in lstTransactionDetails)
                {
                    string transactionID = transactionDetails.transactionID;
                    EntityReference RegardingEntity = transactionDetails.Regarding;

                    StatusRequest req = new StatusRequest();
                    req.PROJECT = "D365";
                    req.command = "verify_payment";
                    req.var1 = transactionID;
                    using (HttpClient client = new HttpClient())
                    {
                        var data = new StringContent(JsonConvert.SerializeObject(req), Encoding.UTF8, "application/json");
                        client.DefaultRequestHeaders.Add("Authorization", authInfo);
                        HttpResponseMessage response = client.PostAsync(intConfig.uri, data).Result;
                        if (response.IsSuccessStatusCode)
                        {
                            var obj = JsonConvert.DeserializeObject<StatusResponse>(response.Content.ReadAsStringAsync().Result);

                            QueryExpression Query = new QueryExpression("hil_paymentstatus");
                            Query.ColumnSet = new ColumnSet("hil_job", "hil_name");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, transactionID);
                            EntityCollection Found = _service.RetrieveMultiple(Query);
                            var item = obj.transaction_details[0];
                            Entity statusPayment = new Entity("hil_paymentstatus");
                            statusPayment.Id = Found[0].Id;
                            statusPayment["hil_mihpayid"] = item.mihpayid;
                            statusPayment["hil_request_id"] = item.request_id;
                            statusPayment["hil_bank_ref_num"] = item.bank_ref_num;

                            statusPayment["hil_amt"] = item.amt;
                            statusPayment["hil_transaction_amount"] = item.transaction_amount;
                            statusPayment["hil_additional_charges"] = item.additional_charges;
                            statusPayment["hil_productinfo"] = item.productinfo;
                            statusPayment["hil_firstname"] = item.firstname;
                            statusPayment["hil_bankcode"] = item.bankcode;
                            statusPayment["hil_udf1"] = item.udf1;
                            statusPayment["hil_udf2"] = item.udf2;
                            statusPayment["hil_udf3"] = item.udf3;
                            statusPayment["hil_udf4"] = item.udf4;
                            statusPayment["hil_udf5"] = item.udf5;
                            statusPayment["hil_field2"] = item.field2;
                            statusPayment["hil_field9"] = item.field9;
                            statusPayment["hil_error_code"] = item.error_code;
                            statusPayment["hil_addedon"] = item.addedon;
                            statusPayment["hil_payment_source"] = item.payment_source;
                            statusPayment["hil_card_type"] = item.card_type;
                            statusPayment["hil_error_message"] = item.error_Message;
                            statusPayment["hil_net_amount_debit"] = item.net_amount_debit;
                            statusPayment["hil_disc"] = item.disc;
                            statusPayment["hil_mode"] = item.mode;
                            statusPayment["hil_pg_type"] = item.PG_TYPE;
                            statusPayment["hil_card_no"] = item.card_no;
                            statusPayment["hil_paymentstatus"] = item.status;
                            statusPayment["hil_unmappedstatus"] = item.unmappedstatus;
                            statusPayment["hil_merchant_utr"] = item.Merchant_UTR;
                            statusPayment["hil_settled_at"] = item.Settled_At;
                            _service.Update(statusPayment);

                            string StatusPay = item.status.ToLower();
                            Console.WriteLine($"{i++}/{lstTransactionDetails.Count} {StatusPay} Payment Status of transactionid " + transactionID);

                            if (Found[0].Contains("hil_job") && RegardingEntity.LogicalName == "msdyn_workorder")
                            {
                                int paymentstatus = 0;
                                if (StatusPay == "not found")
                                {
                                    paymentstatus = (1);
                                }
                                else if (StatusPay == "success")
                                {
                                    paymentstatus = (2);
                                }
                                else if (StatusPay == "pending")
                                {
                                    paymentstatus = (3);
                                }
                                else
                                    paymentstatus = (4);

                                #region Updating Job Payment Status field && bank_ref_num
                                Entity update_RegardingEntity = new Entity(RegardingEntity.LogicalName, Found[0].GetAttributeValue<EntityReference>("hil_job").Id);
                                update_RegardingEntity["hil_paymentstatus"] = new OptionSetValue(paymentstatus);
                                update_RegardingEntity["hil_receiptnumber"] = item.bank_ref_num;
                                _service.Update(update_RegardingEntity);
                                #endregion
                            }
                            else if (RegardingEntity.LogicalName == "invoice")
                            {
                                if (StatusPay.ToLower() == "success".ToLower())
                                {
                                    Entity update_RegardingEntity = new Entity(RegardingEntity.LogicalName, RegardingEntity.Id);
                                    update_RegardingEntity["statecode"] = new OptionSetValue(2);
                                    update_RegardingEntity["statuscode"] = new OptionSetValue(100001);
                                    update_RegardingEntity["hil_bankrefnumber"] = item.bank_ref_num;
                                    _service.Update(update_RegardingEntity);

                                    Query = new QueryExpression("msdyn_paymentdetail");
                                    Query.ColumnSet = new ColumnSet(false);
                                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                                    Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, transactionID);
                                    EntityCollection FoundPaymentDetails = _service.RetrieveMultiple(Query);

                                    Entity update_paymentdetail = new Entity(FoundPaymentDetails[0].LogicalName, FoundPaymentDetails[0].Id);
                                    update_paymentdetail["statecode"] = new OptionSetValue(0);
                                    update_paymentdetail["statuscode"] = new OptionSetValue(910590000);
                                    _service.Update(update_paymentdetail);
                                }
                                else if (StatusPay.ToLower() == "failure".ToLower())
                                {
                                    Query = new QueryExpression("msdyn_paymentdetail");
                                    Query.ColumnSet = new ColumnSet(false);
                                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                                    Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, transactionID);
                                    EntityCollection FoundPaymentDetails = _service.RetrieveMultiple(Query);

                                    Entity update_msdyn_paymentdetail = new Entity(FoundPaymentDetails[0].LogicalName, FoundPaymentDetails[0].Id);
                                    update_msdyn_paymentdetail["msdyn_paymentamount"] = new Money(0);
                                    update_msdyn_paymentdetail["statecode"] = new OptionSetValue(0);
                                    update_msdyn_paymentdetail["statuscode"] = new OptionSetValue(910590001);
                                    _service.Update(update_msdyn_paymentdetail);
                                }
                                else if (StatusPay.ToLower() == "not found".ToLower())
                                {
                                    Query = new QueryExpression("msdyn_paymentdetail");
                                    Query.ColumnSet = new ColumnSet(false);
                                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                                    Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, transactionID);
                                }
                            }
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException("Something went wrong in the API. Error Code : " + (int)response.StatusCode);
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("D365 Internal Error " + ex.Message);
            }
        }
        public void PushAMCDataJob()
        {
            int totalRecords = 0;
            int recordCnt = 0;
            string prodCode = string.Empty;
            string jobId = string.Empty;

            try
            {
                LinkEntity lnkEntOwner = new LinkEntity
                {
                    LinkFromEntityName = "msdyn_workorder",
                    LinkToEntityName = "systemuser",
                    LinkFromAttributeName = "owninguser",
                    LinkToAttributeName = "systemuserid",
                    Columns = new ColumnSet("hil_employeecode"),
                    EntityAlias = "user",
                    JoinOperator = JoinOperator.Inner
                };
                LinkEntity lnkEntAddress = new LinkEntity
                {
                    LinkFromEntityName = "msdyn_workorder",
                    LinkToEntityName = "hil_address",
                    LinkFromAttributeName = "hil_address",
                    LinkToAttributeName = "hil_addressid",
                    Columns = new ColumnSet("hil_street1", "hil_street2", "hil_street3", "hil_state", "hil_pincode", "hil_district", "hil_city"),
                    EntityAlias = "address",
                    JoinOperator = JoinOperator.Inner
                };
                LinkEntity lnkEntConsumer = new LinkEntity
                {
                    LinkFromEntityName = "msdyn_workorder",
                    LinkToEntityName = "contact",
                    LinkFromAttributeName = "hil_customerref",
                    LinkToAttributeName = "contactid",
                    Columns = new ColumnSet("hil_salutation", "fullname", "emailaddress1"),
                    EntityAlias = "contact",
                    JoinOperator = JoinOperator.Inner
                };
                LinkEntity linkEntityForCustomerassets = new LinkEntity
                {
                    LinkFromEntityName = "msdyn_workorder",
                    LinkToEntityName = "msdyn_customerasset",
                    LinkFromAttributeName = "msdyn_customerasset",
                    LinkToAttributeName = "msdyn_customerassetid",
                    Columns = new ColumnSet("hil_invoicedate", "msdyn_product", "msdyn_name", "hil_warrantystatus", "hil_warrantytilldate"),
                    EntityAlias = "custAsset",
                    JoinOperator = JoinOperator.Inner
                };
                LinkEntity linkEntityForPaymentStatus = new LinkEntity
                {
                    LinkFromEntityName = "msdyn_workorder",
                    LinkToEntityName = "hil_paymentstatus",
                    LinkFromAttributeName = "msdyn_workorderid",
                    LinkToAttributeName = "hil_job",
                    Columns = new ColumnSet("hil_bank_ref_num"),
                    EntityAlias = "payment",
                    JoinOperator = JoinOperator.Inner
                };
                linkEntityForPaymentStatus.LinkCriteria.Conditions.Add(new ConditionExpression("hil_paymentstatus", ConditionOperator.Equal, "success"));
                QueryExpression Query = new QueryExpression("msdyn_workorder");
                Query.ColumnSet = new ColumnSet("msdyn_name", "hil_mobilenumber", "msdyn_customerasset", "ownerid", "hil_branch", "hil_receiptamount", "msdyn_timeclosed",
                    "createdon", "hil_actualcharges");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
          //      Query.Criteria.AddCondition(new ConditionExpression("hil_amcsyncstatus", ConditionOperator.Equal, 2)); //Pending Submission
                Query.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrAfter, new DateTime(2024, 11, 4))); //
                Query.Criteria.AddCondition(new ConditionExpression("hil_callsubtype", ConditionOperator.Equal, new Guid("55A71A52-3C0B-E911-A94E-000D3AF06CD4"))); //AMC Call
                Query.Criteria.AddCondition(new ConditionExpression("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"))); //Closed
                Query.Criteria.AddCondition(new ConditionExpression("hil_receiptamount", ConditionOperator.GreaterThan, 0)); //Receipt amount must be grater than 0.
                //Query.Criteria.AddCondition(new ConditionExpression("msdyn_name", ConditionOperator.Equal, "25072432717118")); 

                Query.AddOrder("createdon", OrderType.Ascending);
                Query.LinkEntities.Add(lnkEntOwner);
                Query.LinkEntities.Add(lnkEntAddress);
                Query.LinkEntities.Add(lnkEntConsumer);
                Query.LinkEntities.Add(linkEntityForCustomerassets);
                Query.LinkEntities.Add(linkEntityForPaymentStatus);
                EntityCollection ec = _service.RetrieveMultiple(Query);
                Console.WriteLine("Sync Started:");
                if (ec.Entities.Count > 0)
                {
                    totalRecords = ec.Entities.Count;
                    foreach (Entity ent in ec.Entities)
                    {
                        string fetchxml = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
                                          <entity name='hil_jobsextension'>
                                            <attribute name='hil_jobsextensionid' />
                                            <attribute name='hil_name' />
                                            <attribute name='createdon' />
                                            <order attribute='createdon' descending='false' />
                                            <filter type='and'>
                                              <condition attribute='hil_jobs' operator='eq' value='{ent.Id}' />
                                            </filter>
                                          </entity>
                                        </fetch>";

                        EntityCollection jobextcoll = _service.RetrieveMultiple(Query);

                        recordCnt += 1;
                        prodCode = GetMaterialCode(ent.Id, _service);
                        jobId = ent.GetAttributeValue<string>("msdyn_name");

                        if (prodCode != null)
                        {
                            string _AMCPlan = prodCode.Split('|')[1]; //AMC Plan GUID
                            prodCode = prodCode.Split('|')[0]; //AMC Plan Code

                            AMCInvoiceRequest requestData = new AMCInvoiceRequest();
                            requestData.LT_TABLE = new List<AMCInvoiceData>();
                            AMCInvoiceData invoiceData = new AMCInvoiceData();
                            if (ent.Attributes.Contains("contact.fullname"))
                            {
                                invoiceData.CUSTOMER_NAME = ent.GetAttributeValue<AliasedValue>("contact.fullname").Value.ToString();
                            }
                            if (ent.Attributes.Contains("contact.hil_salutation"))
                            {
                                invoiceData.TITLE = "Mr.";
                            }
                            if (ent.Attributes.Contains("address.hil_street1"))
                            {
                                invoiceData.STREET = ent.GetAttributeValue<AliasedValue>("address.hil_street1").Value.ToString();
                            }
                            if (ent.Attributes.Contains("address.hil_street2"))
                            {
                                invoiceData.HOUSE_NO = ent.GetAttributeValue<AliasedValue>("address.hil_street2").Value.ToString();
                            }
                            if (ent.Attributes.Contains("address.hil_street3"))
                            {
                                invoiceData.STREET4 = ent.GetAttributeValue<AliasedValue>("address.hil_street3").Value.ToString();
                            }
                            if (ent.Attributes.Contains("address.hil_pincode"))
                            {
                                invoiceData.POSTAL_CODE = ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_pincode").Value).Name;
                            }
                            if (ent.Attributes.Contains("address.hil_district"))
                            {
                                invoiceData.DISTRICT = ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_district").Value).Name;
                            }
                            if (ent.Attributes.Contains("address.hil_city"))
                            {
                                invoiceData.CITY = ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_city").Value).Name;
                            }
                            if (ent.Attributes.Contains("address.hil_state"))
                            {
                                Entity entTemp = _service.Retrieve("hil_state", ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_state").Value).Id, new ColumnSet("hil_sapstatecode"));
                                if (entTemp.Attributes.Contains("hil_sapstatecode"))
                                {
                                    invoiceData.REGION_CODE = entTemp.GetAttributeValue<string>("hil_sapstatecode").ToString();
                                }
                            }
                            invoiceData.CALL_ID = ent.GetAttributeValue<string>("msdyn_name");
                            if (ent.Attributes.Contains("contact.emailaddress1"))
                            {
                                invoiceData.EMAIL = ent.GetAttributeValue<AliasedValue>("contact.emailaddress1").Value.ToString();
                            }
                            if (ent.Attributes.Contains("hil_mobilenumber"))
                            {
                                invoiceData.PHONE = ent.GetAttributeValue<string>("hil_mobilenumber");
                            }
                            if (ent.Attributes.Contains("hil_branch"))
                            {
                                invoiceData.BRANCH_NAME = ent.GetAttributeValue<EntityReference>("hil_branch").Name;
                                Entity _entBranch = _service.Retrieve(ent.GetAttributeValue<EntityReference>("hil_branch").LogicalName, ent.GetAttributeValue<EntityReference>("hil_branch").Id, new ColumnSet("hil_mamorandumcode"));
                                if (_entBranch != null)
                                {
                                    invoiceData.SHIP_TO_PARTY = _entBranch.GetAttributeValue<string>("hil_mamorandumcode");
                                    invoiceData.SOLD_TO_PARTY = _entBranch.GetAttributeValue<string>("hil_mamorandumcode");
                                    invoiceData.SHIP_TO_PA_NAME = ent.GetAttributeValue<EntityReference>("hil_branch").Name;
                                }
                            }
                            if (ent.Attributes.Contains("payment.hil_bank_ref_num"))
                            {
                                invoiceData.PAYMENT_REF_NO = ent.GetAttributeValue<AliasedValue>("payment.hil_bank_ref_num").Value.ToString();
                            }
                            invoiceData.REMARKS = "AMC Sale";

                            invoiceData.MATERIAL_CODE = prodCode;

                            invoiceData.CS_SERVICE_PRICE = ent.GetAttributeValue<Money>("hil_receiptamount").Value.ToString();

                            if (ent.Attributes.Contains("custAsset.msdyn_name"))
                            {
                                invoiceData.PRODUCT_SERIAL_NO = ent.GetAttributeValue<AliasedValue>("custAsset.msdyn_name").Value.ToString();
                            }
                            if (ent.Attributes.Contains("custAsset.hil_invoicedate"))
                            {
                                string dopValue = Convert.ToDateTime(ent.GetAttributeValue<AliasedValue>("custAsset.hil_invoicedate").Value).ToString("yyyy-MM-dd");
                                invoiceData.PRODUCT_DOP = dopValue;
                            }
                            if (ent.Attributes.Contains("user.hil_employeecode"))
                            {
                                string _campaignCode = ent.GetAttributeValue<AliasedValue>("user.hil_employeecode").Value.ToString();
                                string[] _campaignCodeArr = _campaignCode.Split('|');
                                if (_campaignCodeArr.Length >= 1)
                                    invoiceData.EMPLOYEE_ID = _campaignCodeArr[0].ToUpper();
                                if (_campaignCodeArr.Length > 1)
                                    invoiceData.EMP_TYPE = _campaignCodeArr[1].ToUpper();
                            }

                            if (ent.Attributes.Contains("custAsset.hil_warrantytilldate"))
                            {
                                string dopValue = Convert.ToDateTime(ent.GetAttributeValue<AliasedValue>("custAsset.hil_warrantytilldate").Value).ToString("yyyy-MM-dd");
                                invoiceData.WARNTY_TILL_DATE = dopValue;
                            }

                            if (ent.Attributes.Contains("msdyn_timeclosed"))
                            {
                                string jobclosedonValue = ent.GetAttributeValue<DateTime>("msdyn_timeclosed").ToString("yyyy-MM-dd");
                                invoiceData.CLOSED_ON = jobclosedonValue;
                            }
                            if (ent.Attributes.Contains("custAsset.hil_warrantystatus"))
                            {
                                OptionSetValue optValue = ((OptionSetValue)ent.GetAttributeValue<AliasedValue>("custAsset.hil_warrantystatus").Value);
                                invoiceData.WARNTY_STATUS = optValue.Value == 1 ? "IN" : "OUT";
                            }
                            else
                            {
                                invoiceData.WARNTY_STATUS = "OUT";
                            }

                            string pricingDateValue = string.Empty;
                            pricingDateValue = ent.GetAttributeValue<DateTime>("msdyn_timeclosed").ToString("yyyy-MM-dd");
                            invoiceData.PRICING_DATE = pricingDateValue;

                            invoiceData.START_DATE = GetWarrantyStartDate(ent.GetAttributeValue<EntityReference>("msdyn_customerasset").Id, ent.GetAttributeValue<DateTime>("msdyn_timeclosed"));
                            if (!string.IsNullOrEmpty(invoiceData.START_DATE) && invoiceData.START_DATE.Length > 0)
                            {
                                invoiceData.END_DATE = GetWarrantyEndDate(new Guid(_AMCPlan), Convert.ToDateTime(invoiceData.START_DATE));
                            }
                            else
                            {
                                if (jobextcoll.Entities.Count > 0)
                                {
                                    Entity JobExtEnt = new Entity(jobextcoll.Entities[0].LogicalName, jobextcoll.Entities[0].Id);
                                    JobExtEnt["hil_amcsyncstatusmessage"] = "Warranty Template Setup is not updated properly.";
                                    _service.Update(JobExtEnt);
                                }
                                Console.WriteLine("ERROR!!! JobId:" + invoiceData.CALL_ID + " Warranty Template Setup is not updated properly.");
                                continue;
                            }

                            decimal _receiptAmt = ent.GetAttributeValue<Money>("hil_receiptamount").Value;
                            decimal _payableAmt = 0;

                            if (ent.Attributes.Contains("hil_actualcharges"))
                            {
                                _payableAmt = ent.GetAttributeValue<Money>("hil_actualcharges").Value;
                            }
                            if (_payableAmt < _receiptAmt)
                            {
                                if (jobextcoll.Entities.Count > 0)
                                {
                                    Entity JobExtEnt = new Entity(jobextcoll.Entities[0].LogicalName, jobextcoll.Entities[0].Id);
                                    JobExtEnt["hil_amcsyncstatusmessage"] = "Payable " + _payableAmt.ToString() + " and Receipt Amount " + _receiptAmt.ToString() + " is mismatch.";
                                    _service.Update(JobExtEnt);
                                }
                                Console.WriteLine("Payable " + _payableAmt.ToString() + " and Receipt Amount " + _receiptAmt.ToString() + " is mismatch. Pending " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                                continue;
                            }

                            invoiceData.ZDS6_AMOUNT = (_payableAmt - _receiptAmt).ToString();
                            invoiceData.DISCOUNT_FLAG = "X";
                            invoiceData.ITEM_NO = "1";
                            requestData.LT_TABLE.Add(invoiceData);

                            var Json = JsonConvert.SerializeObject(requestData);
                            IntegrationConfig objconfig = IntegrationConfiguration("amc_data_v1"); //Get Url and Auth 
                            objconfig.uri = "https://middlewareqa.havells.com:50001/RESTAdapter/dynamics/v1/amc_data";
                            objconfig.Auth = "D365_Havells:QAD365@1234";
                            WebRequest request = WebRequest.Create(objconfig.uri);
                            string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(objconfig.Auth));
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
                            AMCInvoiceRequest resp = JsonConvert.DeserializeObject<AMCInvoiceRequest>(responseFromServer);

                            if (jobextcoll.Entities.Count > 0)
                            {
                                continue;
                                Entity JobExtEnt = new Entity(jobextcoll.Entities[0].LogicalName, new Guid("26f4fd31-c8c8-ef11-b8e9-6045bdaaff05"));
                                JobExtEnt["hil_amcsyncstatusmessage"] = resp.LT_TABLE[0].MESSAGE;
                                _service.Update(JobExtEnt);
                            }
                            if (resp.LT_TABLE[0].STATUS == "S")
                            {
                                Entity entUpdate = new Entity(ent.LogicalName, ent.Id);
                                entUpdate["hil_amcsyncstatus"] = new OptionSetValue(2);
                                Console.WriteLine("Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                                requestData.LT_TABLE[0].MESSAGE = resp.LT_TABLE[0].MESSAGE;
                                requestData.LT_TABLE[0].ORDER_NUMBER = resp.LT_TABLE[0].ORDER_NUMBER;
                                requestData.LT_TABLE[0].BILLING_NUMBER = resp.LT_TABLE[0].BILLING_NUMBER;
                                _service.Update(entUpdate);
                                // SaveAMCSAPInvoice(requestData.LT_TABLE[0]);
                            }
                            else
                            {
                                Console.WriteLine("Not Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                            }
                        }
                        else
                        {
                            if (jobextcoll.Entities.Count > 0)
                            {
                                Entity JobExtEnt = new Entity(jobextcoll.Entities[0].LogicalName, jobextcoll.Entities[0].Id);
                                JobExtEnt["cccc"] = "Product code not found on Job";
                                _service.Update(JobExtEnt);
                            }
                            Console.WriteLine("Product Code does not exist: Job# " + jobId + " /" + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                            continue;
                        }
                    }
                }
                else
                {
                    Console.WriteLine("AMC_Invoice_Sync.Program.Main.PushAMCData :: No record found to sync");
                }
                Console.WriteLine("Sync Ended:");
            }
            catch (Exception ex)
            {
                Console.WriteLine("AMC_Invoice_Sync.Program.Main.PushAMCData :: Error While Loading App Settings:" + ex.Message.ToString());
            }
        }
        public void PushAMCData_OmniChannel()
        {
            IntegrationConfig intConfig = IntegrationConfiguration("amc_data_v1");
            string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));

            int totalRecords = 0;
            int recordCnt = 0;
            string jobId = string.Empty;
            try
            {
                #region Old Query
                LinkEntity lnkEntOwner = new LinkEntity
                {
                    LinkFromEntityName = "invoice",
                    LinkToEntityName = "systemuser",
                    LinkFromAttributeName = "owninguser",
                    LinkToAttributeName = "systemuserid",
                    Columns = new ColumnSet("hil_employeecode"),
                    EntityAlias = "user",
                    JoinOperator = JoinOperator.Inner
                };
                LinkEntity lnkEntAddress = new LinkEntity
                {
                    LinkFromEntityName = "invoice",
                    LinkToEntityName = "hil_address",
                    LinkFromAttributeName = "hil_address",
                    LinkToAttributeName = "hil_addressid",
                    Columns = new ColumnSet("hil_street1", "hil_street2", "hil_street3", "hil_state", "hil_pincode", "hil_district", "hil_city",
                    "hil_branch"),
                    EntityAlias = "address",
                    JoinOperator = JoinOperator.Inner
                };
                LinkEntity lnkEntConsumer = new LinkEntity
                {
                    LinkFromEntityName = "invoice",
                    LinkToEntityName = "contact",
                    LinkFromAttributeName = "customerid",
                    LinkToAttributeName = "contactid",
                    Columns = new ColumnSet("hil_salutation", "fullname", "emailaddress1", "mobilephone"),
                    EntityAlias = "contact",
                    JoinOperator = JoinOperator.Inner
                };
                LinkEntity lnkEntProduct = new LinkEntity
                {
                    LinkFromEntityName = "invoice",
                    LinkToEntityName = "product",
                    LinkFromAttributeName = "productid",
                    LinkToAttributeName = "hil_productcode",
                    EntityAlias = "contact",
                    JoinOperator = JoinOperator.Inner
                };
                lnkEntProduct.LinkCriteria = new FilterExpression(LogicalOperator.And);
                lnkEntProduct.LinkCriteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 910590001);

                QueryExpression Query = new QueryExpression("invoice");
                Query.ColumnSet = new ColumnSet("hil_customerasset", "name", "hil_salestype", "hil_productcode", "totallineitemamount", "hil_receiptamount",
                    "discountamount", "hil_productcode", "msdyn_invoicedate", "createdon", "hil_amcsellingsource", "hil_orderid");

                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 2)); //Paid
                Query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 100001)); //Completed
                Query.Criteria.AddCondition(new ConditionExpression("hil_receiptamount", ConditionOperator.GreaterThan, 0)); //Receipt amount must be grater than 0.
                Query.Criteria.AddCondition(new ConditionExpression("hil_productcode", ConditionOperator.NotNull)); //Customer Asset Contains Data
                Query.Criteria.AddCondition(new ConditionExpression("hil_customerasset", ConditionOperator.NotNull)); //Customer Asset Contains Data
                Query.Criteria.AddCondition(new ConditionExpression("hil_salestype", ConditionOperator.Equal, 3));//AMC Omnichannel
                Query.Criteria.AddCondition(new ConditionExpression("createdon", ConditionOperator.On, "2024-10-04"));

                Query.AddOrder("msdyn_invoicedate", OrderType.Ascending);
                Query.AddOrder("createdon", OrderType.Descending);
                Query.LinkEntities.Add(lnkEntOwner);
                Query.LinkEntities.Add(lnkEntAddress);
                Query.LinkEntities.Add(lnkEntConsumer);
                EntityCollection ec = _service.RetrieveMultiple(Query);
                #endregion

                Console.WriteLine("Sync Started:");
                if (ec.Entities.Count > 0)
                {
                    totalRecords = ec.Entities.Count;
                    foreach (Entity ent in ec.Entities)
                    {
                        Entity entUpdate = new Entity(ent.LogicalName, ent.Id);
                        try
                        {
                            recordCnt += 1;
                            jobId = ent.GetAttributeValue<string>("name");
                            AMCInvoiceRequest requestData = new AMCInvoiceRequest();
                            requestData.LT_TABLE = new List<AMCInvoiceData>();
                            AMCInvoiceData invoiceData = new AMCInvoiceData();
                            if (ent.Attributes.Contains("contact.fullname"))
                            {
                                invoiceData.CUSTOMER_NAME = ent.GetAttributeValue<AliasedValue>("contact.fullname").Value.ToString();
                            }
                            if (ent.Attributes.Contains("contact.hil_salutation"))
                            {
                                invoiceData.TITLE = "Mr.";
                            }
                            if (ent.Attributes.Contains("address.hil_street1"))
                            {
                                invoiceData.STREET = ent.GetAttributeValue<AliasedValue>("address.hil_street1").Value.ToString();
                            }
                            if (ent.Attributes.Contains("address.hil_street2"))
                            {
                                invoiceData.HOUSE_NO = ent.GetAttributeValue<AliasedValue>("address.hil_street2").Value.ToString();
                            }
                            if (ent.Attributes.Contains("address.hil_street3"))
                            {
                                invoiceData.STREET4 = ent.GetAttributeValue<AliasedValue>("address.hil_street3").Value.ToString();
                            }
                            if (ent.Attributes.Contains("address.hil_pincode"))
                            {
                                invoiceData.POSTAL_CODE = ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_pincode").Value).Name;
                            }
                            if (ent.Attributes.Contains("address.hil_district"))
                            {
                                invoiceData.DISTRICT = ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_district").Value).Name;
                            }
                            if (ent.Attributes.Contains("address.hil_city"))
                            {
                                invoiceData.CITY = ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_city").Value).Name;
                            }
                            if (ent.Attributes.Contains("address.hil_state"))
                            {
                                Entity entTemp = _service.Retrieve("hil_state", ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_state").Value).Id, new ColumnSet("hil_sapstatecode"));
                                if (entTemp.Attributes.Contains("hil_sapstatecode"))
                                {
                                    invoiceData.REGION_CODE = entTemp.GetAttributeValue<string>("hil_sapstatecode").ToString();
                                }
                            }
                            string PRODUCT_DOP = "";
                            if (ent.Attributes.Contains("hil_customerasset"))
                            {
                                Entity CustomerAsset = _service.Retrieve("msdyn_customerasset", ent.GetAttributeValue<EntityReference>("hil_customerasset").Id, new ColumnSet("hil_warrantytilldate", "hil_invoicedate"));
                                invoiceData.WARNTY_TILL_DATE = CustomerAsset.Contains("hil_warrantytilldate") ? CustomerAsset.GetAttributeValue<DateTime>("hil_warrantytilldate").ToString("yyyy-MM-dd") : "1900-01-01";
                                PRODUCT_DOP = CustomerAsset.Contains("hil_invoicedate") ? CustomerAsset.GetAttributeValue<DateTime>("hil_invoicedate").ToString("yyyy-MM-dd") : "";
                            }
                            if (string.IsNullOrWhiteSpace(PRODUCT_DOP))
                            {
                                PRODUCT_DOP = ent.GetAttributeValue<DateTime>("msdyn_invoicedate").ToString("yyyy-MM-dd");
                            }
                            invoiceData.PRODUCT_DOP = PRODUCT_DOP;
                            Query = new QueryExpression("msdyn_paymentdetail");
                            Query.ColumnSet = new ColumnSet("msdyn_name", "statuscode");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("msdyn_invoice", ConditionOperator.Equal, ent.Id);
                            Query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 910590000);
                            Query.AddOrder("createdon", OrderType.Descending);
                            EntityCollection FoundPaymentDetails = _service.RetrieveMultiple(Query);
                            if (FoundPaymentDetails.Entities.Count == 0)
                            {
                                Console.WriteLine("No Payment Found");
                                continue;
                            }
                            string transId = FoundPaymentDetails[0].GetAttributeValue<string>("msdyn_name");
                            invoiceData.CALL_ID = transId.Replace("D365_", "");
                            string Sellingsource = ent.Contains("hil_amcsellingsource") ? ent.GetAttributeValue<string>("hil_amcsellingsource") : "";
                            if (Sellingsource == "OneWebsite|22")
                            {
                                if (ent.Contains("hil_orderid"))
                                {
                                    string Orderid = ent.GetAttributeValue<string>("hil_orderid").Split('-')[0].Trim();
                                    invoiceData.PAYMENT_REF_NO = Orderid;
                                }
                                else
                                {
                                    Console.WriteLine("Order Id not found.");
                                    continue;
                                }
                            }
                            else
                            {
                                Query = new QueryExpression("hil_paymentstatus");
                                Query.ColumnSet = new ColumnSet("hil_bank_ref_num");
                                Query.Criteria = new FilterExpression(LogicalOperator.And);
                                Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, transId);
                                Query.Criteria.AddCondition("hil_paymentstatus", ConditionOperator.Equal, "success");
                                Query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 1);
                                Query.AddOrder("createdon", OrderType.Descending);
                                EntityCollection EntPaymentstatus = _service.RetrieveMultiple(Query);
                                if (EntPaymentstatus.Entities.Count > 0)
                                {
                                    invoiceData.PAYMENT_REF_NO = EntPaymentstatus.Entities[0].Contains("hil_bank_ref_num") ? EntPaymentstatus.Entities[0].GetAttributeValue<string>("hil_bank_ref_num") : "";
                                }
                                else
                                {
                                    Console.WriteLine("Bank Reference Number not found.");
                                    continue;
                                }
                            }
                            if (ent.Attributes.Contains("address.hil_branch"))
                            {
                                EntityReference BranchRef = (EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_branch").Value;
                                invoiceData.BRANCH_NAME = BranchRef.Name;
                                Entity _entBranch = _service.Retrieve(BranchRef.LogicalName, BranchRef.Id, new ColumnSet("hil_mamorandumcode"));
                                if (_entBranch != null)
                                {
                                    string hil_mamorandumcode = _entBranch.Contains("hil_mamorandumcode") ? _entBranch.GetAttributeValue<string>("hil_mamorandumcode") : null;
                                    invoiceData.SHIP_TO_PARTY = hil_mamorandumcode;
                                    invoiceData.SOLD_TO_PARTY = hil_mamorandumcode;
                                    invoiceData.SHIP_TO_PA_NAME = BranchRef.Name;
                                }
                            }
                            invoiceData.REMARKS = "AMC Sale";
                            if (ent.Attributes.Contains("contact.emailaddress1"))
                            {
                                invoiceData.EMAIL = ent.GetAttributeValue<AliasedValue>("contact.emailaddress1").Value.ToString();
                            }
                            if (ent.Attributes.Contains("contact.mobilephone"))
                            {
                                invoiceData.PHONE = ent.GetAttributeValue<AliasedValue>("contact.mobilephone").Value.ToString();
                            }
                            if (ent.Attributes.Contains("address.hil_branch"))
                            {
                                invoiceData.BRANCH_NAME = ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_branch").Value).Name;
                            }

                            invoiceData.MATERIAL_CODE = ent.GetAttributeValue<EntityReference>("hil_productcode").Name;
                            invoiceData.CS_SERVICE_PRICE = ent.GetAttributeValue<Money>("hil_receiptamount").Value.ToString();

                            if (ent.Attributes.Contains("createdon"))
                            {
                                string jobclosedonValue = string.Empty;
                                jobclosedonValue = ent.GetAttributeValue<DateTime>("createdon").ToString("yyyy-MM-dd");
                                invoiceData.CLOSED_ON = jobclosedonValue;
                                invoiceData.PRICING_DATE = jobclosedonValue;
                            }
                            invoiceData.WARNTY_STATUS = "OUT";

                            if (ent.Attributes.Contains("hil_amcsellingsource"))
                            {
                                string _campaignCode = ent.GetAttributeValue<string>("hil_amcsellingsource").ToString();
                                string[] _campaignCodeArr = _campaignCode.Split('|');
                                if (_campaignCodeArr.Length >= 1)
                                    invoiceData.EMPLOYEE_ID = _campaignCodeArr[0].ToUpper();
                                if (_campaignCodeArr.Length > 1)
                                    invoiceData.EMP_TYPE = _campaignCodeArr[1].ToUpper();
                            }

                            decimal _receiptAmt = ent.GetAttributeValue<Money>("hil_receiptamount").Value;
                            decimal _payableAmt = 0;

                            if (ent.Attributes.Contains("totallineitemamount"))
                            {
                                _payableAmt = ent.GetAttributeValue<Money>("totallineitemamount").Value;
                            }

                            if (_payableAmt < _receiptAmt) continue;

                            invoiceData.ZDS6_AMOUNT = (_payableAmt - _receiptAmt).ToString();
                            invoiceData.DISCOUNT_FLAG = "X";

                            invoiceData.PRODUCT_SERIAL_NO = ent.GetAttributeValue<EntityReference>("hil_customerasset").Name.ToUpper();

                            string _AMCStartDate = GetWarrantyStartDate(ent.GetAttributeValue<EntityReference>("hil_customerasset").Id, ent.GetAttributeValue<DateTime>("createdon"));
                            if (_AMCStartDate != string.Empty)
                            {
                                string _AMCEndDate = GetWarrantyEndDate(ent.GetAttributeValue<EntityReference>("hil_productcode").Id, Convert.ToDateTime(_AMCStartDate));
                                if (_AMCEndDate != null)
                                {
                                    Console.WriteLine("Start Date " + _AMCStartDate + " || EndDate " + _AMCEndDate);
                                    invoiceData.START_DATE = _AMCStartDate;
                                    invoiceData.END_DATE = _AMCEndDate;
                                    invoiceData.ITEM_NO = "1";
                                    requestData.LT_TABLE.Add(invoiceData);
                                    var Json = JsonConvert.SerializeObject(requestData);

                                    Console.WriteLine(Json);
                                    WebRequest request = WebRequest.Create(intConfig.uri);
                                    request.Headers[HttpRequestHeader.Authorization] = authInfo;
                                    request.Method = "POST";
                                    byte[] byteArray = Encoding.UTF8.GetBytes(Json);
                                    request.ContentType = "application/x-www-form-urlencoded";
                                    request.ContentLength = byteArray.Length;
                                    Stream dataStream = request.GetRequestStream();
                                    dataStream.Write(byteArray, 0, byteArray.Length);
                                    dataStream.Close();
                                    WebResponse response = request.GetResponse();
                                    Console.WriteLine(((HttpWebResponse)response).StatusDescription);
                                    dataStream = response.GetResponseStream();
                                    StreamReader reader = new StreamReader(dataStream);
                                    string responseFromServer = reader.ReadToEnd();
                                    AMCInvoiceRequest resp = JsonConvert.DeserializeObject<AMCInvoiceRequest>(responseFromServer);
                                    Console.WriteLine(responseFromServer);
                                    entUpdate["hil_sapsyncmessage"] = resp.LT_TABLE[0].MESSAGE;

                                    if (resp.LT_TABLE[0].STATUS == "S")
                                    {
                                        entUpdate["statecode"] = new OptionSetValue(2);
                                        entUpdate["statuscode"] = new OptionSetValue(910590000);
                                        entUpdate["hil_sapinvoicenumber"] = resp.LT_TABLE[0].BILLING_NUMBER;
                                        entUpdate["hil_sapsonumber"] = resp.LT_TABLE[0].ORDER_NUMBER;
                                        Console.WriteLine("Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                                        requestData.LT_TABLE[0].MESSAGE = resp.LT_TABLE[0].MESSAGE;
                                        requestData.LT_TABLE[0].ORDER_NUMBER = resp.LT_TABLE[0].ORDER_NUMBER;
                                        requestData.LT_TABLE[0].BILLING_NUMBER = resp.LT_TABLE[0].BILLING_NUMBER;
                                        //SaveAMCSAPInvoice(requestData.LT_TABLE[0]);
                                    }
                                    else
                                    {
                                        Console.WriteLine("Not Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                                    }
                                }
                                else
                                {
                                    entUpdate["hil_sapsyncmessage"] = "Warranty Template does not Exist.";
                                    Console.WriteLine("AMC End Date is NUll: Invoice# " + jobId + " / " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                                }
                            }
                            else
                            {
                                entUpdate["hil_sapsyncmessage"] = "Warranty Start Date does not Exist.";
                                Console.WriteLine("AMC Start Date is NUll: Invoice# " + jobId + " / " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                            }
                        }
                        catch (Exception ex)
                        {
                            entUpdate["hil_sapsyncmessage"] = ex.Message;
                            Console.WriteLine("Error " + ex.Message);
                        }
                        _service.Update(entUpdate);
                    }
                }
                else
                {
                    Console.WriteLine("AMCE2ESyncToSAP.ClsAMCE2E.Program.Main.PushAMCData_OmniChannel :: No record found to sync");
                }
                Console.WriteLine("Sync Ended:");
            }
            catch (Exception ex)
            {
                Console.WriteLine("AMCE2ESyncToSAP.ClsAMCE2E.Program.Main.PushAMCData_OmniChannel :: Error While Loading App Settings:" + ex.Message.ToString());
            }
        }
        #region Check payment status from PayU
        private List<TransactionDetails> getJobsForPaymentStatus()
        {
            string _paymentEffectiveFrom = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
            List<TransactionDetails> lstTransactionDetails = new List<TransactionDetails>();
            string fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='hil_paymentstatus'>
                        <attribute name='hil_paymentstatusid' />
                        <attribute name='hil_name' />
                        <attribute name='hil_job'/>
                        <order attribute='hil_name' descending='false' />
                        <filter type='and'>
                          <condition attribute='hil_paymentstatus' operator='ne' value='success' />
                        </filter>
                        <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='hil_job' link-type='inner' alias='ab'>
                          <attribute name='msdyn_name' />
                          <filter type='and'>
                            <condition attribute='hil_callsubtype' operator='eq' uiname='AMC Call' uitype='hil_callsubtype' value='{{55A71A52-3C0B-E911-A94E-000D3AF06CD4}}' />
                            <condition attribute='msdyn_substatus' operator='in'>
                                <value uiname='Work Done' uitype='msdyn_workordersubstatus'>{{2927FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                                <value uiname='Work Done SMS' uitype='msdyn_workordersubstatus'>{{7E85074C-9C54-E911-A951-000D3AF0677F}}</value>
                                <value uiname='Closed' uitype='msdyn_workordersubstatus'>{{1727FA6C-FA0F-E911-A94E-000D3AF060A1}}</value>
                            </condition>
                            <condition attribute='hil_isocr' operator='ne' value='1' />
                            <condition attribute='hil_receiptamount' operator='gt' value='0' />
                            <condition attribute='hil_amcsyncstatus' operator='eq' value='1' />
                            <condition attribute='createdon' operator='on-or-after' value='{_paymentEffectiveFrom}' />
                            <condition attribute='hil_actualcharges' operator='gt' value='0' />
                            <condition attribute='hil_modeofpayment' operator='in'>
                              <value>1</value>
                              <value>6</value>
                            </condition>
                          </filter>
                        </link-entity>
                      </entity>
                    </fetch>";
            EntityCollection coll = _service.RetrieveMultiple(new FetchExpression(fetchXML));
            foreach (Entity entity in coll.Entities)
            {
                TransactionDetails transactionDetails = new TransactionDetails();
                transactionDetails.transactionID = entity.GetAttributeValue<string>("hil_name");
                transactionDetails.Regarding = entity.GetAttributeValue<EntityReference>("hil_job");
                lstTransactionDetails.Add(transactionDetails);
            }
            return lstTransactionDetails;
        }
        private List<TransactionDetails> getInvoiceForPaymentStatus()
        {
            List<TransactionDetails> lstTransactionDetails = new List<TransactionDetails>();
            string _paymentEffectiveFrom = DateTime.Now.AddDays(-3).ToString("yyyy-MM-dd");

            QueryExpression Query = new QueryExpression("msdyn_paymentdetail");
            Query.ColumnSet = new ColumnSet("msdyn_name", "msdyn_invoice");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);//Active
            Query.Criteria.AddCondition("statuscode", ConditionOperator.NotEqual, 910590000);//Not Received
            Query.Criteria.AddCondition("msdyn_invoice", ConditionOperator.NotNull);
            Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, _paymentEffectiveFrom);
            Query.Orders.Add(new OrderExpression("createdon", OrderType.Descending));
            EntityCollection coll = _service.RetrieveMultiple(Query);
            foreach (Entity entity in coll.Entities)
            {
                TransactionDetails transactionDetails = new TransactionDetails();
                transactionDetails.transactionID = entity.GetAttributeValue<string>("msdyn_name");
                transactionDetails.Regarding = entity.GetAttributeValue<EntityReference>("msdyn_invoice");
                lstTransactionDetails.Add(transactionDetails);
            }
            return lstTransactionDetails;
        }
        #endregion

        #region AMC start and end date calculation
        private string GetWarrantyStartDate(Guid AssetID, DateTime _purchaseDate)
        {
            string WarrantyStartDate = _purchaseDate.ToString("yyyy-MM-dd");
            try
            {
                LinkEntity lnkEntInvoice = new LinkEntity
                {
                    LinkFromEntityName = "hil_unitwarranty",
                    LinkToEntityName = "hil_warrantytemplate",
                    LinkFromAttributeName = "hil_warrantytemplate",
                    LinkToAttributeName = "hil_warrantytemplateid",
                    Columns = new ColumnSet("hil_type"),
                    EntityAlias = "invoice",
                    JoinOperator = JoinOperator.Inner
                };
                QueryExpression Query = new QueryExpression("hil_unitwarranty");
                Query.ColumnSet = new ColumnSet("hil_name", "hil_warrantyenddate", "hil_warrantytemplate");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("hil_customerasset", ConditionOperator.Equal, AssetID));
                Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));  //Active
                Query.AddOrder("hil_warrantyenddate", OrderType.Descending);
                Query.LinkEntities.Add(lnkEntInvoice);
                EntityCollection ec = _service.RetrieveMultiple(Query);
                if (ec.Entities.Count >= 1)
                {
                    int WarrantyType = ((OptionSetValue)ec.Entities[0].GetAttributeValue<AliasedValue>("invoice.hil_type").Value).Value;
                    DateTime _warrantyTempDate = ec.Entities[0].GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);
                    if (WarrantyType == 1 || WarrantyType == 3)
                    {
                        if (_warrantyTempDate >= _purchaseDate)
                        {
                            WarrantyStartDate = _warrantyTempDate.AddDays(1).ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            WarrantyStartDate = _purchaseDate.ToString("yyyy-MM-dd");
                        }
                        return WarrantyStartDate;
                    }
                    else
                    {
                        foreach (Entity entity in ec.Entities)
                        {
                            string fetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_labor'>
                                <attribute name='hil_laborid' />
                                <attribute name='hil_name' />
                                <attribute name='createdon' />
                                <order attribute='hil_name' descending='false' />
                                <filter type='and'>
                                    <condition attribute='hil_includedinwarranty' operator='eq' value='2' />
                                </filter>
                                <link-entity name='hil_warrantytemplate' from='hil_warrantytemplateid' to='hil_warrantytemplateid' link-type='inner' alias='aa'>
                                <filter type='and'>
                                    <condition attribute='hil_warrantytemplateid' operator='eq' value='{entity.GetAttributeValue<EntityReference>("hil_warrantytemplate").Id}'/>
                                </filter>
                                </link-entity>
                                </entity>
                                </fetch>";
                            EntityCollection ec1 = _service.RetrieveMultiple(new FetchExpression(fetch));
                            if (ec1.Entities.Count == 0)
                            {
                                _warrantyTempDate = entity.GetAttributeValue<DateTime>("hil_warrantyenddate").AddMinutes(330);
                                if (_warrantyTempDate >= _purchaseDate)
                                {
                                    WarrantyStartDate = _warrantyTempDate.AddDays(1).ToString("yyyy-MM-dd");
                                }
                                else
                                {
                                    WarrantyStartDate = _purchaseDate.ToString("yyyy-MM-dd");
                                }
                                return WarrantyStartDate;
                            }
                        }
                    }
                }
                if (ec.Entities.Count == 0)
                {
                    Query = new QueryExpression("hil_unitwarranty");
                    Query.ColumnSet = new ColumnSet("hil_warrantystartdate", "hil_warrantytemplate");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition(new ConditionExpression("hil_customerasset", ConditionOperator.Equal, AssetID));
                    EntityCollection ec1 = _service.RetrieveMultiple(Query);
                    if (ec1.Entities.Count == 0)
                    {
                        //WarrantyStartDate = string.Empty;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return WarrantyStartDate;
        }
        private string GetWarrantyEndDate(Guid _AMCPlaneID, DateTime StartDate)
        {
            string WarrantyEndDate = null;
            QueryExpression Query = new QueryExpression("hil_warrantytemplate");
            Query.ColumnSet = new ColumnSet("hil_warrantyperiod");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("hil_amcplan", ConditionOperator.Equal, _AMCPlaneID));
            Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));  //Active
            Query.TopCount = 1;
            Query.AddOrder("createdon", OrderType.Descending);
            EntityCollection ec = _service.RetrieveMultiple(Query);
            if (ec.Entities.Count == 1)
            {
                WarrantyEndDate = StartDate.AddMonths(ec[0].GetAttributeValue<int>("hil_warrantyperiod")).AddDays(-1).ToString("yyyy-MM-dd");
            }
            return WarrantyEndDate;
        }
        #endregion

        #region Supporting Methods
        private ActionResponse SaveAMCSAPInvoice(AMCInvoiceData objamcsapdetails)
        {
            ActionResponse objresult = new ActionResponse();
            try
            {
                QueryExpression query = new QueryExpression("hil_amcstaging");
                query.ColumnSet = new ColumnSet("hil_serailnumber");
                query.NoLock = true;
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition("hil_serailnumber", ConditionOperator.Equal, objamcsapdetails.PRODUCT_SERIAL_NO);
                query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, objamcsapdetails.BILLING_NUMBER);
                EntityCollection objEntity = _service.RetrieveMultiple(query);
                if (objEntity.Entities.Count <= 0)
                {
                    QueryExpression qsCType = new QueryExpression("product");
                    qsCType.ColumnSet = new ColumnSet("name");
                    qsCType.NoLock = true;
                    qsCType.Criteria.AddCondition("name", ConditionOperator.Equal, objamcsapdetails.MATERIAL_CODE);
                    EntityCollection AmcplanGuid = _service.RetrieveMultiple(qsCType);

                    Entity entAMCUpdate = new Entity("hil_amcstaging");
                    entAMCUpdate["hil_callid"] = objamcsapdetails.CALL_ID;
                    entAMCUpdate["hil_serailnumber"] = objamcsapdetails.PRODUCT_SERIAL_NO;
                    entAMCUpdate["hil_mobilenumber"] = objamcsapdetails.PHONE;
                    entAMCUpdate["hil_warrantystartdate"] = Convert.ToDateTime(objamcsapdetails.START_DATE).AddMinutes(330);
                    entAMCUpdate["hil_warrantyenddate"] = Convert.ToDateTime(objamcsapdetails.END_DATE).AddMinutes(330);
                    entAMCUpdate["hil_amcplannameslt"] = objamcsapdetails.MATERIAL_CODE;
                    entAMCUpdate["hil_name"] = objamcsapdetails.BILLING_NUMBER;
                    entAMCUpdate["hil_sapbillingdate"] = DateTime.Now;
                    entAMCUpdate["hil_salesordernumber"] = objamcsapdetails.ORDER_NUMBER;
                    entAMCUpdate["hil_branchmemorandumcode"] = objamcsapdetails.SOLD_TO_PARTY;
                    if (AmcplanGuid.Entities.Count > 0)
                    {
                        if (AmcplanGuid.Entities[0].Id != Guid.Empty)
                            entAMCUpdate["hil_amcplan"] = new EntityReference("product", AmcplanGuid.Entities[0].Id);
                    }
                    string EncryptBIllingNumber = new ClsEncryptDecrypt().EncryptAES256URL(objamcsapdetails.BILLING_NUMBER);
                    entAMCUpdate["hil_sapbillingdocpath"] = ConfigurationManager.AppSettings["Docurl"].ToString() + EncryptBIllingNumber;
                    _service.Create(entAMCUpdate);
                    objresult.Is_Successful = true;
                    objresult.Message = "Success";
                    return objresult;
                }
                else
                {
                    objresult.Is_Successful = false;
                    objresult.Message = "Failed";
                    return objresult;
                }
            }
            catch (Exception ex)
            {
                objresult.Is_Successful = false;
                objresult.Message = "Failed";
                return objresult;
            }
        }
        private string GetMaterialCode(Guid jobGuId, IOrganizationService _service)
        {
            string productCode = null;
            // Select only Replaced Part with Hierarchy Level - AMC and Field Service Product Type - Non-Inventory Product
            string fetchXML = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>";
            fetchXML += "<entity name='msdyn_workorderproduct'>";
            fetchXML += "<attribute name='createdon' />";
            fetchXML += "<attribute name='hil_replacedpart' />";
            fetchXML += "<filter type='and'>";
            fetchXML += "<condition attribute='msdyn_workorder' operator='eq' value='{" + jobGuId + @"}' />";
            fetchXML += "</filter>";
            fetchXML += "<link-entity name='product' from='productid' to='hil_replacedpart' visible='false' link-type='inner' alias='prod'>";
            fetchXML += "<attribute name='productnumber' />";
            fetchXML += "<filter type='and'>";
            fetchXML += "<condition attribute='hil_hierarchylevel' operator='eq' value='910590001' />";
            fetchXML += "<condition attribute='msdyn_fieldserviceproducttype' operator='eq' value='690970001' />";
            fetchXML += "</filter>";
            fetchXML += "</link-entity>";
            fetchXML += "</entity>";
            fetchXML += "</fetch>";
            EntityCollection ec = _service.RetrieveMultiple(new FetchExpression(fetchXML));
            if (ec.Entities.Count > 0)
            {
                if (ec.Entities[0].Attributes.Contains("prod.productnumber"))
                {
                    productCode = ec.Entities[0].GetAttributeValue<AliasedValue>("prod.productnumber").Value.ToString() + "|" + ec.Entities[0].GetAttributeValue<EntityReference>("hil_replacedpart").Id.ToString();
                }
            }
            return productCode;
        }
        private IntegrationConfig IntegrationConfiguration(string Param)
        {
            IntegrationConfig output = new IntegrationConfig();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = _service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            return output;
        }
        #endregion
    }
}
