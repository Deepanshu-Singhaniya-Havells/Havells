using Havells.Crm.CommonLibrary;
using Havells.CRM.WP_AMC_SalesSyncToSAP.Model;
using Havells_Plugin.WOProduct;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Activities.Presentation.Converters;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Description;

namespace Havells.CRM.WP_AMC_SalesSyncToSAP
{
    public class Functions : AzureWebJobsLogs
    {
        static string containerName = "webjob-amcsapinvoices";
        public static IOrganizationService createConnection(string connectionString, string fileName)
        {
            IOrganizationService organizationService = null;
            try
            {
                organizationService = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)organizationService).LastCrmException != null && (((CrmServiceClient)organizationService).LastCrmException.Message == "OrganizationWebProxyClient is null" || ((CrmServiceClient)organizationService).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    string msg = ((CrmServiceClient)organizationService).LastCrmException.Message;
                    Console.WriteLine(msg);
                    CreateOrUpdateLogs(containerName, fileName, msg, null);
                    throw new Exception(((CrmServiceClient)organizationService).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                CreateOrUpdateLogs(containerName, fileName, "Error while Creating Conn: " + ex.Message, null);
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }

            return organizationService;
        }
        public static void getPaymentStatusJob(IOrganizationService service, string fileName)
        {
            string msg = "";
            IntegrationConfig intConfig = IntegrationConfiguration(service, "Send Payment Link", fileName);
            //intConfig.uri = "https://p90ci.havells.com:50001/RESTAdapter/PayUBiz/TransactionAPI";
            string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));

            QueryExpression Query = new QueryExpression("hil_paymentstatus");
            Query.ColumnSet = new ColumnSet("hil_job", "hil_name");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_paymentstatus", ConditionOperator.NotEqual, "success");
            // Query.Criteria.AddCondition("hil_paymentstatus", ConditionOperator.NotEqual, "failure");
            Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, DateTime.Today.AddDays(-7));
            Query.Criteria.AddCondition("hil_job", ConditionOperator.NotNull);
            Query.Orders.Add(new OrderExpression("createdon", OrderType.Ascending));
            EntityCollection paymentCollection = service.RetrieveMultiple(Query);

            int total = paymentCollection.Entities.Count;
            msg = "Total Payment Collection is " + total;
            Console.WriteLine(msg);
            CreateOrUpdateLogs(containerName, fileName, msg, null);

            int count = 0;
            foreach (Entity entity in paymentCollection.Entities)
            {
                count++;
                string transactionID = entity.GetAttributeValue<string>("hil_name");
                msg = count + "/" + total + " Status Updated of transactionid " + transactionID + " Status is " + updatePaymentStatusforJob(transactionID, service, intConfig.uri, authInfo, fileName);
                Console.WriteLine(msg);
                CreateOrUpdateLogs(containerName, fileName, msg, null);
            }
            msg = "-------------------------------------------Done----------------------------------------";
            Console.WriteLine(msg);
            CreateOrUpdateLogs(containerName, fileName, msg, null);



        }
        public static void getPaymentStatusofInvoice(IOrganizationService service, string fileName)
        {
            string msg = "";
            IntegrationConfig intConfig = IntegrationConfiguration(service, "Send Payment Link", fileName);
            //intConfig.uri = "https://p90ci.havells.com:50001/RESTAdapter/PayUBiz/TransactionAPI";
            //intConfig.Auth = "D365_Havells:QAD365@1234";
            string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));

            QueryExpression Query = new QueryExpression("msdyn_paymentdetail");
            Query.ColumnSet = new ColumnSet("msdyn_name", "msdyn_invoice");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("statuscode", ConditionOperator.NotEqual, 910590000);
            Query.Criteria.AddCondition("statuscode", ConditionOperator.NotEqual, 910590001);
            Query.Criteria.AddCondition("msdyn_invoice", ConditionOperator.NotNull);
            Query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, new DateTime(2023, 04, 05));
            //Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Like, "%INV-2023-003091%");
            Query.Orders.Add(new OrderExpression("createdon", OrderType.Descending));
            EntityCollection paymentCollection = service.RetrieveMultiple(Query);

            int total = paymentCollection.Entities.Count;
            msg = "Total Payment Collection is " + total;
            Console.WriteLine(msg);
            CreateOrUpdateLogs(containerName, fileName, msg, null);

            int count = 0;
            foreach (Entity entity in paymentCollection.Entities)
            {
                count++;
                string transactionID = entity.GetAttributeValue<string>("msdyn_name");
                EntityReference Invoice = entity.GetAttributeValue<EntityReference>("msdyn_invoice");
                updatePaymentStatusforInvoice(transactionID, Invoice, service, intConfig.uri, authInfo, fileName);
                msg = count + "/" + total + " Status Updated of transactionid " + transactionID;
                Console.WriteLine(msg);
                CreateOrUpdateLogs(containerName, fileName, msg, null);

            }
            msg = "-------------------------------------------Done----------------------------------------";
            Console.WriteLine(msg);
            CreateOrUpdateLogs(containerName, fileName, msg, null);

        }
        public static void PushAMCData_OmniChannel(IOrganizationService service, string fileName)
        {
            string msg = "";
            IntegrationConfig intConfig = IntegrationConfiguration(service, "amc_data", fileName);
            intConfig.uri = "https://p90ci.havells.com:50001/RESTAdapter/dynamics/amc_data";
            //intConfig.Auth = "D365_Havells:PRDD365@1234
            string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));

            int totalRecords = 0;
            int recordCnt = 0;
            string prodCode = string.Empty;
            string jobId = string.Empty;
            try
            {
                #region Old Query
                LinkEntity lnkEntOwner = new LinkEntity
                {
                    LinkFromEntityName = "invoice",
                    LinkToEntityName = SystemUser.EntityLogicalName,
                    LinkFromAttributeName = "owninguser",
                    LinkToAttributeName = "systemuserid",
                    Columns = new ColumnSet("hil_employeecode"),
                    EntityAlias = "user",
                    JoinOperator = JoinOperator.Inner
                };
                LinkEntity lnkEntAddress = new LinkEntity
                {
                    LinkFromEntityName = "invoice",
                    LinkToEntityName = hil_address.EntityLogicalName,
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
                    LinkToEntityName = Contact.EntityLogicalName,
                    LinkFromAttributeName = "customerid",
                    LinkToAttributeName = "contactid",
                    Columns = new ColumnSet("hil_salutation", "fullname", "emailaddress1", "mobilephone"),
                    EntityAlias = "contact",
                    JoinOperator = JoinOperator.Inner
                };
                LinkEntity lnkEntProduct = new LinkEntity
                {
                    LinkFromEntityName = "invoice",
                    LinkToEntityName = Product.EntityLogicalName,
                    LinkFromAttributeName = "productid",
                    LinkToAttributeName = "hil_productcode",
                    EntityAlias = "contact",
                    JoinOperator = JoinOperator.Inner
                };
                lnkEntProduct.LinkCriteria = new FilterExpression(LogicalOperator.And);
                lnkEntProduct.LinkCriteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 910590001);


                QueryExpression Query = new QueryExpression("invoice");
                Query.ColumnSet = new ColumnSet("hil_customerasset", "name", "hil_salestype", "hil_productcode", "totallineitemamount", "hil_receiptamount",
                    "discountamount", "hil_productcode", "msdyn_invoicedate", "createdon", "hil_amcsellingsource");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 2)); //Paid

                Query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 100001)); //Completed

                Query.Criteria.AddCondition(new ConditionExpression("hil_receiptamount", ConditionOperator.GreaterThan, 0)); //Receipt amount must be grater than 0.
                Query.Criteria.AddCondition(new ConditionExpression("hil_customerasset", ConditionOperator.NotNull)); //Customer Asset Contains Data

                //Query.Criteria.AddCondition(new ConditionExpression("name", ConditionOperator.Equal, "INV-2023-015734")); //Invoice ID

                Query.Criteria.AddCondition(new ConditionExpression("hil_salestype", ConditionOperator.Equal, 3));//AMC Omnichannel
                Query.AddOrder("msdyn_invoicedate", OrderType.Ascending);
                Query.LinkEntities.Add(lnkEntOwner);
                Query.LinkEntities.Add(lnkEntAddress);
                Query.LinkEntities.Add(lnkEntConsumer);
                EntityCollection ec = service.RetrieveMultiple(Query);
                #endregion
                msg = "Sync Started:";
                Console.WriteLine(msg);
                CreateOrUpdateLogs(containerName, fileName, msg, null);

                if (ec.Entities.Count > 0)
                {
                    AMCSAPINvoiceTableCreate _AMCSAPINvoiceTableCreate = new AMCSAPINvoiceTableCreate();

                    totalRecords = ec.Entities.Count;
                    foreach (Entity ent in ec.Entities)
                    {
                        try
                        {
                            recordCnt += 1;
                            prodCode = ent.GetAttributeValue<EntityReference>("hil_productcode").Name;// GetMaterialCode(ent.Id);
                            _AMCSAPINvoiceTableCreate.AMCPlan = ent.GetAttributeValue<EntityReference>("hil_productcode");
                            jobId = ent.GetAttributeValue<string>("name");

                            //if (jobId == "INV-2023-000150")
                            if (prodCode != null)
                            {
                                AMCInvoiceRequest requestData = new AMCInvoiceRequest();
                                requestData.IM_DATA = new System.Collections.Generic.List<AMCInvoiceData>();
                                AMCInvoiceData invoiceData;
                                invoiceData = new AMCInvoiceData();
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
                                    Entity entTemp = service.Retrieve("hil_state", ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_state").Value).Id, new ColumnSet("hil_sapstatecode"));
                                    if (entTemp.Attributes.Contains("hil_sapstatecode"))
                                    {
                                        invoiceData.REGION_CODE = entTemp.GetAttributeValue<string>("hil_sapstatecode").ToString();
                                    }
                                }
                                Query = new QueryExpression("msdyn_paymentdetail");
                                Query.ColumnSet = new ColumnSet("msdyn_name", "statuscode");
                                Query.Criteria = new FilterExpression(LogicalOperator.And);
                                Query.Criteria.AddCondition("msdyn_invoice", ConditionOperator.Equal, ent.Id);
                                Query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 910590000);
                                Query.AddOrder("createdon", OrderType.Descending);
                                EntityCollection FoundPaymentDetails = service.RetrieveMultiple(Query);
                                if (FoundPaymentDetails.Entities.Count == 0)
                                {
                                    msg = "No Payment Found";
                                    Console.WriteLine(msg);
                                    CreateOrUpdateLogs(containerName, fileName, msg, null);

                                    continue;
                                }
                                string transId = FoundPaymentDetails[0].GetAttributeValue<string>("msdyn_name");
                                invoiceData.CALL_ID = transId.Replace("D365_", "");
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

                                invoiceData.MATERIAL_CODE = prodCode;

                                invoiceData.CS_SERVICE_PRICE = ent.GetAttributeValue<Money>("hil_receiptamount").Value.ToString();

                                //if (ent.Attributes.Contains("user.hil_employeecode"))
                                //{
                                //    invoiceData.EMPLOYEE_ID = ent.GetAttributeValue<AliasedValue>("user.hil_employeecode").Value.ToString();
                                //}
                                if (ent.Attributes.Contains("createdon"))
                                {
                                    string jobclosedonValue = string.Empty;
                                    jobclosedonValue = ent.GetAttributeValue<DateTime>("createdon").ToString("yyyy-MM-dd");
                                    invoiceData.CLOSED_ON = jobclosedonValue;

                                    //string pricingDateValue = string.Empty;
                                    //pricingDateValue = ent.GetAttributeValue<DateTime>("createdon").ToString("yyyy-MM-dd");
                                    invoiceData.PRICING_DATE = jobclosedonValue;
                                }
                                invoiceData.WARNTY_STATUS = "OUT";

                                if (ent.Attributes.Contains("hil_amcsellingsource"))
                                {
                                    invoiceData.EMPLOYEE_ID = ent.GetAttributeValue<string>("hil_amcsellingsource").ToString();
                                }

                                Decimal _receiptAmt = ent.GetAttributeValue<Money>("hil_receiptamount").Value;
                                Decimal _payableAmt = 0;

                                if (ent.Attributes.Contains("totallineitemamount"))
                                {
                                    _payableAmt = ent.GetAttributeValue<Money>("totallineitemamount").Value;
                                }

                                if (_payableAmt < _receiptAmt) continue;

                                invoiceData.ZDS6_AMOUNT = (_payableAmt - _receiptAmt).ToString();
                                invoiceData.DISCOUNT_FLAG = "X";

                                invoiceData.PRODUCT_SERIAL_NO = ent.GetAttributeValue<EntityReference>("hil_customerasset").Name;
                                //invoiceData.PRODUCT_DOP = "";
                                if (ent.Attributes.Contains("msdyn_invoicedate"))
                                {
                                    invoiceData.PRODUCT_DOP = ent.GetAttributeValue<DateTime>("msdyn_invoicedate").ToString("yyyy-MM-dd");
                                }

                                string _AMCStartDate = GetWarrantyStartDate(service, ent.GetAttributeValue<EntityReference>("hil_customerasset").Id, ent.GetAttributeValue<DateTime>("createdon"), fileName);// DateTime.Now;
                                if (_AMCStartDate != string.Empty)
                                {
                                    string _AMCEndDate = GetWarrantyEndDate(service, ent.GetAttributeValue<EntityReference>("hil_productcode").Id, Convert.ToDateTime(_AMCStartDate), fileName);
                                    if (_AMCEndDate != null)
                                    {
                                        msg = "Start Date " + _AMCStartDate + " || EndDate " + _AMCEndDate;
                                        Console.WriteLine(msg);
                                        CreateOrUpdateLogs(containerName, fileName, msg, null);

                                        invoiceData.START_DATE = _AMCStartDate;
                                        invoiceData.END_DATE = _AMCEndDate;


                                        _AMCSAPINvoiceTableCreate.MobileNumber = invoiceData.PHONE;
                                        _AMCSAPINvoiceTableCreate.SerialNumber = invoiceData.PRODUCT_SERIAL_NO;
                                        _AMCSAPINvoiceTableCreate.CallID = invoiceData.CALL_ID;
                                        _AMCSAPINvoiceTableCreate.WarrantyStartDate = Convert.ToDateTime(invoiceData.START_DATE);
                                        _AMCSAPINvoiceTableCreate.WarrantyEndDate = Convert.ToDateTime(invoiceData.END_DATE);
                                        requestData.IM_DATA.Add(invoiceData);

                                        var Json = JsonConvert.SerializeObject(requestData);

                                        CreateSAPAMCRecord(service, _AMCSAPINvoiceTableCreate, fileName);

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
                                        msg = ((HttpWebResponse)response).StatusDescription;
                                        Console.WriteLine(msg);
                                        CreateOrUpdateLogs(containerName, fileName, msg, null);

                                        dataStream = response.GetResponseStream();
                                        StreamReader reader = new StreamReader(dataStream);
                                        string responseFromServer = reader.ReadToEnd();
                                        AMCInvoiceData resp = JsonConvert.DeserializeObject<AMCInvoiceData>(responseFromServer);
                                        if (responseFromServer.Replace(@"""", "").IndexOf("STATUS:S,") > 0)
                                        {
                                            SetStateRequest req = new SetStateRequest();
                                            req.State = new OptionSetValue(2);
                                            req.Status = new OptionSetValue(910590000);
                                            req.EntityMoniker = ent.ToEntityReference();
                                            var res = (SetStateResponse)service.Execute(req);
                                            AMCSAPINvoiceTableUpdate _AMCSAPINvoiceTableUpdate = new AMCSAPINvoiceTableUpdate();
                                            _AMCSAPINvoiceTableUpdate.WarrantyStartDate = _AMCSAPINvoiceTableCreate.WarrantyStartDate;
                                            _AMCSAPINvoiceTableUpdate.WarrantyEndDate = _AMCSAPINvoiceTableCreate.WarrantyEndDate;
                                            _AMCSAPINvoiceTableUpdate.SerialNumber = _AMCSAPINvoiceTableCreate.SerialNumber;
                                            _AMCSAPINvoiceTableUpdate.MobileNumber = _AMCSAPINvoiceTableCreate.MobileNumber;
                                            _AMCSAPINvoiceTableUpdate.CallID = _AMCSAPINvoiceTableCreate.CallID;
                                            _AMCSAPINvoiceTableUpdate.invoiceStatus = (int)AMCSAPStatus.Posted;

                                            UpdateSAPAMCRecord(service, _AMCSAPINvoiceTableUpdate, fileName);
                                            msg = "Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name");
                                            Console.WriteLine(msg);
                                            CreateOrUpdateLogs(containerName, fileName, msg, null);
                                        }
                                        else
                                        {
                                            msg = "Not Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name");
                                            Console.WriteLine(msg);
                                            CreateOrUpdateLogs(containerName, fileName, msg, null);
                                        }
                                    }
                                    else
                                    {
                                        msg = "AMC End Date is NUll: Invoice# " + jobId + " / " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name");
                                        Console.WriteLine(msg);
                                        CreateOrUpdateLogs(containerName, fileName, msg, null);
                                    }
                                }
                                else
                                {
                                    msg = "AMC Start Date is NUll: Invoice# " + jobId + " / " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name");
                                    CreateOrUpdateLogs(containerName, fileName, msg, null);
                                    Console.WriteLine(msg);
                                }
                            }
                            else
                            {
                                msg = "Product Code does not exist: Invoice " + jobId + " / " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name");
                                Console.WriteLine(msg);
                                CreateOrUpdateLogs(containerName, fileName, msg, null);
                            }
                        }
                        catch (Exception ex)
                        {
                            msg = "Error " + ex.Message;
                            Console.WriteLine(msg);
                            CreateOrUpdateLogs(containerName, fileName, msg, null);
                        }
                    }
                }
                else
                {
                    msg = "AMC_Invoice_Sync.Program.Main.PushAMCData_WP :: No record found to sync";
                    Console.WriteLine(msg);
                    CreateOrUpdateLogs(containerName, fileName, msg, null);
                }
                msg = "Sync Ended:";
                Console.WriteLine(msg);
                CreateOrUpdateLogs(containerName, fileName, msg, null);
            }
            catch (Exception ex)
            {
                msg = "AMC_Invoice_Sync.Program.Main.PushAMCData_WP :: Error While Loading App Settings:" + ex.Message.ToString();
                Console.WriteLine(msg);
                CreateOrUpdateLogs(containerName, fileName, msg, null);
            }
        }
        public static void PushAMCData(IOrganizationService service, string fileName)
        {
            string msg = "";
            IntegrationConfig intConfig = IntegrationConfiguration(service, "amc_data", fileName);
            //intConfig.uri = "https://p90ci.havells.com:50001/RESTAdapter/dynamics/amc_data";
            //intConfig.Auth = "D365_Havells:PRDD365@1234
            string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));

            int totalRecords = 0;
            int recordCnt = 0;
            string prodCode = string.Empty;
            string jobId = string.Empty;

            try
            {
                LinkEntity lnkEntOwner = new LinkEntity
                {
                    LinkFromEntityName = msdyn_workorder.EntityLogicalName,
                    LinkToEntityName = SystemUser.EntityLogicalName,
                    LinkFromAttributeName = "owninguser",
                    LinkToAttributeName = "systemuserid",
                    Columns = new ColumnSet("hil_employeecode"),
                    EntityAlias = "user",
                    JoinOperator = JoinOperator.Inner
                };
                LinkEntity lnkEntAddress = new LinkEntity
                {
                    LinkFromEntityName = msdyn_workorder.EntityLogicalName,
                    LinkToEntityName = hil_address.EntityLogicalName,
                    LinkFromAttributeName = "hil_address",
                    LinkToAttributeName = "hil_addressid",
                    Columns = new ColumnSet("hil_street1", "hil_street2", "hil_street3", "hil_state", "hil_pincode", "hil_district", "hil_city"),
                    EntityAlias = "address",
                    JoinOperator = JoinOperator.Inner
                };
                LinkEntity lnkEntConsumer = new LinkEntity
                {
                    LinkFromEntityName = msdyn_workorder.EntityLogicalName,
                    LinkToEntityName = Contact.EntityLogicalName,
                    LinkFromAttributeName = "hil_customerref",
                    LinkToAttributeName = "contactid",
                    Columns = new ColumnSet("hil_salutation", "fullname", "emailaddress1"),
                    EntityAlias = "contact",
                    JoinOperator = JoinOperator.Inner
                };
                LinkEntity linkEntityForCustomerassets = new LinkEntity
                {
                    LinkFromEntityName = msdyn_workorder.EntityLogicalName,
                    LinkToEntityName = msdyn_customerasset.EntityLogicalName,
                    LinkFromAttributeName = "msdyn_customerasset",
                    LinkToAttributeName = "msdyn_customerassetid",
                    Columns = new ColumnSet("hil_invoicedate", "msdyn_product", "msdyn_name", "hil_warrantystatus", "hil_warrantytilldate"),
                    EntityAlias = "custAsset",
                    JoinOperator = JoinOperator.Inner
                };
                QueryExpression Query = new QueryExpression(msdyn_workorder.EntityLogicalName);
                Query.ColumnSet = new ColumnSet("msdyn_name", "hil_mobilenumber", "msdyn_customerasset", "ownerid", "hil_branch", "hil_receiptamount", "msdyn_timeclosed", "createdon", "hil_actualcharges");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                //Query.Criteria.AddCondition(new ConditionExpression("hil_amcsyncstatus", ConditionOperator.Equal, 1)); //Pending Submission
                Query.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrAfter, new DateTime(2023, 1, 1))); //
                Query.Criteria.AddCondition(new ConditionExpression("hil_callsubtype", ConditionOperator.Equal, new Guid("55A71A52-3C0B-E911-A94E-000D3AF06CD4"))); //AMC Call
                Query.Criteria.AddCondition(new ConditionExpression("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"))); //Closed
                Query.Criteria.AddCondition(new ConditionExpression("hil_receiptamount", ConditionOperator.GreaterThan, 0)); //Receipt amount must be grater than 0.
                //Query.Criteria.AddCondition(new ConditionExpression("hil_branch", ConditionOperator.Equal, new Guid("FD6E00BD-BDF7-E811-A94C-000D3AF0677F"))); //DELHI Branch.
                Query.Criteria.AddCondition(new ConditionExpression("msdyn_name", ConditionOperator.Equal, "09072323022678")); //1405219429822

                Query.AddOrder("createdon", OrderType.Ascending);
                Query.LinkEntities.Add(lnkEntOwner);
                Query.LinkEntities.Add(lnkEntAddress);
                Query.LinkEntities.Add(lnkEntConsumer);
                Query.LinkEntities.Add(linkEntityForCustomerassets);
                EntityCollection ec = service.RetrieveMultiple(Query);
                msg = "Sync Started:";
                Console.WriteLine(msg);
                CreateOrUpdateLogs(containerName, fileName, msg, null);
                if (ec.Entities.Count > 0)
                {
                    AMCSAPINvoiceTableCreate _AMCSAPINvoiceTableCreate = new AMCSAPINvoiceTableCreate();
                    totalRecords = ec.Entities.Count;
                    foreach (Entity ent in ec.Entities)
                    {
                        try
                        {
                            _AMCSAPINvoiceTableCreate = new AMCSAPINvoiceTableCreate();
                            recordCnt += 1;
                            prodCode = GetMaterialCode(ent.Id, service, fileName);
                            _AMCSAPINvoiceTableCreate.AMCPlan = GetMaterialCodeRef(ent.Id, service, fileName);
                            jobId = ent.GetAttributeValue<string>("msdyn_name");
                            if (prodCode != null)
                            {
                                string _AMCPlan = prodCode.Split('|')[1]; //AMC Plan GUID
                                prodCode = prodCode.Split('|')[0]; //AMC Plan Code

                                AMCInvoiceRequest requestData = new AMCInvoiceRequest();
                                requestData.IM_DATA = new System.Collections.Generic.List<AMCInvoiceData>();
                                AMCInvoiceData invoiceData;
                                invoiceData = new AMCInvoiceData();
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
                                    Entity entTemp = service.Retrieve("hil_state", ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_state").Value).Id, new ColumnSet("hil_sapstatecode"));
                                    if (entTemp.Attributes.Contains("hil_sapstatecode"))
                                    {
                                        invoiceData.REGION_CODE = entTemp.GetAttributeValue<string>("hil_sapstatecode").ToString();
                                    }
                                }
                                invoiceData.CALL_ID = ent.GetAttributeValue<string>("msdyn_name");
                                _AMCSAPINvoiceTableCreate.CallID = invoiceData.CALL_ID;
                                if (ent.Attributes.Contains("contact.emailaddress1"))
                                {
                                    invoiceData.EMAIL = ent.GetAttributeValue<AliasedValue>("contact.emailaddress1").Value.ToString();
                                }
                                if (ent.Attributes.Contains("hil_mobilenumber"))
                                {
                                    invoiceData.PHONE = ent.GetAttributeValue<string>("hil_mobilenumber");
                                }
                                _AMCSAPINvoiceTableCreate.MobileNumber = invoiceData.PHONE;
                                if (ent.Attributes.Contains("hil_branch"))
                                {
                                    invoiceData.BRANCH_NAME = ent.GetAttributeValue<EntityReference>("hil_branch").Name;
                                }

                                invoiceData.MATERIAL_CODE = prodCode; //((EntityReference)(ent.GetAttributeValue<AliasedValue>("custAsset.msdyn_product").Value)).Name;

                                invoiceData.CS_SERVICE_PRICE = ent.GetAttributeValue<Money>("hil_receiptamount").Value.ToString();

                                if (ent.Attributes.Contains("custAsset.msdyn_name"))
                                {
                                    invoiceData.PRODUCT_SERIAL_NO = ent.GetAttributeValue<AliasedValue>("custAsset.msdyn_name").Value.ToString();
                                }
                                if (ent.Attributes.Contains("custAsset.hil_invoicedate"))
                                {
                                    string dopValue = string.Empty;
                                    dopValue = Convert.ToDateTime(ent.GetAttributeValue<AliasedValue>("custAsset.hil_invoicedate").Value).ToString("yyyy-MM-dd");
                                    invoiceData.PRODUCT_DOP = dopValue;
                                }
                                if (ent.Attributes.Contains("user.hil_employeecode"))
                                {
                                    invoiceData.EMPLOYEE_ID = ent.GetAttributeValue<AliasedValue>("user.hil_employeecode").Value.ToString();
                                }

                                if (ent.Attributes.Contains("custAsset.hil_warrantytilldate"))
                                {
                                    string dopValue = string.Empty;
                                    dopValue = Convert.ToDateTime(ent.GetAttributeValue<AliasedValue>("custAsset.hil_warrantytilldate").Value).ToString("yyyy-MM-dd");
                                    invoiceData.WARNTY_TILL_DATE = dopValue;
                                }

                                if (ent.Attributes.Contains("msdyn_timeclosed"))
                                {
                                    string jobclosedonValue = string.Empty;
                                    jobclosedonValue = ent.GetAttributeValue<DateTime>("msdyn_timeclosed").ToString("yyyy-MM-dd");
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
                                //pricingDateValue = ent.GetAttributeValue<DateTime>("createdon").ToString("yyyy-MM-dd");  // Changed on 06/Sep/2021 as discussed with Vinay & Ashish Tondon
                                pricingDateValue = ent.GetAttributeValue<DateTime>("msdyn_timeclosed").ToString("yyyy-MM-dd");
                                invoiceData.PRICING_DATE = pricingDateValue;

                                invoiceData.START_DATE = GetWarrantyStartDate(service, ent.GetAttributeValue<EntityReference>("msdyn_customerasset").Id, ent.GetAttributeValue<DateTime>("msdyn_timeclosed"), fileName);
                                _AMCSAPINvoiceTableCreate.SerialNumber = ent.GetAttributeValue<EntityReference>("msdyn_customerasset").Name;
                                if (!string.IsNullOrEmpty(invoiceData.START_DATE) && invoiceData.START_DATE.Length > 0)
                                {
                                    invoiceData.END_DATE = GetWarrantyEndDate(service, new Guid(_AMCPlan), Convert.ToDateTime(invoiceData.START_DATE), fileName);

                                    _AMCSAPINvoiceTableCreate.WarrantyStartDate = Convert.ToDateTime(invoiceData.START_DATE);
                                    _AMCSAPINvoiceTableCreate.WarrantyEndDate = Convert.ToDateTime(invoiceData.END_DATE);
                                }
                                else
                                {
                                    msg = "ERROR!!! JobId:" + invoiceData.CALL_ID + " Warranty Template Setup is not updated properly.";
                                    Console.WriteLine(msg);
                                    CreateOrUpdateLogs(containerName, fileName, msg, null);
                                    continue;
                                }
                                Decimal _receiptAmt = ent.GetAttributeValue<Money>("hil_receiptamount").Value;
                                Decimal _payableAmt = 0;

                                if (ent.Attributes.Contains("hil_actualcharges"))
                                {
                                    _payableAmt = ent.GetAttributeValue<Money>("hil_actualcharges").Value;
                                }

                                if (_payableAmt < _receiptAmt)
                                {
                                    msg = "Payable " + _payableAmt.ToString() + " and Receipt Amount " + _receiptAmt.ToString() + " is mismatch. Pending " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name");
                                    Console.WriteLine(msg);
                                    CreateOrUpdateLogs(containerName, fileName, msg, null);
                                    continue;
                                }

                                invoiceData.ZDS6_AMOUNT = (_payableAmt - _receiptAmt).ToString();
                                invoiceData.DISCOUNT_FLAG = "X";
                                requestData.IM_DATA.Add(invoiceData);

                                CreateSAPAMCRecord(service, _AMCSAPINvoiceTableCreate, fileName);

                                var Json = JsonConvert.SerializeObject(requestData);
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
                                msg = ((HttpWebResponse)response).StatusDescription;
                                CreateOrUpdateLogs(containerName, fileName, msg, null);
                                Console.WriteLine(msg);

                                dataStream = response.GetResponseStream();
                                StreamReader reader = new StreamReader(dataStream);
                                string responseFromServer = reader.ReadToEnd();
                                AMCInvoiceData resp = JsonConvert.DeserializeObject<AMCInvoiceData>(responseFromServer);
                                if (responseFromServer.Replace(@"""", "").IndexOf("STATUS:S,") > 0)
                                {
                                    Entity entUpdate = new Entity(msdyn_workorder.EntityLogicalName);
                                    entUpdate.Id = ent.Id;
                                    entUpdate["hil_amcsyncstatus"] = new OptionSetValue(2);
                                    service.Update(entUpdate);

                                    AMCSAPINvoiceTableUpdate _AMCSAPINvoiceTableUpdate = new AMCSAPINvoiceTableUpdate();
                                    _AMCSAPINvoiceTableUpdate.WarrantyStartDate = _AMCSAPINvoiceTableCreate.WarrantyStartDate;
                                    _AMCSAPINvoiceTableUpdate.WarrantyEndDate = _AMCSAPINvoiceTableCreate.WarrantyEndDate;
                                    _AMCSAPINvoiceTableUpdate.SerialNumber = _AMCSAPINvoiceTableCreate.SerialNumber;
                                    _AMCSAPINvoiceTableUpdate.MobileNumber = _AMCSAPINvoiceTableCreate.MobileNumber;
                                    _AMCSAPINvoiceTableUpdate.CallID = _AMCSAPINvoiceTableCreate.CallID;
                                    _AMCSAPINvoiceTableUpdate.invoiceStatus = (int)AMCSAPStatus.Posted;

                                    UpdateSAPAMCRecord(service, _AMCSAPINvoiceTableUpdate, fileName);
                                    msg = "AMC has been Synced with SAP Asset# " + invoiceData.PRODUCT_SERIAL_NO + " Job#" + jobId + " Closed On#" + invoiceData.CLOSED_ON + " ::: " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name");
                                    Console.WriteLine(msg);

                                    CreateOrUpdateLogs(containerName, fileName, msg, null);


                                }
                                else
                                {
                                    Entity entUpdate = new Entity(msdyn_workorder.EntityLogicalName);
                                    entUpdate.Id = ent.Id;
                                    entUpdate["hil_amcsyncstatus"] = new OptionSetValue(3);
                                    service.Update(entUpdate);
                                    msg = "Not Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name");
                                    Console.WriteLine(msg);
                                    CreateOrUpdateLogs(containerName, fileName, msg, null);
                                }
                            }
                            else
                            {
                                msg = "Product Code does not exist: Job# " + jobId + " /" + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name");
                                Console.WriteLine(msg);
                                CreateOrUpdateLogs(containerName, fileName, msg, null);
                            }
                        }
                        catch (Exception ex)
                        {
                            msg = $"{ex.Message}";
                            Console.WriteLine(msg);
                            CreateOrUpdateLogs(containerName, fileName, msg, null);
                        }
                    }
                }
                else
                {
                    msg = "AMC_Invoice_Sync.Program.Main.PushAMCData :: No record found to sync";
                    Console.WriteLine(msg);
                    CreateOrUpdateLogs(containerName, fileName, msg, null);
                }
                Console.WriteLine("Sync Ended:");

                CreateOrUpdateLogs(containerName, fileName, "Sync Ended:", null);
            }
            catch (Exception ex)
            {
                msg = "AMC_Invoice_Sync.Program.Main.PushAMCData :: Error While Loading App Settings:" + ex.Message.ToString();
                Console.WriteLine(msg);
                CreateOrUpdateLogs(containerName, fileName, msg, null);
            }
        }
        public static void PushAMCData_WP(IOrganizationService service, string fileName)
        {
            IntegrationConfig intConfig = IntegrationConfiguration(service, "amc_data", fileName);
            string authInfo = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));

            int totalRecords = 0;
            int recordCnt = 0;
            string prodCode = string.Empty;
            string jobId = string.Empty;

            try
            {
                LinkEntity lnkEntOwner = new LinkEntity
                {
                    LinkFromEntityName = "invoice",
                    LinkToEntityName = SystemUser.EntityLogicalName,
                    LinkFromAttributeName = "owninguser",
                    LinkToAttributeName = "systemuserid",
                    Columns = new ColumnSet("hil_employeecode"),
                    EntityAlias = "user",
                    JoinOperator = JoinOperator.Inner
                };
                LinkEntity lnkEntAddress = new LinkEntity
                {
                    LinkFromEntityName = "invoice",
                    LinkToEntityName = hil_address.EntityLogicalName,
                    LinkFromAttributeName = "hil_address",
                    LinkToAttributeName = "hil_addressid",
                    Columns = new ColumnSet("hil_street1", "hil_street2", "hil_street3", "hil_state", "hil_pincode", "hil_district", "hil_city", "hil_branch"),
                    EntityAlias = "address",
                    JoinOperator = JoinOperator.Inner
                };
                LinkEntity lnkEntConsumer = new LinkEntity
                {
                    LinkFromEntityName = "invoice",
                    LinkToEntityName = Contact.EntityLogicalName,
                    LinkFromAttributeName = "customerid",
                    LinkToAttributeName = "contactid",
                    Columns = new ColumnSet("hil_salutation", "fullname", "emailaddress1", "mobilephone"),
                    EntityAlias = "contact",
                    JoinOperator = JoinOperator.Inner
                };
                QueryExpression Query = new QueryExpression("invoice");
                Query.ColumnSet = new ColumnSet("hil_customerasset", "name", "hil_salestype", "hil_productcode", "totallineitemamount", "hil_receiptamount", "discountamount", "hil_productcode", "msdyn_invoicedate");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("pricelevelid", ConditionOperator.Equal, new Guid("2F3A3303-1064-ED11-9562-6045BDA55A4D"))); //FG Sale
                Query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 100001)); //Completed
                Query.Criteria.AddCondition(new ConditionExpression("hil_receiptamount", ConditionOperator.GreaterThan, 0));
                //Query.Criteria.AddCondition(new ConditionExpression("name", ConditionOperator.Equal, "INV-2023-000225")); //Invoice ID

                Query.AddOrder("msdyn_invoicedate", OrderType.Ascending);
                Query.LinkEntities.Add(lnkEntOwner);
                Query.LinkEntities.Add(lnkEntAddress);
                Query.LinkEntities.Add(lnkEntConsumer);
                EntityCollection ec = service.RetrieveMultiple(Query);
                Console.WriteLine("Sync Started:");
                if (ec.Entities.Count > 0)
                {
                    totalRecords = ec.Entities.Count;
                    foreach (Entity ent in ec.Entities)
                    {
                        try
                        {
                            recordCnt += 1;
                            prodCode = ent.GetAttributeValue<EntityReference>("hil_productcode").Name;// GetMaterialCode(ent.Id);
                            jobId = ent.GetAttributeValue<string>("name");
                            if (prodCode != null)
                            {

                                AMCInvoiceRequest requestData = new AMCInvoiceRequest();
                                requestData.IM_DATA = new System.Collections.Generic.List<AMCInvoiceData>();
                                AMCInvoiceData invoiceData;
                                invoiceData = new AMCInvoiceData();
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
                                    Entity entTemp = service.Retrieve("hil_state", ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_state").Value).Id, new ColumnSet("hil_sapstatecode"));
                                    if (entTemp.Attributes.Contains("hil_sapstatecode"))
                                    {
                                        invoiceData.REGION_CODE = entTemp.GetAttributeValue<string>("hil_sapstatecode").ToString();
                                    }
                                }

                                Query = new QueryExpression("msdyn_paymentdetail");
                                Query.ColumnSet = new ColumnSet("msdyn_name", "statuscode");
                                Query.Criteria = new FilterExpression(LogicalOperator.And);
                                Query.Criteria.AddCondition("msdyn_invoice", ConditionOperator.Equal, ent.Id);
                                Query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 910590000);
                                Query.AddOrder("createdon", OrderType.Descending);
                                EntityCollection FoundPaymentDetails = service.RetrieveMultiple(Query);
                                if (FoundPaymentDetails.Entities.Count == 0)
                                {
                                    Console.WriteLine("No Payment Found");
                                    continue;
                                }
                                string transId = FoundPaymentDetails[0].GetAttributeValue<string>("msdyn_name");
                                invoiceData.CALL_ID = transId.Replace("S365_", "");
                                //invoiceData.CALL_ID = transId;
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
                                    invoiceData.BRANCH_NAME = ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_city").Value).Name;
                                }
                                invoiceData.MATERIAL_CODE = prodCode;

                                invoiceData.CS_SERVICE_PRICE = ent.GetAttributeValue<Money>("hil_receiptamount").Value.ToString();

                                if (ent.Attributes.Contains("user.hil_employeecode"))
                                {
                                    invoiceData.EMPLOYEE_ID = ent.GetAttributeValue<AliasedValue>("user.hil_employeecode").Value.ToString();
                                }
                                if (ent.Attributes.Contains("msdyn_invoicedate"))
                                {
                                    string jobclosedonValue = string.Empty;
                                    jobclosedonValue = ent.GetAttributeValue<DateTime>("msdyn_invoicedate").ToString("yyyy-MM-dd");
                                    invoiceData.CLOSED_ON = jobclosedonValue;
                                }
                                invoiceData.WARNTY_STATUS = "OUT";
                                string pricingDateValue = string.Empty;
                                pricingDateValue = ent.GetAttributeValue<DateTime>("msdyn_invoicedate").ToString("yyyy-MM-dd");
                                invoiceData.PRICING_DATE = pricingDateValue;


                                Decimal _receiptAmt = ent.GetAttributeValue<Money>("hil_receiptamount").Value;
                                Decimal _payableAmt = 0;

                                if (ent.Attributes.Contains("totallineitemamount"))
                                {
                                    _payableAmt = ent.GetAttributeValue<Money>("totallineitemamount").Value;
                                }

                                if (_payableAmt < _receiptAmt) continue;

                                invoiceData.ZDS6_AMOUNT = (_payableAmt - _receiptAmt).ToString();
                                invoiceData.DISCOUNT_FLAG = "X";

                                if (ent.Contains("hil_customerasset"))
                                {
                                    invoiceData.PRODUCT_SERIAL_NO = ent.GetAttributeValue<EntityReference>("hil_customerasset").Name;
                                    invoiceData.PRODUCT_DOP = "";
                                }
                                //DateTime _AMCStartDate = DateTime.Now;
                                //DateTime _AMCEndDate = _AMCStartDate.AddDays(365);
                                //invoiceData.START_DATE = _AMCStartDate.Year.ToString() + "-0" + _AMCStartDate.Month.ToString() + "-" + _AMCStartDate.Day.ToString();
                                //invoiceData.END_DATE = _AMCEndDate.Year.ToString() + "-0" + _AMCEndDate.Month.ToString() + "-" + _AMCEndDate.Day.ToString(); ;

                                requestData.IM_DATA.Add(invoiceData);

                                var Json = JsonConvert.SerializeObject(requestData);

                                WebRequest request = WebRequest.Create(intConfig.uri);
                                request.Headers[HttpRequestHeader.Authorization] = authInfo;
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
                                AMCInvoiceData resp = JsonConvert.DeserializeObject<AMCInvoiceData>(responseFromServer);
                                if (responseFromServer.Replace(@"""", "").IndexOf("STATUS:S,") > 0)
                                {
                                    SetStateRequest req = new SetStateRequest();
                                    req.State = new OptionSetValue(2);
                                    req.Status = new OptionSetValue(910590000);
                                    req.EntityMoniker = ent.ToEntityReference();
                                    var res = (SetStateResponse)service.Execute(req);
                                    Console.WriteLine("Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                                }
                                else
                                {
                                    Console.WriteLine("Not Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                                }
                            }
                            else
                            {
                                Console.WriteLine("Product Code does not exist: Job# " + jobId + " /" + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
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
                    Console.WriteLine("AMC_Invoice_Sync.Program.Main.PushAMCData_WP :: No record found to sync");
                }
                Console.WriteLine("Sync Ended:");
            }
            catch (Exception ex)
            {
                Console.WriteLine("AMC_Invoice_Sync.Program.Main.PushAMCData_WP :: Error While Loading App Settings:" + ex.Message.ToString());
            }
        }

        static void updatePaymentStatusforInvoice(String transactionID, EntityReference Invoice, IOrganizationService service, string URL, string authInfo, string fileName)
        {
            try
            {
                String StatusPay = "";
                StatusRequest req = new StatusRequest();
                req.PROJECT = "D365";
                req.command = "verify_payment";
                req.var1 = transactionID;
                var client = new RestClient(URL);
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);

                request.AddHeader("Authorization", authInfo);
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("application/json", Newtonsoft.Json.JsonConvert.SerializeObject(req), ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);

                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<StatusResponse>(response.Content);
                QueryExpression Query = new QueryExpression("hil_paymentstatus");
                Query.ColumnSet = new ColumnSet("hil_job", "hil_name");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, transactionID);
                EntityCollection Found = service.RetrieveMultiple(Query);
                foreach (var item in obj.transaction_details)
                {
                    Entity statusPayment = new Entity("hil_paymentstatus");
                    statusPayment.Id = Found[0].Id;
                    statusPayment["hil_mihpayid"] = item.mihpayid;
                    statusPayment["hil_request_id"] = item.request_id;
                    statusPayment["hil_bank_ref_num"] = item.bank_ref_num;

                    statusPayment["hil_amt"] = item.amt;
                    statusPayment["hil_transaction_amount"] = item.transaction_amount;
                    //statusPayment.txnid = item.txnid;
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
                    service.Update(statusPayment);
                    StatusPay = item.status.ToLower();
                }
                Console.WriteLine(StatusPay + " Payment Status of transactionid " + transactionID);
                if (StatusPay.ToLower() == "success".ToLower())
                {
                    SetStateRequest req2 = new SetStateRequest();
                    req2.State = new OptionSetValue(2);
                    req2.Status = new OptionSetValue(100001);
                    req2.EntityMoniker = Invoice;
                    var res = (SetStateResponse)service.Execute(req2);

                    Query = new QueryExpression("msdyn_paymentdetail");
                    Query.ColumnSet = new ColumnSet(false);
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, transactionID);
                    EntityCollection FoundPaymentDetails = service.RetrieveMultiple(Query);

                    req2 = new SetStateRequest();
                    req2.State = new OptionSetValue(0);
                    req2.Status = new OptionSetValue(910590000);
                    req2.EntityMoniker = FoundPaymentDetails[0].ToEntityReference();
                    res = (SetStateResponse)service.Execute(req2);
                }
                else if (StatusPay.ToLower() == "failure".ToLower())
                {
                    Query = new QueryExpression("msdyn_paymentdetail");
                    Query.ColumnSet = new ColumnSet(false);
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, transactionID);
                    EntityCollection FoundPaymentDetails = service.RetrieveMultiple(Query);

                    Entity entity = new Entity(FoundPaymentDetails[0].LogicalName, FoundPaymentDetails[0].Id);
                    entity["msdyn_paymentamount"] = new Money(0);
                    service.Update(entity);

                    SetStateRequest req1 = new SetStateRequest();
                    req1.State = new OptionSetValue(0);
                    req1.Status = new OptionSetValue(910590001);
                    req1.EntityMoniker = FoundPaymentDetails[0].ToEntityReference();
                    service.Execute(req1);
                }
                else if (StatusPay.ToLower() == "not found".ToLower())
                {
                    Query = new QueryExpression("msdyn_paymentdetail");
                    Query.ColumnSet = new ColumnSet(false);
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, transactionID);
                    //EntityCollection FoundPaymentDetails = service.RetrieveMultiple(Query);
                    //SetStateRequest req1 = new SetStateRequest();
                    //req1.State = new OptionSetValue(0);
                    //req1.Status = new OptionSetValue(910590001);
                    //req1.EntityMoniker = FoundPaymentDetails[0].ToEntityReference();
                    //service.Execute(req1);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("D365 Internal Error " + ex);
            }
        }
        static string updatePaymentStatusforJob(String transactionID, IOrganizationService service, string URL, string authInfo, string fileName)
        {
            string payStatus = string.Empty;
            try
            {
                StatusRequest req = new StatusRequest();
                req.PROJECT = "D365";
                req.command = "verify_payment";
                req.var1 = transactionID;
                var client = new RestClient(URL);
                client.Timeout = -1;
                var request = new RestRequest(Method.POST);

                request.AddHeader("Authorization", authInfo);
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("application/json", Newtonsoft.Json.JsonConvert.SerializeObject(req), ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);

                var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<StatusResponse>(response.Content);
                QueryExpression Query = new QueryExpression("hil_paymentstatus");
                Query.ColumnSet = new ColumnSet("hil_job", "hil_name");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, transactionID);
                EntityCollection Found = service.RetrieveMultiple(Query);
                string bank_ref_number = string.Empty;
                if (Found.Entities.Count > 0)
                {
                    Entity statusPayment = new Entity("hil_paymentstatus");
                    statusPayment.Id = Found[0].Id;
                    if (obj.transaction_details[0].mihpayid != null)
                    {
                        statusPayment["hil_mihpayid"] = obj.transaction_details[0].mihpayid.ToString();
                    }
                    if (obj.transaction_details[0].request_id != null)
                    {
                        statusPayment["hil_request_id"] = obj.transaction_details[0].request_id;
                    }
                    if (obj.transaction_details[0].bank_ref_num != null)
                    {
                        statusPayment["hil_bank_ref_num"] = obj.transaction_details[0].bank_ref_num;
                        bank_ref_number = obj.transaction_details[0].bank_ref_num;
                    }
                    if (obj.transaction_details[0].amt != null)
                    {
                        statusPayment["hil_amt"] = obj.transaction_details[0].amt;
                    }
                    if (obj.transaction_details[0].transaction_amount != null)
                    {
                        statusPayment["hil_transaction_amount"] = obj.transaction_details[0].transaction_amount;
                    }
                    if (obj.transaction_details[0].additional_charges != null)
                    {
                        statusPayment["hil_additional_charges"] = obj.transaction_details[0].additional_charges;
                    }
                    if (obj.transaction_details[0].productinfo != null)
                    {
                        statusPayment["hil_productinfo"] = obj.transaction_details[0].productinfo;
                    }
                    if (obj.transaction_details[0].firstname != null)
                    {
                        statusPayment["hil_firstname"] = obj.transaction_details[0].firstname;
                    }
                    if (obj.transaction_details[0].bankcode != null)
                    {
                        statusPayment["hil_bankcode"] = obj.transaction_details[0].bankcode;
                    }
                    if (obj.transaction_details[0].udf1 != null)
                    {
                        statusPayment["hil_udf1"] = obj.transaction_details[0].udf1;
                    }
                    if (obj.transaction_details[0].udf2 != null)
                    {
                        statusPayment["hil_udf2"] = obj.transaction_details[0].udf2;
                    }
                    if (obj.transaction_details[0].udf3 != null)
                    {
                        statusPayment["hil_udf3"] = obj.transaction_details[0].udf3;
                    }
                    if (obj.transaction_details[0].udf4 != null)
                    {
                        statusPayment["hil_udf4"] = obj.transaction_details[0].udf4;
                    }
                    if (obj.transaction_details[0].udf5 != null)
                    {
                        statusPayment["hil_udf5"] = obj.transaction_details[0].udf5;
                    }
                    if (obj.transaction_details[0].field2 != null)
                    {
                        statusPayment["hil_field2"] = obj.transaction_details[0].field2;
                    }
                    if (obj.transaction_details[0].field9 != null)
                    {
                        statusPayment["hil_field9"] = obj.transaction_details[0].field9.Length > 100 ? obj.transaction_details[0].field9.Substring(0, 99) : obj.transaction_details[0].field9;
                    }
                    if (obj.transaction_details[0].error_code != null)
                    {
                        statusPayment["hil_error_code"] = obj.transaction_details[0].error_code;
                    }
                    if (obj.transaction_details[0].addedon != null)
                    {
                        statusPayment["hil_addedon"] = obj.transaction_details[0].addedon;
                    }
                    if (obj.transaction_details[0].payment_source != null)
                    {
                        statusPayment["hil_payment_source"] = obj.transaction_details[0].payment_source;
                    }
                    if (obj.transaction_details[0].card_type != null)
                    {
                        statusPayment["hil_card_type"] = obj.transaction_details[0].card_type;
                    }
                    if (obj.transaction_details[0].error_Message != null)
                    {
                        statusPayment["hil_error_message"] = obj.transaction_details[0].error_Message;
                    }
                    if (obj.transaction_details[0].net_amount_debit != null)
                    {
                        statusPayment["hil_net_amount_debit"] = obj.transaction_details[0].net_amount_debit;
                    }
                    if (obj.transaction_details[0].disc != null)
                    {
                        statusPayment["hil_disc"] = obj.transaction_details[0].disc;
                    }
                    if (obj.transaction_details[0].mode != null)
                    {
                        statusPayment["hil_mode"] = obj.transaction_details[0].mode;
                    }
                    if (obj.transaction_details[0].PG_TYPE != null)
                    {
                        statusPayment["hil_pg_type"] = obj.transaction_details[0].PG_TYPE;
                    }
                    if (obj.transaction_details[0].card_no != null)
                    {
                        statusPayment["hil_card_no"] = obj.transaction_details[0].card_no;
                    }
                    if (obj.transaction_details[0].status != null)
                    {
                        statusPayment["hil_paymentstatus"] = obj.transaction_details[0].status;
                    }
                    if (obj.transaction_details[0].unmappedstatus != null)
                    {
                        statusPayment["hil_unmappedstatus"] = obj.transaction_details[0].unmappedstatus;
                    }
                    if (obj.transaction_details[0].Merchant_UTR != null)
                    {
                        statusPayment["hil_merchant_utr"] = obj.transaction_details[0].Merchant_UTR;
                    }
                    if (obj.transaction_details[0].Settled_At != null)
                    {
                        statusPayment["hil_settled_at"] = obj.transaction_details[0].Settled_At;
                    }
                    service.Update(statusPayment);

                    if (Found[0].Contains("hil_job"))
                    {
                        int paymentstatus = 0;

                        if (obj.transaction_details[0].status == "Not Found")
                        {
                            paymentstatus = (1);
                        }
                        else if (obj.transaction_details[0].status == "success")
                        {
                            paymentstatus = (2);
                        }
                        else if (obj.transaction_details[0].status == "pending")
                        {
                            paymentstatus = (3);
                        }
                        else
                            paymentstatus = (4);
                        payStatus = obj.transaction_details[0].status;
                        updatePaymentinJob(service, Found[0].GetAttributeValue<EntityReference>("hil_job").Id, paymentstatus, bank_ref_number, fileName);
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("D365 Internal Error " + ex);
            }
            return payStatus;
        }
        static void updatePaymentinJob(IOrganizationService service, Guid JobId, int paymentstatus, string bank_ref_number, string fileName)
        {
            #region Updating Job Payment Status field

            Entity _updateJob = new Entity("msdyn_workorder", JobId);
            _updateJob["hil_paymentstatus"] = new OptionSetValue(paymentstatus);
            _updateJob["hil_receiptnumber"] = bank_ref_number;
            service.Update(_updateJob);

            #endregion
        }
        public static IntegrationConfig IntegrationConfiguration(IOrganizationService service, string Param, string fileName)
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
                //if (output.uri.Contains("middleware.havells.com"))
                //{
                //    output.uri = output.uri.Replace("middleware", "p90ci");
                //}
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
            return output;
        }
        #region AMC Start And ENd Date Calculation
        static string GetWarrantyStartDate(IOrganizationService service, Guid AssetID, DateTime _purchaseDate, string fileName)
        {
            string WarrantyStartDate = _purchaseDate.ToString("yyyy-MM-dd");
            try
            {
                LinkEntity lnkEntInvoice = new LinkEntity
                {
                    LinkFromEntityName = hil_unitwarranty.EntityLogicalName,
                    LinkToEntityName = hil_warrantytemplate.EntityLogicalName,
                    LinkFromAttributeName = "hil_warrantytemplate",
                    LinkToAttributeName = "hil_warrantytemplateid",
                    Columns = new ColumnSet("hil_type"),
                    EntityAlias = "invoice",
                    JoinOperator = JoinOperator.Inner
                };
                QueryExpression Query = new QueryExpression(hil_unitwarranty.EntityLogicalName);
                Query.ColumnSet = new ColumnSet("hil_name", "hil_warrantyenddate", "hil_warrantytemplate");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("hil_customerasset", ConditionOperator.Equal, AssetID));
                Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));  //Active
                Query.AddOrder("hil_warrantyenddate", OrderType.Descending);
                Query.LinkEntities.Add(lnkEntInvoice);
                EntityCollection ec = service.RetrieveMultiple(Query);
                if (ec.Entities.Count > 0)
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
                            EntityCollection ec1 = service.RetrieveMultiple(new FetchExpression(fetch));
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
                    Query = new QueryExpression(hil_unitwarranty.EntityLogicalName);
                    Query.ColumnSet = new ColumnSet("hil_warrantystartdate", "hil_warrantytemplate");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition(new ConditionExpression("hil_customerasset", ConditionOperator.Equal, AssetID));
                    Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));  //Active
                    EntityCollection ec1 = service.RetrieveMultiple(Query);
                    if (ec1.Entities.Count == 0)
                    {
                        //WarrantyStartDate = string.Empty;
                        Console.WriteLine("ERROR!!! No Standard Warranty is found with Asset.");
                    }
                    else
                    {
                        Console.WriteLine("ERROR!!! Warranty Template is not defined.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return WarrantyStartDate;
        }
        static string GetWarrantyEndDate(IOrganizationService service, Guid _AMCPlaneID, DateTime StartDate, string fileName)
        {
            string WarrantyEndDate = null;
            QueryExpression Query = new QueryExpression(hil_warrantytemplate.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("hil_warrantyperiod");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("hil_amcplan", ConditionOperator.Equal, _AMCPlaneID));
            Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));  //Active
            Query.TopCount = 1;
            Query.AddOrder("createdon", OrderType.Descending);
            EntityCollection ec = service.RetrieveMultiple(Query);
            if (ec.Entities.Count == 1)
            {
                WarrantyEndDate = StartDate.AddMonths(ec[0].GetAttributeValue<int>("hil_warrantyperiod")).AddDays(-1).ToString("yyyy-MM-dd");
            }
            return WarrantyEndDate;
        }
        #endregion
        static string GetMaterialCode(Guid jobGuId, IOrganizationService service, string fileName)
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
            EntityCollection ec = service.RetrieveMultiple(new FetchExpression(fetchXML));
            if (ec.Entities.Count > 0)
            {
                if (ec.Entities[0].Attributes.Contains("prod.productnumber"))
                {
                    productCode = ec.Entities[0].GetAttributeValue<AliasedValue>("prod.productnumber").Value.ToString() + "|" + ec.Entities[0].GetAttributeValue<EntityReference>("hil_replacedpart").Id.ToString();
                }
            }
            return productCode;
        }
        static EntityReference GetMaterialCodeRef(Guid jobGuId, IOrganizationService service, string fileName)
        {
            EntityReference productCode = null;
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
            EntityCollection ec = service.RetrieveMultiple(new FetchExpression(fetchXML));
            if (ec.Entities.Count > 0)
            {
                if (ec.Entities[0].Attributes.Contains("prod.productnumber"))
                {
                    productCode = ec.Entities[0].GetAttributeValue<EntityReference>("hil_replacedpart");
                }
            }
            return productCode;
        }
        static void CreateSAPAMCRecord(IOrganizationService service, AMCSAPINvoiceTableCreate _AMCSAPINvoiceTableCreate, string fileName)
        {
            Entity entity = new Entity("hil_amcstaging");
            entity["hil_callid"] = _AMCSAPINvoiceTableCreate.CallID;
            entity["hil_serailnumber"] = _AMCSAPINvoiceTableCreate.SerialNumber;
            entity["hil_mobilenumber"] = _AMCSAPINvoiceTableCreate.MobileNumber;
            entity["hil_warrantystartdate"] = _AMCSAPINvoiceTableCreate.WarrantyStartDate;
            entity["hil_warrantyenddate"] = _AMCSAPINvoiceTableCreate.WarrantyEndDate;
            entity["hil_amcplan"] = _AMCSAPINvoiceTableCreate.AMCPlan;
            entity["hil_amcinvoicestatus"] = new OptionSetValue((int)AMCSAPStatus.Draft);
            QueryExpression Query = new QueryExpression("hil_amcstaging");
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("hil_callid", ConditionOperator.Equal, _AMCSAPINvoiceTableCreate.CallID));
            Query.Criteria.AddCondition(new ConditionExpression("hil_serailnumber", ConditionOperator.Equal, _AMCSAPINvoiceTableCreate.SerialNumber));
            Query.Criteria.AddCondition(new ConditionExpression("hil_mobilenumber", ConditionOperator.Equal, _AMCSAPINvoiceTableCreate.MobileNumber));
            Query.Criteria.AddCondition(new ConditionExpression("hil_warrantystartdate", ConditionOperator.Equal, _AMCSAPINvoiceTableCreate.WarrantyStartDate));
            Query.Criteria.AddCondition(new ConditionExpression("hil_warrantyenddate", ConditionOperator.Equal, _AMCSAPINvoiceTableCreate.WarrantyEndDate));
            EntityCollection ec = service.RetrieveMultiple(Query);
            if (ec.Entities.Count == 1)
            {
                entity.Id = ec.Entities[0].Id;
                service.Update(entity);
            }
            else if (ec.Entities.Count == 0)
            {
                service.Create(entity);
            }
        }
        static void UpdateSAPAMCRecord(IOrganizationService service, AMCSAPINvoiceTableUpdate _AMCSAPINvoiceTableUpdate, string fileName)
        {
            QueryExpression Query = new QueryExpression("hil_amcstaging");
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("hil_callid", ConditionOperator.Equal, _AMCSAPINvoiceTableUpdate.CallID));
            Query.Criteria.AddCondition(new ConditionExpression("hil_serailnumber", ConditionOperator.Equal, _AMCSAPINvoiceTableUpdate.SerialNumber));
            Query.Criteria.AddCondition(new ConditionExpression("hil_mobilenumber", ConditionOperator.Equal, _AMCSAPINvoiceTableUpdate.MobileNumber));
            Query.Criteria.AddCondition(new ConditionExpression("hil_warrantystartdate", ConditionOperator.Equal, _AMCSAPINvoiceTableUpdate.WarrantyStartDate));
            Query.Criteria.AddCondition(new ConditionExpression("hil_warrantyenddate", ConditionOperator.Equal, _AMCSAPINvoiceTableUpdate.WarrantyEndDate));
            EntityCollection ec = service.RetrieveMultiple(Query);
            if (ec.Entities.Count == 1)
            {
                Entity entity = new Entity(ec[0].LogicalName, ec[0].Id);
                if (_AMCSAPINvoiceTableUpdate.SAPBillDoc != null)
                    entity["hil_name"] = _AMCSAPINvoiceTableUpdate.SAPBillDoc;
                if (_AMCSAPINvoiceTableUpdate.SAPBillDate != null)
                    entity["hil_sapbillingdate"] = _AMCSAPINvoiceTableUpdate.SAPBillDate;
                if (_AMCSAPINvoiceTableUpdate.invoiceStatus != null)
                    entity["hil_amcinvoicestatus"] = new OptionSetValue((int)_AMCSAPINvoiceTableUpdate.invoiceStatus);
                if (_AMCSAPINvoiceTableUpdate.AMCPlanName != null)
                    entity["hil_amcplannameslt"] = _AMCSAPINvoiceTableUpdate.AMCPlanName;
                service.Update(entity);
            }
        }

    }
}
