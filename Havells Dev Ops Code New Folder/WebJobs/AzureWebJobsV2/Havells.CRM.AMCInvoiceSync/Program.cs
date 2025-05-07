using System;
using Microsoft.Xrm.Sdk;
using System.IO;
using Microsoft.Xrm.Sdk.Query;
using System.Text;
using Microsoft.Xrm.Tooling.Connector;
using Newtonsoft.Json;
using System.Net;
using Havells.CRM.AMCInvoiceSync.Model;
using System.Web.Services.Description;
using Microsoft.Crm.Sdk.Messages;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;

namespace Havells.CRM.AMCInvoiceSync
{
    public class Program
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion

        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                //GetPendingPaymentStatus.getPaymentStatus(_service);
                GetPendingPaymentStatus.getPaymentStatusofInvoice(_service);
                //Program.CreateAMCDocument_Invoice(_service);
                //PushAMCData_DEV(); //AMC Omnichannel
                //PushAMCData(); // AMC Selling Job
                //PushAMCData_WP(); // Waterpurifier Upselling Finished Goods
                PushAMCData_OmniChannel();
            }
            else
            {

            }
        }
        static void PushAMCData_DEV()
        {
            string sUserName = "D365_Havells";
            string sPassword = "PRDD365@1234";
            //string sPassword = "QAD365@1234";
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
                Query.ColumnSet = new ColumnSet("msdyn_name", "hil_mobilenumber", "ownerid", "hil_branch", "hil_receiptamount", "msdyn_timeclosed", "createdon", "hil_actualcharges");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("hil_amcsyncstatus", ConditionOperator.Equal, 1)); //Pending Submission
                Query.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrAfter, new DateTime(2024, 01, 01))); //
                Query.Criteria.AddCondition(new ConditionExpression("hil_callsubtype", ConditionOperator.Equal, new Guid("55A71A52-3C0B-E911-A94E-000D3AF06CD4"))); //AMC Call
                Query.Criteria.AddCondition(new ConditionExpression("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"))); //Closed
                Query.Criteria.AddCondition(new ConditionExpression("hil_receiptamount", ConditionOperator.GreaterThan, 0)); //Receipt amount must be grater than 0.
                //Query.Criteria.AddCondition(new ConditionExpression("hil_branch", ConditionOperator.Equal, new Guid("FD6E00BD-BDF7-E811-A94C-000D3AF0677F"))); //DELHI Branch.
                //Query.Criteria.AddCondition(new ConditionExpression("msdyn_name", ConditionOperator.Equal, "20072215989529")); //1405219429822

                Query.AddOrder("createdon", OrderType.Ascending);
                Query.LinkEntities.Add(lnkEntOwner);
                Query.LinkEntities.Add(lnkEntAddress);
                Query.LinkEntities.Add(lnkEntConsumer);
                Query.LinkEntities.Add(linkEntityForCustomerassets);
                EntityCollection ec = _service.RetrieveMultiple(Query);
                Console.WriteLine("Sync Started:");
                if (ec.Entities.Count > 0)
                {
                    totalRecords = ec.Entities.Count;
                    foreach (Entity ent in ec.Entities)
                    {
                        recordCnt += 1;
                        prodCode = GetMaterialCode(ent.Id);
                        jobId = ent.GetAttributeValue<string>("msdyn_name");
                        if (prodCode != null)
                        {
                            prodCode = prodCode.Split('|')[0]; //AMC Plan Code
                            string _AMCPlan = prodCode.Split('|')[1]; //AMC Plan GUID
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

                            Decimal _receiptAmt = ent.GetAttributeValue<Money>("hil_receiptamount").Value;
                            Decimal _payableAmt = 0;

                            if (ent.Attributes.Contains("hil_actualcharges"))
                            {
                                _payableAmt = ent.GetAttributeValue<Money>("hil_actualcharges").Value;
                            }

                            if (_payableAmt < _receiptAmt) { Console.WriteLine("Pending " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name")); continue; }

                            invoiceData.ZDS6_AMOUNT = (_payableAmt - _receiptAmt).ToString();
                            invoiceData.DISCOUNT_FLAG = "X";

                            invoiceData.START_DATE = "20230215";
                            invoiceData.END_DATE = "20240214";

                            requestData.IM_DATA.Add(invoiceData);

                            var Json = JsonConvert.SerializeObject(requestData);
                            WebRequest request = WebRequest.Create("https://middleware.havells.com:50001/RESTAdapter/dynamics/amc_data");
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
                            AMCInvoiceData resp = JsonConvert.DeserializeObject<AMCInvoiceData>(responseFromServer);
                            if (responseFromServer.Replace(@"""", "").IndexOf("STATUS:S,") > 0)
                            {
                                Entity entUpdate = new Entity(msdyn_workorder.EntityLogicalName);
                                entUpdate.Id = ent.Id;
                                entUpdate["hil_amcsyncstatus"] = new OptionSetValue(2);
                                _service.Update(entUpdate);
                                Console.WriteLine("Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                            }
                            else
                            {
                                Entity entUpdate = new Entity(msdyn_workorder.EntityLogicalName);
                                entUpdate.Id = ent.Id;
                                entUpdate["hil_amcsyncstatus"] = new OptionSetValue(3);
                                _service.Update(entUpdate);
                                Console.WriteLine("Not Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                            }
                        }
                        else
                        {
                            Console.WriteLine("Product Code does not exist: Invoice " + jobId + " /" + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
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
        static void PushAMCData()
        {
            string sUserName = "D365_Havells";
            string sPassword = "PRDD365@1234";
            //string sPassword = "QAD365@1234";
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
                Query.Criteria.AddCondition(new ConditionExpression("hil_amcsyncstatus", ConditionOperator.Equal, 1)); //Pending Submission
                Query.Criteria.AddCondition(new ConditionExpression("msdyn_timeclosed", ConditionOperator.OnOrAfter, new DateTime(2023, 4, 4))); //
                Query.Criteria.AddCondition(new ConditionExpression("hil_callsubtype", ConditionOperator.Equal, new Guid("55A71A52-3C0B-E911-A94E-000D3AF06CD4"))); //AMC Call
                Query.Criteria.AddCondition(new ConditionExpression("msdyn_substatus", ConditionOperator.Equal, new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"))); //Closed
                Query.Criteria.AddCondition(new ConditionExpression("hil_receiptamount", ConditionOperator.GreaterThan, 0)); //Receipt amount must be grater than 0.

                Query.AddOrder("createdon", OrderType.Ascending);
                Query.LinkEntities.Add(lnkEntOwner);
                Query.LinkEntities.Add(lnkEntAddress);
                Query.LinkEntities.Add(lnkEntConsumer);
                Query.LinkEntities.Add(linkEntityForCustomerassets);
                EntityCollection ec = _service.RetrieveMultiple(Query);
                Console.WriteLine("Sync Started:");
                if (ec.Entities.Count > 0)
                {
                    totalRecords = ec.Entities.Count;
                    foreach (Entity ent in ec.Entities)
                    {
                        recordCnt += 1;
                        prodCode = GetMaterialCode(ent.Id);
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

                            invoiceData.START_DATE = GetWarrantyStartDate(_service, ent.GetAttributeValue<EntityReference>("msdyn_customerasset").Id, ent.GetAttributeValue<DateTime>("msdyn_timeclosed"));
                            if (!string.IsNullOrEmpty(invoiceData.START_DATE) && invoiceData.START_DATE.Length > 0)
                            {
                                invoiceData.END_DATE = GetWarrantyEndDate(_service, new Guid(_AMCPlan), Convert.ToDateTime(invoiceData.START_DATE));
                            }
                            else
                            {
                                Console.WriteLine("ERROR!!! JobId:" + invoiceData.CALL_ID + " Warranty Template Setup is not updated properly.");
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
                                Console.WriteLine("Payable " + _payableAmt.ToString() + " and Receipt Amount " + _receiptAmt.ToString() + " is mismatch. Pending " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                                continue;
                            }

                            invoiceData.ZDS6_AMOUNT = (_payableAmt - _receiptAmt).ToString();
                            invoiceData.DISCOUNT_FLAG = "X";
                            requestData.IM_DATA.Add(invoiceData);

                            var Json = JsonConvert.SerializeObject(requestData);
                            WebRequest request = WebRequest.Create("https://middleware.havells.com:50001/RESTAdapter/dynamics/amc_data");
                            //WebRequest request = WebRequest.Create(https://p90ci.havells.com:50001/RESTAdapter/dynamics/amc_data);
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
                            AMCInvoiceData resp = JsonConvert.DeserializeObject<AMCInvoiceData>(responseFromServer);
                            if (responseFromServer.Replace(@"""", "").IndexOf("STATUS:S,") > 0)
                            {
                                Entity entUpdate = new Entity(msdyn_workorder.EntityLogicalName);
                                entUpdate.Id = ent.Id;
                                entUpdate["hil_amcsyncstatus"] = new OptionSetValue(2);
                                _service.Update(entUpdate);
                                Console.WriteLine("AMC has been Synced with SAP Asset# " + invoiceData.PRODUCT_SERIAL_NO + " Job#" + jobId + " Closed On#" + invoiceData.CLOSED_ON + " ::: " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                            }
                            else
                            {
                                Entity entUpdate = new Entity(msdyn_workorder.EntityLogicalName);
                                entUpdate.Id = ent.Id;
                                entUpdate["hil_amcsyncstatus"] = new OptionSetValue(3);
                                _service.Update(entUpdate);
                                Console.WriteLine("Not Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                            }
                        }
                        else
                        {
                            Console.WriteLine("Product Code does not exist: Job# " + jobId + " /" + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
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
        #region AMC Start And ENd Date Calculation
        static string GetWarrantyStartDatenew(IOrganizationService service, Guid AssetID, DateTime _purchaseDate)
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
                    Query = new QueryExpression(hil_unitwarranty.EntityLogicalName);
                    Query.ColumnSet = new ColumnSet("hil_warrantystartdate", "hil_warrantytemplate");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition(new ConditionExpression("hil_customerasset", ConditionOperator.Equal, AssetID));
                    //Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));  //Active
                    EntityCollection ec1 = service.RetrieveMultiple(Query);
                    if (ec1.Entities.Count == 0)
                    {
                        WarrantyStartDate = string.Empty;
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
        static string GetWarrantyEndDatenew(IOrganizationService service, Guid _AMCPlaneID, DateTime StartDate)
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

        //static void PushAMCData_WP()
        //{
        //    string sUserName = "D365_Havells";
        //    //string sPassword = "DEVD365@1234";
        //    //string sPassword = "QAD365@1234";
        //    string sPassword = "PRDD365@1234";
        //    int totalRecords = 0;
        //    int recordCnt = 0;
        //    string prodCode = string.Empty;
        //    string jobId = string.Empty;

        //    try
        //    {
        //        LinkEntity lnkEntOwner = new LinkEntity
        //        {
        //            LinkFromEntityName = "invoice",
        //            LinkToEntityName = SystemUser.EntityLogicalName,
        //            LinkFromAttributeName = "owninguser",
        //            LinkToAttributeName = "systemuserid",
        //            Columns = new ColumnSet("hil_employeecode"),
        //            EntityAlias = "user",
        //            JoinOperator = JoinOperator.Inner
        //        };
        //        LinkEntity lnkEntAddress = new LinkEntity
        //        {
        //            LinkFromEntityName = "invoice",
        //            LinkToEntityName = hil_address.EntityLogicalName,
        //            LinkFromAttributeName = "hil_address",
        //            LinkToAttributeName = "hil_addressid",
        //            Columns = new ColumnSet("hil_street1", "hil_street2", "hil_street3", "hil_state", "hil_pincode", "hil_district", "hil_city", "hil_branch"),
        //            EntityAlias = "address",
        //            JoinOperator = JoinOperator.Inner
        //        };
        //        LinkEntity lnkEntConsumer = new LinkEntity
        //        {
        //            LinkFromEntityName = "invoice",
        //            LinkToEntityName = Contact.EntityLogicalName,
        //            LinkFromAttributeName = "customerid",
        //            LinkToAttributeName = "contactid",
        //            Columns = new ColumnSet("hil_salutation", "fullname", "emailaddress1", "mobilephone"),
        //            EntityAlias = "contact",
        //            JoinOperator = JoinOperator.Inner
        //        };

        //        QueryExpression Query = new QueryExpression("invoice");
        //        Query.ColumnSet = new ColumnSet("name","hil_salestype", "hil_productcode","totallineitemamount", "hil_receiptamount", "discountamount", "hil_productcode", "msdyn_invoicedate");
        //        Query.Criteria = new FilterExpression(LogicalOperator.And);
        //        //Query.Criteria.AddCondition(new ConditionExpression("msdyn_invoicedate", ConditionOperator.OnOrAfter, new DateTime(2022, 2, 04))); //
        //        Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 2)); //Posted
        //        Query.Criteria.AddCondition(new ConditionExpression("hil_receiptamount", ConditionOperator.GreaterThan, 0)); //Receipt amount must be grater than 0.
        //        //Query.Criteria.AddCondition(new ConditionExpression("name", ConditionOperator.Equal, "INV-2022-000026")); //Invoice ID

        //        Query.AddOrder("msdyn_invoicedate", OrderType.Ascending);
        //        Query.LinkEntities.Add(lnkEntOwner);
        //        Query.LinkEntities.Add(lnkEntAddress);
        //        Query.LinkEntities.Add(lnkEntConsumer);
        //        //Query.LinkEntities.Add(linkEntityForCustomerassets);
        //        EntityCollection ec = _service.RetrieveMultiple(Query);
        //        Console.WriteLine("Sync Started:");
        //        if (ec.Entities.Count > 0)
        //        {
        //            totalRecords = ec.Entities.Count;
        //            foreach (Entity ent in ec.Entities)
        //            {
        //                recordCnt += 1;
        //                prodCode = ent.GetAttributeValue<EntityReference>("hil_productcode").Name;// GetMaterialCode(ent.Id);
        //                jobId = ent.GetAttributeValue<string>("name");
        //                if (prodCode != null)
        //                {
        //                    AMCInvoiceRequest requestData = new AMCInvoiceRequest();
        //                    requestData.IM_DATA = new System.Collections.Generic.List<AMCInvoiceData>();
        //                    AMCInvoiceData invoiceData;
        //                    invoiceData = new AMCInvoiceData();
        //                    if (ent.Attributes.Contains("contact.fullname"))
        //                    {
        //                        invoiceData.CUSTOMER_NAME = ent.GetAttributeValue<AliasedValue>("contact.fullname").Value.ToString();
        //                    }
        //                    if (ent.Attributes.Contains("contact.hil_salutation"))
        //                    {
        //                        invoiceData.TITLE = "Mr.";
        //                    }
        //                    if (ent.Attributes.Contains("address.hil_street1"))
        //                    {
        //                        invoiceData.STREET = ent.GetAttributeValue<AliasedValue>("address.hil_street1").Value.ToString();
        //                    }
        //                    if (ent.Attributes.Contains("address.hil_street2"))
        //                    {
        //                        invoiceData.HOUSE_NO = ent.GetAttributeValue<AliasedValue>("address.hil_street2").Value.ToString();
        //                    }
        //                    if (ent.Attributes.Contains("address.hil_street3"))
        //                    {
        //                        invoiceData.STREET4 = ent.GetAttributeValue<AliasedValue>("address.hil_street3").Value.ToString();
        //                    }
        //                    if (ent.Attributes.Contains("address.hil_pincode"))
        //                    {
        //                        invoiceData.POSTAL_CODE = ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_pincode").Value).Name;
        //                    }
        //                    if (ent.Attributes.Contains("address.hil_district"))
        //                    {
        //                        invoiceData.DISTRICT = ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_district").Value).Name;
        //                    }
        //                    if (ent.Attributes.Contains("address.hil_city"))
        //                    {
        //                        invoiceData.CITY = ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_city").Value).Name;
        //                    }
        //                    if (ent.Attributes.Contains("address.hil_state"))
        //                    {
        //                        Entity entTemp = _service.Retrieve("hil_state", ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_state").Value).Id, new ColumnSet("hil_sapstatecode"));
        //                        if (entTemp.Attributes.Contains("hil_sapstatecode"))
        //                        {
        //                            invoiceData.REGION_CODE = entTemp.GetAttributeValue<string>("hil_sapstatecode").ToString();
        //                        }
        //                    }
        //                    invoiceData.CALL_ID = ent.GetAttributeValue<string>("name");
        //                    if (ent.Attributes.Contains("contact.emailaddress1"))
        //                    {
        //                        invoiceData.EMAIL = ent.GetAttributeValue<AliasedValue>("contact.emailaddress1").Value.ToString();
        //                    }
        //                    if (ent.Attributes.Contains("contact.mobilephone"))
        //                    {
        //                        invoiceData.PHONE = ent.GetAttributeValue<AliasedValue>("contact.mobilephone").Value.ToString();
        //                    }
        //                    if (ent.Attributes.Contains("address.hil_branch"))
        //                    {
        //                        invoiceData.BRANCH_NAME = ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_city").Value).Name;
        //                    }

        //                    //if (ent.Attributes.Contains("hil_branch"))
        //                    //{
        //                    //    invoiceData.BRANCH_NAME = ent.GetAttributeValue<EntityReference>("hil_branch").Name;
        //                    //}

        //                    invoiceData.MATERIAL_CODE = prodCode; //((EntityReference)(ent.GetAttributeValue<AliasedValue>("custAsset.msdyn_product").Value)).Name;

        //                    invoiceData.CS_SERVICE_PRICE = ent.GetAttributeValue<Money>("hil_receiptamount").Value.ToString();

        //                    //if (ent.Attributes.Contains("custAsset.msdyn_name"))
        //                    //{
        //                    //    invoiceData.PRODUCT_SERIAL_NO = ent.GetAttributeValue<AliasedValue>("custAsset.msdyn_name").Value.ToString();
        //                    //}
        //                    //if (ent.Attributes.Contains("custAsset.hil_invoicedate"))
        //                    //{
        //                    //    string dopValue = string.Empty;
        //                    //    dopValue = Convert.ToDateTime(ent.GetAttributeValue<AliasedValue>("custAsset.hil_invoicedate").Value).ToString("yyyy-MM-dd");
        //                    //    invoiceData.PRODUCT_DOP = dopValue;
        //                    //}
        //                    if (ent.Attributes.Contains("user.hil_employeecode"))
        //                    {
        //                        invoiceData.EMPLOYEE_ID = ent.GetAttributeValue<AliasedValue>("user.hil_employeecode").Value.ToString();
        //                    }
        //                    //if (ent.Attributes.Contains("custAsset.hil_warrantytilldate"))
        //                    //{
        //                    //    string dopValue = string.Empty;
        //                    //    dopValue = Convert.ToDateTime(ent.GetAttributeValue<AliasedValue>("custAsset.hil_warrantytilldate").Value).ToString("yyyy-MM-dd");
        //                    //    invoiceData.WARNTY_TILL_DATE = dopValue;
        //                    //}
        //                    if (ent.Attributes.Contains("msdyn_invoicedate"))
        //                    {
        //                        string jobclosedonValue = string.Empty;
        //                        jobclosedonValue = ent.GetAttributeValue<DateTime>("msdyn_invoicedate").ToString("yyyy-MM-dd");
        //                        invoiceData.CLOSED_ON = jobclosedonValue;
        //                    }
        //                    //if (ent.Attributes.Contains("custAsset.hil_warrantystatus"))
        //                    //{
        //                    //    OptionSetValue optValue = ((OptionSetValue)ent.GetAttributeValue<AliasedValue>("custAsset.hil_warrantystatus").Value);
        //                    //    invoiceData.WARNTY_STATUS = optValue.Value == 1 ? "IN" : "OUT";
        //                    //}
        //                    //else
        //                    //{
        //                        invoiceData.WARNTY_STATUS = "OUT";
        //                    //}
        //                    string pricingDateValue = string.Empty;
        //                    //pricingDateValue = ent.GetAttributeValue<DateTime>("createdon").ToString("yyyy-MM-dd");  // Changed on 06/Sep/2021 as discussed with Vinay & Ashish Tondon
        //                    pricingDateValue = ent.GetAttributeValue<DateTime>("msdyn_invoicedate").ToString("yyyy-MM-dd");
        //                    invoiceData.PRICING_DATE = pricingDateValue;

        //                    Decimal _receiptAmt = ent.GetAttributeValue<Money>("hil_receiptamount").Value;
        //                    Decimal _payableAmt = 0;

        //                    if (ent.Attributes.Contains("totallineitemamount"))
        //                    {
        //                        _payableAmt = ent.GetAttributeValue<Money>("totallineitemamount").Value;
        //                    }

        //                    if (_payableAmt < _receiptAmt) continue;

        //                    invoiceData.ZDS6_AMOUNT = (_payableAmt - _receiptAmt).ToString();
        //                    invoiceData.DISCOUNT_FLAG = "X";

        //                    requestData.IM_DATA.Add(invoiceData);

        //                    var Json = JsonConvert.SerializeObject(requestData);
        //                    //WebRequest request = WebRequest.Create("https://middlewaredev.havells.com:50001/RESTAdapter/dynamics/amc_data");
        //                    //WebRequest request = WebRequest.Create("https://middlewareqa.havells.com:50001/RESTAdapter/dynamics/amc_data");
        //                    //WebRequest request = WebRequest.Create("https://middleware.havells.com:50001/RESTAdapter/dynamics/amc_data");
        //                    WebRequest request = WebRequest.Create("https://p90ci.havells.com:50001/RESTAdapter/dynamics/amc_data");
        //                    string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
        //                    request.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
        //                    // Set the Method property of the request to POST.  
        //                    request.Method = "POST";
        //                    byte[] byteArray = Encoding.UTF8.GetBytes(Json);
        //                    // Set the ContentType property of the WebRequest.  
        //                    request.ContentType = "application/x-www-form-urlencoded";
        //                    // Set the ContentLength property of the WebRequest.  
        //                    request.ContentLength = byteArray.Length;
        //                    // Get the request stream.  
        //                    Stream dataStream = request.GetRequestStream();
        //                    // Write the data to the request stream.  
        //                    dataStream.Write(byteArray, 0, byteArray.Length);
        //                    // Close the Stream object.  
        //                    dataStream.Close();
        //                    // Get the response.  
        //                    WebResponse response = request.GetResponse();
        //                    // Display the status.  
        //                    Console.WriteLine(((HttpWebResponse)response).StatusDescription);
        //                    // Get the stream containing content returned by the server.  
        //                    dataStream = response.GetResponseStream();
        //                    // Open the stream using a StreamReader for easy access.  
        //                    StreamReader reader = new StreamReader(dataStream);
        //                    // Read the content.  
        //                    string responseFromServer = reader.ReadToEnd();
        //                    // responseFromServer = responseFromServer.Replace("ZBAPI_CREATE_SALES_ORDER.Response", "ZBAPI_CREATE_SALES_ORDER"); // removed for new response
        //                    AMCInvoiceData resp = JsonConvert.DeserializeObject<AMCInvoiceData>(responseFromServer);
        //                    if (responseFromServer.Replace(@"""", "").IndexOf("STATUS:S,") > 0)
        //                    {
        //                        //Entity entUpdate = new Entity(msdyn_workorder.EntityLogicalName);
        //                        //entUpdate.Id = ent.Id;
        //                        //entUpdate["hil_amcsyncstatus"] = new OptionSetValue(2);
        //                        //_service.Update(entUpdate);
        //                        //Console.WriteLine("Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
        //                    }
        //                    else
        //                    {
        //                        Entity entUpdate = new Entity(msdyn_workorder.EntityLogicalName);
        //                        entUpdate.Id = ent.Id;
        //                        entUpdate["hil_amcsyncstatus"] = new OptionSetValue(3);
        //                        _service.Update(entUpdate);
        //                        Console.WriteLine("Not Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
        //                    }
        //                }
        //                else
        //                {
        //                    Console.WriteLine("Product Code does not exist: Invoice " + jobId + " /" + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
        //                }
        //            }
        //        }
        //        else
        //        {
        //            Console.WriteLine("AMC_Invoice_Sync.Program.Main.PushAMCData :: No record found to sync");
        //        }
        //        Console.WriteLine("Sync Ended:");
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("AMC_Invoice_Sync.Program.Main.PushAMCData :: Error While Loading App Settings:" + ex.Message.ToString());
        //    }
        //}
        static void PushAMCData_WP()
        {
            IntegrationConfig intConfig = Program.IntegrationConfiguration(_service, "amc_data");
            //intConfig.uri = "https://middlewareqa.havells.com:50001/RESTAdapter/dynamics/amc_data";
            //intConfig.Auth = "D365_Havells:QAD365@1234";
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
                lnkEntProduct.LinkCriteria.AddCondition("producttypecode", ConditionOperator.Equal, 1);
                lnkEntProduct.LinkCriteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 5);


                QueryExpression Query = new QueryExpression("invoice");
                Query.ColumnSet = new ColumnSet("hil_customerasset", "name", "hil_salestype", "hil_productcode", "totallineitemamount", "hil_receiptamount",
                    "discountamount", "hil_productcode", "msdyn_invoicedate");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 2)); //Paid
                Query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 100001)); //Completed
                Query.Criteria.AddCondition(new ConditionExpression("hil_receiptamount", ConditionOperator.GreaterThan, 0)); //Receipt amount must be grater than 0.
                Query.Criteria.AddCondition(new ConditionExpression("hil_customerasset", ConditionOperator.Null)); //Invoice ID
                Query.Criteria.AddCondition(new ConditionExpression("name", ConditionOperator.Equal, "INV-2023-000613")); //Invoice ID
                Query.AddOrder("msdyn_invoicedate", OrderType.Ascending);
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
                                Entity entTemp = _service.Retrieve("hil_state", ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_state").Value).Id, new ColumnSet("hil_sapstatecode"));
                                if (entTemp.Attributes.Contains("hil_sapstatecode"))
                                {
                                    invoiceData.REGION_CODE = entTemp.GetAttributeValue<string>("hil_sapstatecode").ToString();
                                }
                            }

                            Query = new QueryExpression("msdyn_paymentdetail");
                            Query.ColumnSet = new ColumnSet("msdyn_name", "statuscode");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("msdyn_invoice", ConditionOperator.Equal, ent.Id);
                            //Query.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 910590000);
                            Query.AddOrder("createdon", OrderType.Descending);
                            EntityCollection FoundPaymentDetails = _service.RetrieveMultiple(Query);
                            if (FoundPaymentDetails.Entities.Count == 0)
                            {
                                Console.WriteLine("No Payment Found");
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

                            invoiceData.PRODUCT_SERIAL_NO = ent.GetAttributeValue<EntityReference>("hil_customerasset").Name;
                            invoiceData.PRODUCT_DOP = "";
                            DateTime _AMCStartDate = DateTime.Now;
                            DateTime _AMCEndDate = _AMCStartDate.AddDays(365);
                            invoiceData.START_DATE = _AMCStartDate.Year.ToString() + "-0" + _AMCStartDate.Month.ToString() + "-" + _AMCStartDate.Day.ToString();
                            invoiceData.END_DATE = _AMCEndDate.Year.ToString() + "-0" + _AMCEndDate.Month.ToString() + "-" + _AMCEndDate.Day.ToString(); ;

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
                                var res = (SetStateResponse)_service.Execute(req);
                                Console.WriteLine("Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                            }
                            else
                            {
                                Console.WriteLine("Not Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                            }
                        }
                        else
                        {
                            Console.WriteLine("Product Code does not exist: Invoice " + jobId + " /" + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
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

        static void PushAMCData_OmniChannel()
        {
            IntegrationConfig intConfig = Program.IntegrationConfiguration(_service, "amc_data");
            //intConfig.uri = "https://p90ci.havells.com:50001/RESTAdapter/dynamics/amc_data";
            //intConfig.Auth = "D365_Havells:PRDD365@1234";
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
                    "discountamount", "hil_productcode", "msdyn_invoicedate", "createdon", "hil_amcsellingsource","");

                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 2)); //Paid
                Query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 100001)); //Completed
                //Query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 910590000)); //Posted
                Query.Criteria.AddCondition(new ConditionExpression("hil_receiptamount", ConditionOperator.GreaterThan, 0)); //Receipt amount must be grater than 0.
                Query.Criteria.AddCondition(new ConditionExpression("hil_customerasset", ConditionOperator.NotNull)); //Customer Asset Contains Data
                //Query.Criteria.AddCondition(new ConditionExpression("name", ConditionOperator.Equal, "INV-2023-036516")); //Invoice ID
                Query.Criteria.AddCondition(new ConditionExpression("hil_salestype", ConditionOperator.Equal, 3));//AMC Omnichannel
                Query.Criteria.AddCondition(new ConditionExpression("createdon", ConditionOperator.OnOrAfter, "2024-02-01")); //Paid
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
                        try
                        {
                            recordCnt += 1;
                            prodCode = ent.GetAttributeValue<EntityReference>("hil_productcode").Name;// GetMaterialCode(ent.Id);
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
                                    Entity entTemp = _service.Retrieve("hil_state", ((EntityReference)ent.GetAttributeValue<AliasedValue>("address.hil_state").Value).Id, new ColumnSet("hil_sapstatecode"));
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
                                EntityCollection FoundPaymentDetails = _service.RetrieveMultiple(Query);
                                if (FoundPaymentDetails.Entities.Count == 0)
                                {
                                    Console.WriteLine("No Payment Found");
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
                                    string _campaignCode = ent.GetAttributeValue<string>("hil_amcsellingsource").ToString();
                                    string[] _campaignCodeArr = _campaignCode.Split('|');
                                    if (_campaignCodeArr.Length >= 1)
                                        invoiceData.EMPLOYEE_ID = _campaignCodeArr[0].ToUpper();
                                    if (_campaignCodeArr.Length > 1)
                                        invoiceData.EMP_TYPE = _campaignCodeArr[1].ToUpper();
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

                                string _AMCStartDate = GetWarrantyStartDate(_service, ent.GetAttributeValue<EntityReference>("hil_customerasset").Id, ent.GetAttributeValue<DateTime>("createdon"));// DateTime.Now;
                                if (_AMCStartDate != string.Empty)
                                {
                                    string _AMCEndDate = GetWarrantyEndDate(_service, ent.GetAttributeValue<EntityReference>("hil_productcode").Id, Convert.ToDateTime(_AMCStartDate));
                                    if (_AMCEndDate != null)
                                    {
                                        Console.WriteLine("Start Date " + _AMCStartDate + " || EndDate " + _AMCEndDate);
                                        invoiceData.START_DATE = _AMCStartDate;
                                        invoiceData.END_DATE = _AMCEndDate;

                                        requestData.IM_DATA.Add(invoiceData);

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
                                        AMCInvoiceData resp = JsonConvert.DeserializeObject<AMCInvoiceData>(responseFromServer);
                                        Console.WriteLine(responseFromServer);
                                        if (responseFromServer.Replace(@"""", "").IndexOf("STATUS:S,") > 0)
                                        {
                                            SetStateRequest req = new SetStateRequest();
                                            req.State = new OptionSetValue(2);
                                            req.Status = new OptionSetValue(910590000);
                                            req.EntityMoniker = ent.ToEntityReference();
                                            var res = (SetStateResponse)_service.Execute(req);
                                            Console.WriteLine("Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                                        }
                                        else
                                        {
                                            Console.WriteLine("Not Done " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("AMC End Date is NUll: Invoice# " + jobId + " / " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("AMC Start Date is NUll: Invoice# " + jobId + " / " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                                }
                            }
                            else
                            {
                                Console.WriteLine("Product Code does not exist: Invoice " + jobId + " / " + recordCnt.ToString() + "/" + totalRecords.ToString() + ":" + ent.GetAttributeValue<string>("msdyn_name"));
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error " + ex.Message);
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
        static void SendSMS()
        {

            try
            {
                SMSRequestModel smsReq = new SMSRequestModel() { MobileNumber = "8285906486", Message = "Test SMS" };
                SMSResponseModel smsRes = new SMSResponseModel();

                var Json = JsonConvert.SerializeObject(smsReq);
                WebRequest request = WebRequest.Create("http://chatapiqa.havells.com/sms/OTPCentralService.svc/SendSMS");
                // Set the Method property of the request to POST.  
                request.Method = "POST";
                byte[] byteArray = Encoding.UTF8.GetBytes(Json);
                // Set the ContentType property of the WebRequest.  
                request.ContentType = "application/json";
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
                SMSResponseModel resp = JsonConvert.DeserializeObject<SMSResponseModel>(responseFromServer);
                if (resp.ResponseMessage == "Success")
                {
                    Console.WriteLine("Done");
                }
                else
                {
                    Console.WriteLine("Not Done");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("AMC_Invoice_Sync.Program.Main.PushAMCData :: Error While Loading App Settings:" + ex.Message.ToString());
            }
            Console.ReadLine();
        }
        static string GetMaterialCode(Guid jobGuId)
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
        #region AMC Start And ENd Date Calculation
        static string GetWarrantyStartDate(IOrganizationService service, Guid AssetID, DateTime _purchaseDate)
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
                    Query = new QueryExpression(hil_unitwarranty.EntityLogicalName);
                    Query.ColumnSet = new ColumnSet("hil_warrantystartdate", "hil_warrantytemplate");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition(new ConditionExpression("hil_customerasset", ConditionOperator.Equal, AssetID));
                    EntityCollection ec1 = service.RetrieveMultiple(Query);
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
        static string GetWarrantyEndDate(IOrganizationService service, Guid _AMCPlaneID, DateTime StartDate)
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


        #region App Setting Load/CRM Connection
        static IOrganizationService ConnectToCRM(string connectionString)
        {
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;

        }
        #endregion
    }
    public class IntegrationConfig
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
    public class SparePart
    {
        public string partFamily { get; set; }
        public int count { get; set; }
    }
}
