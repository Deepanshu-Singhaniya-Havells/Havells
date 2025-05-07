using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;

namespace UpdateCustomerLoyaltyData
{
    internal class ClsUpdateCustomerLoyaltyData
    {
        private readonly ServiceClient _service;
        public ClsUpdateCustomerLoyaltyData(ServiceClient service)
        {
            _service = service;
        }
        public void UpdateCustomerLoyaltyData()
        {
            try
            {
                if (_service.IsReady)
                {
                    ResponseLoyalty response = new ResponseLoyalty();
                    string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                        <entity name='msdyn_customerasset'>
                          <attribute name='createdon' />
                          <attribute name='msdyn_product' />
                          <attribute name='msdyn_name' />
                          <attribute name='hil_invoicevalue' />   
                           <attribute name='hil_invoicedate' />  
                          <attribute name='hil_productsubcategorymapping' />
                          <attribute name='hil_productcategory' />
                          <attribute name='msdyn_customerassetid' />  
                           <attribute name='hil_customer' /> 
                          <attribute name='hil_modelname' /> 
                          <order attribute='createdon' descending='true' />
                          <filter type='and'>
                            <condition attribute='hil_invoicedate' operator='last-x-months' value='6' /> 
                           <condition attribute='hil_pushforloyaltyprograms' operator='ne' value='1'/>     
                           </filter>
                          <link-entity name='contact' from='contactid' to='hil_customer' visible='false' link-type='inner' alias='co'>
                          <attribute name='mobilephone'/>
                            <filter type='and'>
                                <condition attribute='hil_loyaltyprogramenabled' operator='eq' value='1' />
                              </filter>
                          </link-entity>
                          <link-entity name='product' from='productid' to='hil_productcategory' link-type='inner' alias='ak'>
                          <attribute name='hil_division' />
                          <attribute name='hil_materialgroup' />
                            <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='al'>
                              <filter type='and'>
                                <condition attribute='hil_eligibleforloyaltyprograms' operator='eq' value='1' />
                              </filter>
                            </link-entity>
                          </link-entity>
                        </entity>
                      </fetch>";
                    Console.WriteLine("Getting Customer Assest for loyalty Programs");
                    EntityCollection entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entcoll.Entities.Count > 0)
                    {
                        Console.WriteLine("Getting Customer Assest count: {0} for loyalty Programs", entcoll.Entities.Count);
                        foreach (var c in entcoll.Entities)
                        {
                            string MobileNumber = c.Contains("co.mobilephone") ? c.GetAttributeValue<AliasedValue>("co.mobilephone").Value.ToString() : null;
                            string Model = c.Contains("hil_modelname") ? c.Attributes["hil_modelname"].ToString() : null;
                            decimal InvoiceValue = c.Contains("hil_invoicevalue") ? decimal.Round(c.GetAttributeValue<decimal>("hil_invoicevalue"), 2) : 0;
                            string Division = c.Contains("ak.hil_division") ? ((EntityReference)c.GetAttributeValue<AliasedValue>("ak.hil_division").Value).Name : null;
                            string Materialgroup = c.Contains("ak.hil_materialgroup") ? ((EntityReference)c.GetAttributeValue<AliasedValue>("ak.hil_materialgroup").Value).Name : null;
                            string CustomerAsset = c.Attributes["msdyn_name"].ToString();
                            if (MobileNumber == null)
                            {
                                Console.WriteLine("Customer Assest: {0} Mobile Number not found.", CustomerAsset);
                                continue;
                            }
                            if (InvoiceValue == 0)
                            {
                                Console.WriteLine("Invalid Invoice value.");
                                continue;
                            }
                            else if (Model == null)
                            {
                                Console.WriteLine("Model not found.");
                                continue;
                            }
                            else if (Division == null || Materialgroup == null)
                            {
                                Console.WriteLine("Product Configuration(Division/Materialgroup) not found.");
                                continue;
                            }
                            Console.WriteLine("Processing Customer Assest: {0} for Customer: {1}", CustomerAsset, MobileNumber);

                            SKUBillDetails BillDetails = new SKUBillDetails();
                            BillDetails.TransactionItems = new TransactionItems();
                            BillDetails.TransactionItems.TransactionItem = new List<TransactionItem>();
                            BillDetails.PaymentMode = new PaymentMode();
                            BillDetails.PaymentMode.TenderItem = new List<TenderItem>();

                            TransactionItem objTransactionItem = new TransactionItem();
                            objTransactionItem.ItemType = "s";
                            objTransactionItem.ItemQty = "1";
                            objTransactionItem.ItemName = Model;
                            objTransactionItem.Size = null;
                            objTransactionItem.Unit = "1";
                            objTransactionItem.ItemDiscount = "0";
                            objTransactionItem.ItemTax = "0";
                            objTransactionItem.TotalPrice = InvoiceValue;
                            objTransactionItem.BilledPrice = InvoiceValue;
                            objTransactionItem.Department = Division;
                            objTransactionItem.Category = Division;
                            objTransactionItem.Group = Materialgroup;
                            objTransactionItem.ItemId = CustomerAsset;
                            objTransactionItem.RefBillNo = "";
                            BillDetails.TransactionItems.TransactionItem.Add(objTransactionItem);

                            TenderItem tenderItem = new TenderItem();
                            tenderItem.TenderCode = "Cash";
                            tenderItem.TenderID = "";
                            tenderItem.TenderValue = InvoiceValue;

                            BillDetails.PaymentMode.TenderItem.Add(tenderItem);
                            BillDetails.CountryCode = "91";
                            BillDetails.TransactionDate = c.GetAttributeValue<DateTime>("hil_invoicedate").ToString("dd MMM yyyy");
                            BillDetails.BillNo = CustomerAsset;
                            BillDetails.Channel = "Online";
                            BillDetails.CustomerType = "Loyalty";
                            BillDetails.BillValue = InvoiceValue;
                            BillDetails.PointsRedeemed = "";
                            BillDetails.PointsValueRedeemed = "";
                            BillDetails.PrimaryOrderNumber = CustomerAsset;
                            BillDetails.ShippingCharges = "";
                            BillDetails.PreDelivery = "False";
                            BillDetails.SKUOfferCode = "";

                            /////////////////////// ER API Call for Data Post /////////////////////////////////// 

                            var data = new StringContent(JsonConvert.SerializeObject(BillDetails), System.Text.Encoding.UTF8, "application/json");
                            data.Headers.Add("LoginUserId", MobileNumber.ToString());
                            data.Headers.Add("OperationToken", "");
                            HttpClient client = new HttpClient();
                            Console.WriteLine("Push customer Asset Data to ER API for Earn loyalty points");
                            HttpResponseMessage response1 = client.PostAsync("https://middlewaredev.havells.com:50001/RESTAdapter/ER/SaveSKUBillDetails?IM_PROJECT=D365", data).Result;
                            if (response1.IsSuccessStatusCode)
                            {
                                var result = response1.Content.ReadAsStringAsync().Result;
                                response = JsonConvert.DeserializeObject<ResponseLoyalty>(result);
                                if (response.ReturnMessage.Contains("Success."))
                                {
                                    Entity entinvoice = new Entity("msdyn_customerasset", entcoll.Entities[0].Id);
                                    entinvoice["hil_pushforloyaltyprograms"] = true;
                                    _service.Update(entinvoice);
                                    Console.WriteLine(response.ReturnMessage);
                                }
                                else
                                {
                                    Console.WriteLine(response.ReturnMessage);
                                }
                            }
                            Console.WriteLine("Completed.");

                            /////////////////////// ER Tier Update API Call ///////////////////////
                            TierModel tierModel = new TierModel();
                            var data1 = new StringContent(JsonConvert.SerializeObject(tierModel), System.Text.Encoding.UTF8, "application/json");
                            data1.Headers.Add("LoginUserId", MobileNumber.ToString());
                            data1.Headers.Add("OperationToken", "");
                            HttpClient client1 = new HttpClient();
                            Console.WriteLine("Get customer Tier data from ER API");
                            HttpResponseMessage responsetier = client1.PostAsync("https://middlewaredev.havells.com:50001/RESTAdapter/ER/GetTierDetails?IM_PROJECT=D365", data1).Result;
                            if (responsetier.IsSuccessStatusCode)
                            {
                                try
                                {
                                    var resulttier = responsetier.Content.ReadAsStringAsync().Result;
                                    dynamic TierResponse = JsonConvert.DeserializeObject<dynamic>(resulttier);
                                    string? Tier = TierResponse != null ? TierResponse["TierName"] : null;
                                    if (!String.IsNullOrWhiteSpace(Tier))
                                    {
                                        int TierValue = 0;
                                        switch (Tier.ToUpper())
                                        {
                                            case "PLUS":
                                            case "BASE":
                                                TierValue = 1;
                                                break;
                                            case "PREMIUM":
                                                TierValue = 2;
                                                break;
                                            case "PRIVILEDGE":
                                                TierValue = 3;
                                                break;
                                            default:
                                                TierValue = 0;
                                                break;
                                        }
                                        Guid CustomerGuid = c.GetAttributeValue<EntityReference>("hil_customer").Id;
                                        Entity UpdateTier = new Entity("contact", CustomerGuid);
                                        UpdateTier["hil_loyaltyprogramtier"] = new OptionSetValue(TierValue);
                                        _service.Update(UpdateTier);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex.Message);
                                    continue;
                                }
                            }
                            Console.WriteLine("Completed..");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Error: D365 Service is unavailable!");
                }
            }
            catch (Exception ex)
            {
                Console.Write("Error: " + ex.Message);
            }
        }
    }
    public class TierModel
    {
        public string CountryCode { get; set; } = "91";
    }
    public class ResponseLoyalty
    {
        public string ReturnCode { get; set; }
        public string ReturnMessage { get; set; }
        public string OrderLevelCredits { get; set; }
        public string TransactionId { get; set; }
    }

    public class SKUBillDetails
    {
        public string CountryCode { get; set; }
        public string TransactionDate { get; set; }
        public string BillNo { get; set; }
        public string Channel { get; set; }
        public string CustomerType { get; set; }
        public decimal BillValue { get; set; }
        public string PointsRedeemed { get; set; }
        public string PointsValueRedeemed { get; set; }
        public string PrimaryOrderNumber { get; set; }
        public string ShippingCharges { get; set; }
        public string PreDelivery { get; set; }
        public string SKUOfferCode { get; set; }
        public TransactionItems TransactionItems { get; set; }
        public PaymentMode PaymentMode { get; set; }

    }

    public class TransactionItem
    {
        public string ItemType { get; set; }
        public string ItemQty { get; set; }
        public string ItemName { get; set; }
        public string Size { get; set; }
        public string Unit { get; set; }
        public string ItemDiscount { get; set; }
        public string ItemTax { get; set; }
        public decimal TotalPrice { get; set; }
        public decimal BilledPrice { get; set; }
        public string Department { get; set; }
        public string Category { get; set; }
        public string Group { get; set; }
        public string ItemId { get; set; }
        public string RefBillNo { get; set; }
    }

    public class TransactionItems
    {
        public List<TransactionItem> TransactionItem { get; set; }
    }

    public class TenderItem
    {
        public string TenderCode { get; set; }
        public string TenderID { get; set; }
        public decimal TenderValue { get; set; }
    }

    public class PaymentMode
    {
        public List<TenderItem> TenderItem { get; set; }
    }
}
