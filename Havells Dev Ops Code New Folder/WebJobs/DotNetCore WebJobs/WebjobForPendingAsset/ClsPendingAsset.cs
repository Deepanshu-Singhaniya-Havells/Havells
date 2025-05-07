using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Text;

namespace WebjobForPendingAsset
{
    internal class ClsPendingAsset
    {
        private readonly ServiceClient _service;
        public ClsPendingAsset(ServiceClient service)
        {
            _service = service;
        }
        public void InsertPendingAssetinERtable()
        {
            try
            {
                if (_service.IsReady)
                {
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
                           <attribute name='hil_productsubcategory' />
                           <attribute name='hil_source' />
                           <attribute name='hil_customer' />  
                           <attribute name='hil_modelname' /> 
                           <order attribute='createdon' descending='true' />
                           <filter type='and'>
                             <filter type='or'>
                                 <filter type='or'>
                                     <condition attribute='hil_branchheadapprovalstatus' operator='eq' value='1' />
                                     <condition attribute='statuscode' operator='eq' value='910590001' />                           
                                 </filter>
                             //<condition attribute='hil_source' operator='in'>
                             //    <value>6</value>
                             //    <value>5</value>
                             //    <value>7</value>
                             //    <value>12</value>
                             //</condition>
                             </filter>
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
                            <attribute name='hil_sapcode' />
                            <attribute name='hil_materialgroup' />
                             <link-entity name='hil_productcatalog' from='hil_productcode' to='productid' link-type='inner' alias='al'>
                               <filter type='and'>
                                 <condition attribute='hil_eligibleforloyaltyprograms' operator='eq' value='1' />
                               </filter>
                             </link-entity>
                           </link-entity>
                         </entity>
                       </fetch>";
                    Console.WriteLine("Getting Customer Asset Data Start");
                    EntityCollection entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entcoll.Entities.Count > 0)
                    {
                        Console.WriteLine("Getting Customer Asset Data count: {0}", entcoll.Entities.Count);
                        foreach (var c in entcoll.Entities)
                        {
                            decimal InvoiceValue = 0;
                            Console.WriteLine("Customer Asset Data details Start");
                            string MobileNumber = c.Contains("co.mobilephone") ? c.GetAttributeValue<AliasedValue>("co.mobilephone").Value.ToString() : null;
                            string Materialgroup = c.Contains("hil_productsubcategory") ? c.GetAttributeValue<EntityReference>("hil_productsubcategory").Name : null;
                            EntityReference productnumber = c.Contains("msdyn_product") ? (c.GetAttributeValue<EntityReference>("msdyn_product")) : null;
                            string Model = c.Contains("hil_modelname") ? c.Attributes["hil_modelname"].ToString() : null;
                            string Division = c.Contains("ak.hil_sapcode") ? c.GetAttributeValue<AliasedValue>("ak.hil_sapcode").Value.ToString() : null;
                            string CustomerAsset = c.Attributes["msdyn_name"].ToString();
                            string Source = c.Contains("hil_source") ? c.GetAttributeValue<OptionSetValue>("hil_source").Value.ToString() : "";
                            string invoiceDate = c.GetAttributeValue<DateTime>("hil_invoicedate").ToString("dd MMM yyyy");
                            Guid AssestID = new Guid(c.Attributes["msdyn_customerassetid"].ToString());
                            string CustomerName = c.Contains("hil_customer") ? c.GetAttributeValue<EntityReference>("hil_customer").Name : null;
                            string Category = c.Contains("hil_productcategory") ? c.GetAttributeValue<EntityReference>("hil_productcategory").Name : null;
                            Console.WriteLine("SAP Price API Initiate");
                            var ParamModelData = new MC_PRICEModel();
                            ParamModelData.LT_TABLE = new LT_TABLE();
                            ParamModelData.LT_TABLE.MATNR = productnumber.Name;
                            Console.WriteLine("Getting SAP price on SKU {0}", productnumber.Name);
                            var Paramdata = new StringContent(JsonConvert.SerializeObject(ParamModelData), System.Text.Encoding.UTF8, "application/json");

                            QueryExpression qe = new QueryExpression("hil_integrationconfiguration");
                            qe.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                            qe.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "Ecom_priceinfo");
                            Entity enColl = _service.RetrieveMultiple(qe)[0];
                            String URL = enColl.GetAttributeValue<string>("hil_url");
                            String Auth = enColl.GetAttributeValue<string>("hil_username") + ":" + enColl.GetAttributeValue<string>("hil_password");
                            Console.WriteLine("SAP API Call");
                            HttpClient Reqclient = new HttpClient();
                            var byteArray = Encoding.ASCII.GetBytes(Auth);
                            Reqclient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                            HttpResponseMessage Result = Reqclient.PostAsync(URL, Paramdata).Result;
                            if (Result.IsSuccessStatusCode)
                            {
                                Console.WriteLine("SAP API Status Success");
                                Response obj = JsonConvert.DeserializeObject<Response>(Result.Content.ReadAsStringAsync().Result);
                                if (obj != null)
                                {
                                    foreach (var item in obj.LT_TABLE)
                                    {
                                        if (item.KSCHL != null && item.KSCHL == "ZWEB")
                                        {
                                            if (!string.IsNullOrEmpty(item.KBETR))
                                                InvoiceValue = Math.Round(((Convert.ToDecimal(item.KBETR) * 70) / 100), 2);
                                            Console.WriteLine("SAP API Price: {0}", InvoiceValue);
                                        }
                                    }
                                }
                            }

                            if (InvoiceValue > 0)
                            {
                                Console.WriteLine("Data Insertion Start in Easy Reward Loyalty Program");
                                Entity ERCreate = new Entity("hil_easyrewardloyaltyprogram");
                                ERCreate["hil_division"] = Division;
                                ERCreate["hil_invoicevalue"] = InvoiceValue;
                                ERCreate["hil_customerasset"] = CustomerAsset;
                                ERCreate["hil_mobilenumber"] = MobileNumber;
                                ERCreate["hil_materialgroup"] = Materialgroup;
                                ERCreate["hil_invoicedate"] = Convert.ToDateTime(invoiceDate);
                                ERCreate["hil_source"] = Source;
                                ERCreate["hil_productsyned"] = false;
                                ERCreate["hil_syncstatus"] = new OptionSetValue(1);
                                ERCreate["hil_name"] = CustomerName;
                                ERCreate["hil_category"] = Category;
                                ERCreate["hil_synccount"] = 0;
                                var res = _service.Create(ERCreate);
                                Console.WriteLine("Data Insertion Success in Easy Reward Loyalty Program");
                                if (res != Guid.Empty)
                                {
                                    Entity entinvoice = new Entity("msdyn_customerasset", AssestID);
                                    entinvoice["hil_pushforloyaltyprograms"] = true;
                                    _service.Update(entinvoice);
                                    Console.WriteLine("Customer Asset Pushforloyaltyprograms Status Update");
                                }
                            }
                            Console.WriteLine("Completed.");
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
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
    public class LT_TABLE
    {
        public string MATNR { get; set; }
    }
    internal class Response
    {
        public List<LTTABLE> LT_TABLE { get; set; }
    }
    internal class LTTABLE
    {
        public string MATNR { get; set; }
        public string KSCHL { get; set; }
        public string DATAB { get; set; }
        public string KBETR { get; set; }
        public string KONWA { get; set; }
        public string DATBI { get; set; }
        public string DELETE_FLAG { get; set; }
        public string CREATEDBY { get; set; }
        public string CTIMESTAMP { get; set; }
        public string MODIFYBY { get; set; }
        public string MTIMESTAMP { get; set; }
    }
    public class MC_PRICEModel
    {
        public LT_TABLE LT_TABLE { get; set; }
    }
}
