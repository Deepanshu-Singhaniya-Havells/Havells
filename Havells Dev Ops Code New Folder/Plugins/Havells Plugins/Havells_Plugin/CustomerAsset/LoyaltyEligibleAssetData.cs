using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace Havells_Plugin.CustomerAsset
{
    public class LoyaltyEligibleAssetData : IPlugin
    {

        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {

            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                tracingService.Trace("First Step");
                if (context.InputParameters.Contains("Target")
                && context.InputParameters["Target"] is Entity && context.PrimaryEntityName.ToLower() == "msdyn_customerasset" && context.Depth == 1)
                {
                    tracingService.Trace("Execution Start");
                    Entity entity = (Entity)context.InputParameters["Target"];
                    Entity _entCA = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_customer"));
                    if (_entCA.Contains("hil_customer"))
                    {
                        Entity _customerEligible = service.Retrieve(_entCA.GetAttributeValue<EntityReference>("hil_customer").LogicalName, _entCA.GetAttributeValue<EntityReference>("hil_customer").Id, new ColumnSet("hil_loyaltyprogramenabled"));
                        bool _isEnabledForLoyaltyProg = false;
                        if (_customerEligible.Contains("hil_loyaltyprogramenabled"))
                        {
                            _isEnabledForLoyaltyProg = _customerEligible.GetAttributeValue<bool>("hil_loyaltyprogramenabled");
                        }

                        if (_isEnabledForLoyaltyProg)
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
                                 <condition attribute='hil_source' operator='in'>
                                     <value>6</value>
                                     <value>5</value>
                                     <value>7</value>
                                     <value>12</value>
                                 </condition>
                                 </filter>
                                <condition attribute='hil_invoicedate' operator='last-x-months' value='6' /> 
                                <condition attribute='hil_pushforloyaltyprograms' operator='ne' value='1'/>      
                                 <condition attribute='msdyn_customerassetid' operator='eq' value='{entity.Id}' />   
                               </filter>
                              <link-entity name='contact' from='contactid' to='hil_customer' visible='false' link-type='inner' alias='co'>
                                 <attribute name='mobilephone'/> 
                                  <attribute name='hil_loyaltyprogramenabled'/>                                    
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
                            tracingService.Trace("Getting Customer Information for Tier Update loyalty Programs");

                            EntityCollection entcoll = service.RetrieveMultiple(new FetchExpression(_fetchXML));

                            if (entcoll.Entities.Count > 0)
                            {
                                tracingService.Trace("step-2");
                                bool hil_loyaltyprogramenabled = (bool)(entcoll.Entities[0].Contains("co.hil_loyaltyprogramenabled") ? entcoll.Entities[0].GetAttributeValue<AliasedValue>("co.hil_loyaltyprogramenabled").Value : false);
                                tracingService.Trace(hil_loyaltyprogramenabled.ToString());
                                if (hil_loyaltyprogramenabled)
                                {
                                    //check if customer is Loyalty enabled
                                    tracingService.Trace(entcoll.Entities.Count.ToString());
                                    foreach (var c in entcoll.Entities)
                                    {
                                        decimal InvoiceValue = 0;
                                        try
                                        {
                                            //decimal InvoiceValue = c.Contains("hil_invoicevalue") ? decimal.Round(c.GetAttributeValue<decimal>("hil_invoicevalue"), 2) : 0;
                                            string MobileNumber = c.Contains("co.mobilephone") ? c.GetAttributeValue<AliasedValue>("co.mobilephone").Value.ToString() : null;
                                            string Materialgroup = c.Contains("hil_productsubcategory") ? c.GetAttributeValue<EntityReference>("hil_productsubcategory").Name : null;
                                            EntityReference productnumber = c.Contains("msdyn_product") ? (c.GetAttributeValue<EntityReference>("msdyn_product")) : null;
                                            string Model = c.Contains("hil_modelname") ? c.Attributes["hil_modelname"].ToString() : null;
                                            string Division = c.Contains("ak.hil_sapcode") ? c.GetAttributeValue<AliasedValue>("ak.hil_sapcode").Value.ToString() : null;
                                            string Category = c.Contains("hil_productcategory") ? c.GetAttributeValue<EntityReference>("hil_productcategory").Name : null;
                                            string CustomerAsset = c.Attributes["msdyn_name"].ToString();
                                            string Source = c.Contains("hil_source") ? c.GetAttributeValue<OptionSetValue>("hil_source").Value.ToString() : "";
                                            string invoiceDate = c.GetAttributeValue<DateTime>("hil_invoicedate").ToString("dd MMM yyyy");
                                            Guid AssestID = new Guid(c.Attributes["msdyn_customerassetid"].ToString());
                                            tracingService.Trace(AssestID.ToString());
                                            string CustomerName = c.Contains("hil_customer") ? c.GetAttributeValue<EntityReference>("hil_customer").Name : null;

                                            var ParamModelData = new MC_PRICEModel();
                                            ParamModelData.LT_TABLE = new LT_TABLE();
                                            ParamModelData.LT_TABLE.MATNR = productnumber.Name;

                                            var Paramdata = new StringContent(JsonConvert.SerializeObject(ParamModelData), System.Text.Encoding.UTF8, "application/json");

                                            QueryExpression qe = new QueryExpression("hil_integrationconfiguration");
                                            qe.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                                            qe.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "Ecom_priceinfo");
                                            Entity enColl = service.RetrieveMultiple(qe)[0];
                                            String URL = enColl.GetAttributeValue<string>("hil_url");
                                            String Auth = enColl.GetAttributeValue<string>("hil_username") + ":" + enColl.GetAttributeValue<string>("hil_password");
                                            
                                            HttpClient Reqclient = new HttpClient();
                                            var byteArray = Encoding.ASCII.GetBytes(Auth);
                                            Reqclient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                                            tracingService.Trace("Push customer Asset Data to ER API for Earn loyalty points");
                                            HttpResponseMessage Result = Reqclient.PostAsync(URL, Paramdata).Result;
                                            if (Result.IsSuccessStatusCode)
                                            {
                                                Response obj = JsonConvert.DeserializeObject<Response>(Result.Content.ReadAsStringAsync().Result);
                                                if (obj != null)
                                                {

                                                    foreach (var item in obj.LT_TABLE)
                                                    {
                                                        if (item.KSCHL != null && item.KSCHL == "ZWEB")
                                                        {
                                                            if (!string.IsNullOrEmpty(item.KBETR))
                                                                InvoiceValue = Math.Round(((Convert.ToDecimal(item.KBETR) * 70) / 100), 2);
                                                        }
                                                    }
                                                }
                                            }

                                            if (InvoiceValue > 0)
                                            {
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
                                                ERCreate["hil_synccount"] = 0;
                                                ERCreate["hil_category"] = Category;
                                                var res = service.Create(ERCreate);

                                                if (res != null)
                                                {
                                                    Entity entinvoice = new Entity("msdyn_customerasset", AssestID);
                                                    entinvoice["hil_pushforloyaltyprograms"] = true;
                                                    service.Update(entinvoice);
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.CustomerAssetPreCreate.Execute Error " + ex.Message);
                                        }
                                    }
                                }
                                else
                                {
                                    string MobileNumber = entcoll.Entities[0].Contains("co.mobilephone") ? entcoll.Entities[0].GetAttributeValue<AliasedValue>("co.mobilephone").Value.ToString() : null;
                                    tracingService.Trace(MobileNumber);
                                    var ParamModelData = new TrackUsers();
                                    ParamModelData.phoneNumber = MobileNumber;
                                    ParamModelData.countryCode = "+91";
                                    ParamModelData.traits = new trait();
                                    ParamModelData.traits.name = entcoll.Entities[0].Contains("hil_customer") ? entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_customer").Name : null;

                                    var RESULT = UpdateWhatsAppUser(ParamModelData, service);
                                    tracingService.Trace(RESULT.ToString());
                                    var ParamData = new TrackEvents();

                                    ParamData.traits = new trait();
                                    ParamData.phoneNumber = MobileNumber;

                                    ParamData.countryCode = "+91";
                                    ParamData.@event = "loyalty_enroll";
                                    tracingService.Trace("SendSMS", MobileNumber);
                                    var result = SendMessage(ParamData, service);
                                    tracingService.Trace("Send sms result", result.ToString());
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.CustomerAssetPreCreate.Execute Error " + ex.Message);

            }
        }

        public sendMesgRes SendMessage(TrackEvents campaignDetails, IOrganizationService service)
        {
            IntegrationConfig intConfig = IntegrationConfiguration(service, "WhatsApp Campaign Details");
            string url = intConfig.uri;
            string authInfo = intConfig.Auth;
            string Campaigndata = JsonConvert.SerializeObject(campaignDetails);
            using (var client = new HttpClient())
            {
                HttpContent objCampaign = new StringContent(Campaigndata);
                client.BaseAddress = new Uri(url);
                client.DefaultRequestHeaders.Add("Authorization", authInfo);
                var result = client.PostAsync(url, objCampaign).Result;
                return JsonConvert.DeserializeObject<sendMesgRes>(result.Content.ReadAsStringAsync().Result);
            }
        }
        public UpdateAppUserRes UpdateWhatsAppUser(TrackUsers userDetails, IOrganizationService service)
        {
            IntegrationConfig intConfig = IntegrationConfiguration(service, "Update WhatsApp User");
            string url = intConfig.uri;
            string authInfo = intConfig.Auth;
            string data = JsonConvert.SerializeObject(userDetails);
            using (var client = new HttpClient())
            {
                HttpContent objcontent = new StringContent(data);
                client.BaseAddress = new Uri(url);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Add("Authorization", authInfo);
                var result = client.PostAsync(url, objcontent).Result;
                return JsonConvert.DeserializeObject<UpdateAppUserRes>(result.Content.ReadAsStringAsync().Result);
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
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + " " + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
            return output;
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
            public string SourceType { get; set; }
            public TransactionItems TransactionItems { get; set; }
            public PaymentMode PaymentMode { get; set; }

        }

        public class TransactionItem
        {
            public string ItemType { get; set; }
            public string ItemQty { get; set; }
            public string ItemName { get; set; }
            public string Size { get; set; }
            public decimal Unit { get; set; }
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


        public class TrackUsers
        {
            public string phoneNumber { get; set; }
            public string countryCode { get; set; }
            public string tags { get; set; }
            public trait traits { get; set; }
        }
        public class TrackEvents
        {
            public string phoneNumber { get; set; }
            public string countryCode { get; set; }
            public string @event { get; set; }
            public trait traits { get; set; }
        }
        public class trait
        {
            public string name { get; set; }
        }
        public class UpdateAppUserRes
        {
            public string result { get; set; }
            public string message { get; set; }
        }
        public class sendMesgRes
        {
            public string result { get; set; }
            public string message { get; set; }
            public string id { get; set; }
        }
        public class IntegrationConfig
        {
            public string uri { get; set; }
            public string Auth { get; set; }
        }
    }
}
