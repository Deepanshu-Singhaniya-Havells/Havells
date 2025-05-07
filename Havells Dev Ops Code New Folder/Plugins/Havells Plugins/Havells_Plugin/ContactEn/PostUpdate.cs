using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace Havells_Plugin.ContactEn
{
    public class PostUpdate : IPlugin
    {
        private static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == Contact.EntityLogicalName && context.MessageName.ToUpper() == "UPDATE" && context.Depth == 1)
                {
                    tracingService.Trace("1");
                    Entity entity = (Entity)context.InputParameters["Target"];
                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));
                    //ContactEn.Common.SetFullAddress(entity, context.MessageName.ToUpper(), service);
                    if (entity.Contains("hil_loyaltyprogramenabled") || entity.Contains("hil_loyaltyprogramtier"))
                    {
                        Entity _custLoyaltyInfo = service.Retrieve(entity.LogicalName, entity.Id,new ColumnSet("hil_loyaltyprogramenabled", "hil_loyaltyprogramtier"));

                        bool Isloyaltyprogramenabled = _custLoyaltyInfo.Contains("hil_loyaltyprogramenabled") ? _custLoyaltyInfo.GetAttributeValue<bool>("hil_loyaltyprogramenabled") : false;
                        if (Isloyaltyprogramenabled)
                        {
                            OptionSetValue loyaltyTier = _custLoyaltyInfo.Contains("hil_loyaltyprogramtier") ? _custLoyaltyInfo.GetAttributeValue<OptionSetValue>("hil_loyaltyprogramtier") : null;
                            if (loyaltyTier != null)
                            {
                                Entity _entUpdate = new Entity(entity.LogicalName, entity.Id);
                                _entUpdate["entityimage"] = loyaltyTier.Value == 3 ? Convert.FromBase64String(LoyaltyMemnberTierImage.GoldImage) : loyaltyTier.Value == 2 ? Convert.FromBase64String(LoyaltyMemnberTierImage.SilverImage) : Convert.FromBase64String(LoyaltyMemnberTierImage.BrongeImage);
                                service.Update(_entUpdate);
                            }
                            //Fetch All Applicable Products and push to Easy Rewards Platform
                            PushLoyaltyEligibleData(service, entity);
                        }
                        else 
                        {
                            Entity _entUpdate = new Entity(entity.LogicalName, entity.Id);
                            _entUpdate["hil_loyaltyprogramtier"] = null;
                            _entUpdate["entityimage"] = null;
                            service.Update(_entUpdate);
                        }
                    }
                    Common.CallAPIInsertUpdateCustomer(service, entity, tracingService);
                }
            }
            catch (Exception ex)
            {
                Entity _address = new Entity(((Entity)context.InputParameters["Target"]).LogicalName);
                _address["hil_errormsg"] = ex.Message;
                _address.Id = ((Entity)context.InputParameters["Target"]).Id;
                service.Update(_address);
            }
        }

        private void PushLoyaltyEligibleData(IOrganizationService service, Entity entity) {
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
                        </filter>
                        <link-entity name='contact' from='contactid' to='hil_customer' visible='false' link-type='inner' alias='co'>
                            <attribute name='mobilephone'/>                            
                            <filter type='and'>
                                <condition attribute='hil_loyaltyprogramenabled' operator='eq' value='1' />     
                                <condition attribute='contactid' operator='eq' value='{entity.Id}' />   
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
            EntityCollection entcoll = service.RetrieveMultiple(new FetchExpression(_fetchXML));

            if (entcoll.Entities.Count > 0)
            {
                foreach (var c in entcoll.Entities)
                {
                    decimal InvoiceValue = 0;
                    try
                    {
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
                        tracingService.Trace("Initiate SAP Api for Price information !!");
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
                            tracingService.Trace("Step-4");
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
                            tracingService.Trace("Record Created Sucessfully on EasyRewardLoyality Table in D365");

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
                        throw new InvalidPluginExecutionException("Havells_Plugin.Consumer.ConsumerPostUpdate.Execute Error " + ex.Message);

                    }
                }
            }
        }
    }
    #region  START Model of SAP AND EASY REWARD API 
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
    #endregion
}
