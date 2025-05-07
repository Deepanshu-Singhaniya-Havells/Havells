using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RestSharp;
using Newtonsoft.Json;
using System.Text;

namespace HavellsNewPlugin.SAP_IntegrationForOrderCreation
{
    public class Integration : IPlugin
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
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace(entity.LogicalName);
                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_tenderno"));
                    EntityReference tenderID = entity.GetAttributeValue<EntityReference>("hil_tenderno");
                    QueryExpression qsChecklist = new QueryExpression("hil_orderchecklist");
                    qsChecklist.ColumnSet = new ColumnSet("hil_tenderno", "hil_nameofclientcustomercode", "hil_paymentterms", "hil_approvedtransporter", "hil_projectname", "hil_poloinofooter", "hil_poloino", "hil_podate", "hil_prices", "hil_paymentterms", "hil_clearedformanufacturing", "hil_approveddatasheetgtp", "hil_approvedqap", "hil_inspection", "hil_ldterms", "hil_approvedtransporter", "hil_bankguaranteerequired", "hil_valueofbg", "hil_warrantyperiod", "hil_clearedon", "hil_drumlengthschedule", "hil_inspectionagency", "hil_specialinstructions", "hil_typeofdrum");
                    qsChecklist.NoLock = true;
                    qsChecklist.Criteria = new FilterExpression(LogicalOperator.And);
                    qsChecklist.Criteria.AddCondition("hil_tenderno", ConditionOperator.Equal, tenderID.Id);
                    qsChecklist.Criteria.AddCondition("hil_synwithsap", ConditionOperator.Equal, false);
                    qsChecklist.Criteria.AddCondition("hil_approvalstatus", ConditionOperator.Equal, 1);
                    EntityCollection entCol = service.RetrieveMultiple(qsChecklist);
                    if (entCol.Entities.Count > 0)
                    {
                        GetOANumber(service, entCol);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error " + ex.Message);
            }
        }
        static dynamic orderChecklistId;
        public static void GetOANumber(IOrganizationService service, EntityCollection OdrChkLstColl)
        {
            try
            {

                RootCheckList RCL = new RootCheckList();
                List<ITTABLEHEADER> iTTABLEHEADERLst = new List<ITTABLEHEADER>();
                List<ITTABLEITEM> iTTABLEITEMLst = new List<ITTABLEITEM>();

                foreach (Entity lstPO in OdrChkLstColl.Entities)
                {
                    ITTABLEHEADER iTTABLEHEADER = new ITTABLEHEADER();
                    iTTABLEHEADER.ZTENDERNO = Convert.ToString(lstPO.GetAttributeValue<EntityReference>("hil_tenderno").Name);
                    iTTABLEHEADER.ZNAMEOFCLIENT = lstPO.Contains("hil_nameofclientcustomercode") ?
                        Convert.ToString(service.Retrieve("account", lstPO.GetAttributeValue<EntityReference>("hil_nameofclientcustomercode").Id, new ColumnSet("accountnumber")).GetAttributeValue<string>("accountnumber")) : "";//lstPO.GetAttributeValue<EntityReference>("hil_nameofclientcustomercode").Name) : "";
                    iTTABLEHEADER.ZPROJNAME = Convert.ToString(lstPO.GetAttributeValue<string>("hil_projectname"));
                    iTTABLEHEADER.ZPO_FOOTER = Convert.ToString(lstPO.GetAttributeValue<string>("hil_poloinofooter"));
                    iTTABLEHEADER.ZPO_LOINO = Convert.ToString(lstPO.GetAttributeValue<string>("hil_poloino"));
                    iTTABLEHEADER.ZPO_DATE = Convert.ToString(dateFormatter(lstPO.GetAttributeValue<DateTime>("hil_podate")));
                    iTTABLEHEADER.ZPRICES = (lstPO.Contains("hil_prices") ? lstPO.FormattedValues["hil_prices"] : "").ToString();
                    iTTABLEHEADER.ZPAYTERM = "B004";//lstPO.Contains("hil_paymentterms") ? Convert.ToString(lstPO.FormattedValues["hil_paymentterms"]) : "";
                    if (lstPO.Contains("hil_clearedformanufacturing"))
                    {
                        iTTABLEHEADER.ZMANUFCT_CLEAR = lstPO.GetAttributeValue<bool>("hil_clearedformanufacturing") == true ? "YES" : "NO";
                    }
                    else
                    {
                        iTTABLEHEADER.ZMANUFCT_CLEAR = "NO";
                    }
                    iTTABLEHEADER.ZUSAGE = "P";//usase indicator master
                    if (lstPO.Contains("hil_approveddatasheetgtp"))
                    {
                        iTTABLEHEADER.ZDATA_GTP = lstPO.GetAttributeValue<bool>("hil_approveddatasheetgtp") == true ? "YES" : "NO";
                    }
                    else
                    {
                        iTTABLEHEADER.ZDATA_GTP = "NO";
                    }
                    if (lstPO.Contains("hil_approvedqap"))
                    {
                        iTTABLEHEADER.ZQAP = lstPO.GetAttributeValue<bool>("hil_approvedqap") == true ? "YES" : "NO";
                    }
                    else
                    {
                        iTTABLEHEADER.ZQAP = "NO";
                    }
                    if (lstPO.Contains("hil_inspection"))
                    {
                        iTTABLEHEADER.ZINSPECTION = lstPO.GetAttributeValue<bool>("hil_inspection") == true ? "YES" : "NO";
                    }
                    iTTABLEHEADER.ZZLEADCOD = "NA";

                    var LDTerm = "L.D. Terms : " + (lstPO.Contains("hil_ldterms") ? Convert.ToString(lstPO.GetAttributeValue<string>("hil_ldterms")) : "").ToString();
                    var ApprovedTransporter = " ApprovedTransporter :" + (lstPO.Contains("hil_approvedtransporter") ? Convert.ToString(lstPO.FormattedValues["hil_approvedtransporter"]) : "").ToString();
                    var BankGuaranteeRequired = "BankGuaranteeRequired : " + (lstPO.Contains("hil_bankguaranteerequired") ? lstPO.GetAttributeValue<bool>("hil_bankguaranteerequired") : false).ToString();
                    var CategoryOfBankGurantee = "CategoryOfBankGurantee : " + (lstPO.Contains("hil_categoryofbankguarantee") ? lstPO.FormattedValues["hil_categoryofbankguarantee"] : "").ToString();
                    var Valueofbg = "Value Of Bg % :" + (lstPO.Contains("hil_valueofbg") ? lstPO.FormattedValues["hil_valueofbg"] : "").ToString();
                    var warrantyperiod = "Warranty Period :" + (lstPO.Contains("hil_warrantyperiod") ? lstPO.GetAttributeValue<EntityReference>("hil_warrantyperiod").Name : "").ToString();
                    var Dispatchclearedon = "Dispatchclearedon : " + (lstPO.Contains("hil_clearedon") ? dateFormatter(lstPO.GetAttributeValue<DateTime>("hil_clearedon")).ToString() : "").ToString();
                    var drumlengthschedule = "Drum length schedule : " + (lstPO.Contains("hil_drumlengthschedule") ? lstPO.GetAttributeValue<bool>("hil_drumlengthschedule") : false).ToString();
                    var inspectionagency = "Inspectionagency : " + (lstPO.Contains("hil_inspectionagency") ? lstPO.GetAttributeValue<string>("hil_inspectionagency") : "").ToString();
                    var specialInstruction = "Specialinstructions :" + (lstPO.Contains("hil_specialinstructions") ? lstPO.GetAttributeValue<string>("hil_specialinstructions") : "").ToString();
                    var hil_typeofdrum = "Type Of Drum :" + (lstPO.Contains("hil_typeofdrum") ? lstPO.FormattedValues["hil_typeofdrum"] : "").ToString();
                    iTTABLEHEADER.ZHEADER_TEXT = LDTerm + " , " + ApprovedTransporter + " , " + BankGuaranteeRequired + " , " + CategoryOfBankGurantee + Valueofbg + " , " + warrantyperiod + " , " + Dispatchclearedon + " , " + hil_typeofdrum + " , " + drumlengthschedule + " , " + inspectionagency + " , " + specialInstruction;
                    iTTABLEHEADERLst.Add(iTTABLEHEADER);

                    orderChecklistId = lstPO.Id;
                    QueryExpression query = new QueryExpression("hil_tenderproduct");
                    query.ColumnSet = new ColumnSet("hil_tenderid", "hil_product", "hil_poqty", "hil_plantcode", "hil_basicpriceinrsmtr",
                        "hil_freightcharges", "hil_hopricespecialconstructions", "hil_toleranceupperlimit", "hil_tolerancelowerlimit", "hil_name", "hil_inspectiontype");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition(new ConditionExpression("hil_orderchecklistid", ConditionOperator.Equal, lstPO.Id));
                    LinkEntity EntityA = new LinkEntity("hil_tenderproduct", "hil_deliveryschedule", "hil_tenderproductid", "hil_tenderproduct", JoinOperator.Inner);
                    EntityA.Columns = new ColumnSet("hil_deliverydate", "hil_name", "hil_deliveryscheduleid");
                    EntityA.EntityAlias = "PEnq";
                    query.LinkEntities.Add(EntityA);

                    EntityCollection entColl = service.RetrieveMultiple(query);
                    foreach (Entity product in entColl.Entities)
                    {
                        ITTABLEITEM iTTABLEITEM = new ITTABLEITEM();
                        iTTABLEITEM.ZTENDERNO = product.Contains("hil_tenderid") ? Convert.ToString(product.GetAttributeValue<EntityReference>("hil_tenderid").Name) : "";
                        iTTABLEITEM.ZMATNR = product.Contains("hil_product") ? product.GetAttributeValue<EntityReference>("hil_product").Name : "";
                        iTTABLEITEM.ZQUANTITY = product.Contains("hil_poqty") ? Convert.ToString(product.GetAttributeValue<decimal>("hil_poqty")) : "";
                        iTTABLEITEM.ZUOM = "M";
                        iTTABLEITEM.ZWERKS = product.Contains("hil_plantcode") ? product.GetAttributeValue<EntityReference>("hil_plantcode").Name : "";
                        iTTABLEITEM.ZNET_PRICE = product.Contains("hil_basicpriceinrsmtr") ? Convert.ToString(product.GetAttributeValue<Money>("hil_basicpriceinrsmtr").Value) : "";
                        iTTABLEITEM.ZZFRE = product.Contains("hil_freightcharges") ? Convert.ToString(product.GetAttributeValue<Money>("hil_freightcharges").Value) : "";
                        iTTABLEITEM.ZLD_PENALTY = product.Contains("hil_inspectiontype") ? Convert.ToString(service.Retrieve("hil_inspectiontype", product.GetAttributeValue<EntityReference>("hil_inspectiontype").Id, new ColumnSet("hil_code")).GetAttributeValue<string>("hil_code")) : "";
                        iTTABLEITEM.ZINDIVIDUAL = Convert.ToString(product.GetAttributeValue<Decimal>("hil_toleranceupperlimit"));
                        iTTABLEITEM.ZOVERALL = Convert.ToString(product.GetAttributeValue<Decimal>("hil_tolerancelowerlimit"));
                        iTTABLEITEM.ZAPP_PRICE = product.Contains("hil_hopricespecialconstructions") ? Convert.ToString(product.GetAttributeValue<Money>("hil_hopricespecialconstructions").Value) : "";
                        iTTABLEITEM.ZDELIVERY_DATE = dateFormatter((DateTime)product.GetAttributeValue<AliasedValue>("PEnq.hil_deliverydate").Value).ToString();
                        iTTABLEITEM.ZTENDER_LINE = Convert.ToString(product.GetAttributeValue<AliasedValue>("PEnq.hil_name").Value);
                        iTTABLEITEMLst.Add(iTTABLEITEM);
                    }
                }
                RCL.IT_TABLE_HEADER = iTTABLEHEADERLst;
                RCL.IT_TABLE_ITEM = iTTABLEITEMLst;
                IntegrationConfig integration = Models.IntegrationConfiguration(service, "GetLookupMasterDataMasterForHomeAdvisory");

                string _authInfo = integration.Auth;
                _authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authInfo));
                _authInfo = "Basic " + _authInfo;
                String sUrl = integration.uri;

                string data = JsonConvert.SerializeObject(RCL);
                var client = new RestClient(sUrl);
                //client.Timeout = -1;
                var request = new RestRequest();
                request.AddHeader("Authorization", _authInfo);
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("application/json", data, ParameterType.RequestBody);
                RestResponse response = client.Execute(request, Method.Post);
                dynamic returndatas = JsonConvert.DeserializeObject<RootCheckListReturn>(response.Content);
                Guid headerid = Guid.Empty;
                if (Convert.ToString(returndatas.RETURN).ToUpper() == "Sucess".ToUpper())
                {
                    if (returndatas.ET_TABLE.Count > 0)
                    {
                        List<oaheader> lstoaheader = new List<oaheader>();
                        foreach (var item in returndatas.ET_TABLE)
                        {
                            Entity OAHeader = new Entity("hil_oaheader");
                            var finditem = lstoaheader.Find(lstitem => lstitem.hil_oaheaderid == item.ORDERNUMBER);
                            if (finditem == null)
                            {
                                oaheader _Oaheader = new oaheader();
                                _Oaheader.hil_oaheaderid = item.ORDERNUMBER;
                                OAHeader["hil_name"] = Convert.ToString(item.ORDERNUMBER);
                                QueryExpression QueryTenderId = new QueryExpression("hil_tender");
                                QueryTenderId.ColumnSet = new ColumnSet("hil_tenderid");
                                QueryTenderId.NoLock = true;
                                QueryTenderId.Criteria = new FilterExpression(LogicalOperator.And);
                                QueryTenderId.Criteria.AddCondition("hil_name", ConditionOperator.Equal, item.ZTENDERNO);
                                EntityCollection entCol3 = service.RetrieveMultiple(QueryTenderId);
                                if (entCol3.Entities.Count > 0)
                                {
                                    OAHeader["hil_tenderid"] = entCol3[0].ToEntityReference();
                                }
                                OAHeader["hil_shiptopartyname"] = item.SHIP_NAME;
                                OAHeader["hil_shiptopartycode"] = item.SHIP_CODE;
                                QueryExpression QuerySalesOffice = new QueryExpression("hil_salesoffice");
                                QuerySalesOffice.ColumnSet = new ColumnSet("hil_name");
                                QuerySalesOffice.NoLock = true;
                                QuerySalesOffice.Criteria = new FilterExpression(LogicalOperator.And);
                                QuerySalesOffice.Criteria.AddCondition("hil_sapcode", ConditionOperator.Equal, item.VKBUR);
                                EntityCollection entCol2 = service.RetrieveMultiple(QuerySalesOffice);
                                if (entCol2.Entities.Count > 0)
                                {
                                    OAHeader["hil_salesoffice"] = entCol2[0].ToEntityReference();
                                }
                                OAHeader["hil_responsemessage"] = item.REMARKS;
                                OAHeader["hil_orderchecklistid"] = new EntityReference("hil_orderchecklist", orderChecklistId);
                                OAHeader["hil_oacreatedon"] = Convert.ToDateTime(item.ORDERCREATION);
                                OAHeader["hil_datasheetforgtp"] = item.ZDATA_GTP;
                                QueryExpression QueryCustName = new QueryExpression("account");
                                QueryCustName.ColumnSet = new ColumnSet("name");
                                QueryCustName.NoLock = true;
                                QueryCustName.Criteria = new FilterExpression(LogicalOperator.And);
                                QueryCustName.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, item.CUST_CODE);
                                EntityCollection entCol4 = service.RetrieveMultiple(QueryCustName);
                                if (entCol4.Entities.Count > 0)
                                {
                                    OAHeader["hil_customername"] = entCol4[0].ToEntityReference();
                                }
                                OAHeader["hil_customercode"] = item.CUST_CODE;
                                OAHeader["hil_creditlimit"] = Convert.ToDecimal(item.CREDITLIMIT);
                                OAHeader["hil_creditdayscustomercode"] = item.CREDITDAY_CUST;
                                OAHeader["hil_creditdays"] = Convert.ToInt32(item.CREDITDAYS);
                                lstoaheader.Add(_Oaheader);
                                headerid = service.Create(OAHeader);
                            }
                            Entity OAProduct = new Entity("hil_oaproduct");

                            OAProduct["hil_deliverydate"] = Convert.ToDateTime(item.DELIVERY_DATE);
                            QueryExpression Querydeliverscheduleid = new QueryExpression("hil_deliveryschedule");
                            Querydeliverscheduleid.ColumnSet = new ColumnSet("hil_deliveryscheduleid");
                            Querydeliverscheduleid.NoLock = true;
                            Querydeliverscheduleid.Criteria = new FilterExpression(LogicalOperator.And);
                            Querydeliverscheduleid.Criteria.AddCondition("hil_name", ConditionOperator.Equal, item.ZTENDER_LINE);
                            EntityCollection entCol1 = service.RetrieveMultiple(Querydeliverscheduleid);
                            if (entCol1.Entities.Count > 0)
                            {
                                OAProduct["hil_deliveryschedule"] = entCol1[0].ToEntityReference();
                            }

                            OAProduct["hil_materialcategorydescription"] = Convert.ToString(item.CATEGORY_DISC);
                            OAProduct["hil_materialgroup"] = item.MATKL;
                            OAProduct["hil_netvalue"] = new Money(Convert.ToDecimal(item.VALUE));
                            QueryExpression QueryProductCode = new QueryExpression("product");
                            QueryProductCode.ColumnSet = new ColumnSet("name");
                            QueryProductCode.NoLock = true;
                            QueryProductCode.Criteria = new FilterExpression(LogicalOperator.And);
                            QueryProductCode.Criteria.AddCondition("productnumber", ConditionOperator.Equal, item.ITEM_CODE);
                            EntityCollection entCol5 = service.RetrieveMultiple(QueryProductCode);
                            if (entCol5.Entities.Count > 0)
                            {
                                OAProduct["hil_productcode"] = entCol5[0].ToEntityReference();
                            }
                            OAProduct["hil_productdescription"] = item.ITEM_DESC;
                            OAProduct["hil_productionplanningremarks"] = "";
                            OAProduct["hil_quantity"] = Convert.ToDecimal(item.QTY) * 1000;
                            OAProduct["hil_stockinhand"] = Convert.ToDecimal(item.STOCK);
                            OAProduct["hil_oaheader"] = new EntityReference(OAHeader.LogicalName, headerid);
                            service.Create(OAProduct);

                            Entity orderchecklist = new Entity("hil_orderchecklist");
                            orderchecklist.Id = orderChecklistId;
                            orderchecklist["hil_synwithsap"] = true;
                            service.Update(orderchecklist);
                        }
                    }
                }
                else
                {
                    Entity orderchecklist = new Entity("hil_orderchecklist");
                    orderchecklist.Id = orderChecklistId;
                    orderchecklist.Attributes["hil_syncwithsapremark"] = Convert.ToString(returndatas.RETURN);
                    service.Update(orderchecklist);
                }

            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error: " + ex.Message);
            }
        }
        public static string dateFormatter(DateTime date)
        {
            string _enquiryDatetime = string.Empty;
            if (date.Year.ToString().PadLeft(4, '0') != "0001")
                _enquiryDatetime = date.Year.ToString() + "-" + date.Month.ToString().PadLeft(2, '0') + "-" + date.Day.ToString().PadLeft(2, '0');
            return _enquiryDatetime;
        }


    }
}
