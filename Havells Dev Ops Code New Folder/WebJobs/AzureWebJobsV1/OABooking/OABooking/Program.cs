using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel.Description;
using System.Text;

namespace OABooking
{
    public class Program
    {
        static IOrganizationService service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Program Started");
                LoadAppSettings();
                // Class1.getTenderProduct(service);
                // changeAutoNumber("93b8869c-5831-ec11-b6e6-002248d4c580");
                //GetReadinessDates(service);

                QueryExpression qsChecklist = new QueryExpression("hil_orderchecklist");
                qsChecklist.ColumnSet = new ColumnSet("ownerid", "hil_name", "hil_tenderno", "hil_typeoforder", "hil_nameofclientcustomercode", "hil_paymentterms", "hil_approvedtransporter", "hil_projectname",
                    "hil_poloinofooter", "hil_poloino", "hil_podate", "hil_prices", "hil_paymentterms", "hil_clearedformanufacturing", "hil_approveddatasheetgtp",
                    "hil_approvedqap", "hil_inspection", "hil_ldterms", "hil_approvedtransporter", "hil_bankguaranteerequired", "hil_valueofbg", "hil_warrantyperiod",
                    "hil_clearedon", "hil_drumlengthschedule", "hil_inspectionagency", "hil_specialinstructions", "hil_typeofdrum", "hil_consigneeaddress");
                qsChecklist.Criteria = new FilterExpression(LogicalOperator.And);
                if (args.Length > 0)
                {
                    qsChecklist.Criteria.AddCondition("hil_orderchecklistid", ConditionOperator.Equal, new Guid(args[0]));
                }
                else
                {
                    qsChecklist.Criteria.AddCondition("hil_syncwithsapremark", ConditionOperator.Null);
                }
                qsChecklist.Criteria.AddCondition("hil_clearedformanufacturing", ConditionOperator.Equal, true);
                qsChecklist.Criteria.AddCondition("hil_synwithsap", ConditionOperator.Equal, false);
                qsChecklist.Criteria.AddCondition("hil_approvalstatus", ConditionOperator.Equal, 1);
                qsChecklist.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                EntityCollection entCol = service.RetrieveMultiple(qsChecklist);
                Console.WriteLine("Count of OCL " + entCol.Entities.Count);
                if (entCol.Entities.Count > 0)
                {
                    GetOANumberPlantWise(entCol);
                }
                Console.WriteLine("Program Terminated Secussfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex.Message);
                Console.WriteLine("Program Terminated with error");
            }

        }
        public static void GetOANumberPlantWise(EntityCollection entCol)
        {
            Console.WriteLine("Total Record " + entCol.Entities.Count);
            foreach (Entity lstPO in entCol.Entities)
            {
                string CustomerName = string.Empty;
                string ProjectName = string.Empty;
                string OANumber = string.Empty;
                string EnqNo = string.Empty;
                string orderchecklistno = string.Empty;
                string ORDERTYPE = lstPO.FormattedValues["hil_typeoforder"];

                EntityReference tenderOwner = new EntityReference();
                RootCheckList RCL = new RootCheckList();
                List<ITTABLEHEADER> iTTABLEHEADERLst = new List<ITTABLEHEADER>();
                EntityReference owner = lstPO.GetAttributeValue<EntityReference>("ownerid");

                string[] ordertypearr = ORDERTYPE.Split('-');
                ORDERTYPE = ordertypearr[1];
                Console.WriteLine("Retriving User Config");
                QueryExpression query1 = new QueryExpression("hil_userbranchmapping");
                query1.ColumnSet = new ColumnSet("ownerid");
                query1.NoLock = true;
                query1.Criteria = new FilterExpression(LogicalOperator.And);
                query1.Criteria.AddCondition("hil_user", ConditionOperator.Equal, owner.Id);
                LinkEntity EntityP = new LinkEntity("hil_buhead", "systemuser", "hil_buhead", "systemuserid", JoinOperator.LeftOuter);
                EntityP.Columns = new ColumnSet("domainname");
                EntityP.EntityAlias = "Pbuhead";
                query1.LinkEntities.Add(EntityP);
                LinkEntity EntityQ = new LinkEntity("hil_zonalhead", "systemuser", "hil_zonalhead", "systemuserid", JoinOperator.LeftOuter);
                EntityQ.Columns = new ColumnSet("domainname");
                EntityQ.EntityAlias = "Pzonalhead";
                query1.LinkEntities.Add(EntityQ);
                LinkEntity EntityR = new LinkEntity("hil_branchproducthead", "systemuser", "hil_branchproducthead", "systemuserid", JoinOperator.LeftOuter);
                EntityR.Columns = new ColumnSet("domainname");
                EntityR.EntityAlias = "PBPhead";
                query1.LinkEntities.Add(EntityR);
                LinkEntity EntityS = new LinkEntity("hil_user", "systemuser", "hil_user", "systemuserid", JoinOperator.LeftOuter);
                EntityS.Columns = new ColumnSet("domainname");
                EntityS.EntityAlias = "Puser";
                query1.LinkEntities.Add(EntityS);
                EntityCollection entColUser = service.RetrieveMultiple(query1);
                ITTABLEHEADER iTTABLEHEADER = new ITTABLEHEADER();
                Console.WriteLine("User Config Retrived count is " + entColUser.Entities.Count);
                if (entColUser.Entities.Count > 0)
                {
                    iTTABLEHEADER.ZSALES_REP = Convert.ToString(entColUser.Entities[0].GetAttributeValue<AliasedValue>("Puser.domainname").Value); ;
                    iTTABLEHEADER.ZBRANCH_PH = Convert.ToString(entColUser.Entities[0].GetAttributeValue<AliasedValue>("PBPhead.domainname").Value);
                    iTTABLEHEADER.ZZONAL_HEAD = Convert.ToString(entColUser.Entities[0].GetAttributeValue<AliasedValue>("Pzonalhead.domainname").Value);
                    iTTABLEHEADER.ZBUSINES_UH = Convert.ToString(entColUser.Entities[0].GetAttributeValue<AliasedValue>("Pbuhead.domainname").Value);
                }
                if (lstPO.Contains("hil_tenderno"))
                {
                    iTTABLEHEADER.ZTENDERNO = Convert.ToString(lstPO.GetAttributeValue<EntityReference>("hil_tenderno").Name);
                }
                else
                {
                    iTTABLEHEADER.ZTENDERNO = Convert.ToString(lstPO.GetAttributeValue<string>("hil_name")); ;
                }
                iTTABLEHEADER.ZORDERTYPE = ORDERTYPE;

                iTTABLEHEADER.ZNAMEOFCLIENT = lstPO.Contains("hil_nameofclientcustomercode") ?
                    Convert.ToString(service.Retrieve("account", lstPO.GetAttributeValue<EntityReference>("hil_nameofclientcustomercode").Id, new ColumnSet("accountnumber")).GetAttributeValue<string>("accountnumber"))
                    : "";
                iTTABLEHEADER.ZPROJNAME = Convert.ToString(lstPO.GetAttributeValue<string>("hil_projectname"));
                Entity tender = null; //declare a empty Entity
                Console.WriteLine("retriving Owner");
                if (lstPO.Contains("hil_tenderno"))
                {
                    EnqNo = Convert.ToString(lstPO.GetAttributeValue<EntityReference>("hil_tenderno").Name);
                    tender = service.Retrieve(
                            lstPO.GetAttributeValue<EntityReference>("hil_tenderno").LogicalName,
                            lstPO.GetAttributeValue<EntityReference>("hil_tenderno").Id,
                            new ColumnSet("ownerid")
                        );
                    tenderOwner = tender.GetAttributeValue<EntityReference>("ownerid");
                }
                else
                {
                    tenderOwner = lstPO.GetAttributeValue<EntityReference>("ownerid");
                }
                Console.WriteLine("retrived Owner");
                ProjectName = Convert.ToString(lstPO.GetAttributeValue<string>("hil_projectname"));
                CustomerName = iTTABLEHEADER.ZNAMEOFCLIENT;
                orderchecklistno = Convert.ToString(lstPO.GetAttributeValue<string>("hil_name"));
                iTTABLEHEADER.ZPO_FOOTER = Convert.ToString(lstPO.GetAttributeValue<string>("hil_poloinofooter"));
                iTTABLEHEADER.ZPO_LOINO = Convert.ToString(lstPO.GetAttributeValue<string>("hil_poloino"));
                iTTABLEHEADER.ZPO_DATE = Convert.ToString(dateFormatter(lstPO.GetAttributeValue<DateTime>("hil_podate").AddMinutes(330)));
                iTTABLEHEADER.ZPRICES = (lstPO.Contains("hil_prices") ? lstPO.FormattedValues["hil_prices"] : "").ToString();
                if (lstPO.Contains("hil_clearedformanufacturing"))
                {
                    iTTABLEHEADER.ZMANUFCT_CLEAR = lstPO.GetAttributeValue<bool>("hil_clearedformanufacturing") == true ? "YES" : "NO";
                }
                else
                {
                    iTTABLEHEADER.ZMANUFCT_CLEAR = "NO";
                }
                iTTABLEHEADER.ZUSAGE = ORDERTYPE == "DS" ? "S" : "P"; //usase indicator master
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

                if (lstPO.Contains("hil_consigneeaddress"))
                    iTTABLEHEADER.ZSHIP_PARTY = lstPO.GetAttributeValue<string>("hil_consigneeaddress");
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
                RCL.IT_TABLE_HEADER = iTTABLEHEADERLst;
                orderChecklistId = lstPO.Id;

                Console.WriteLine("header Details Competed..");


                QueryExpression queryDistinctPlancode = new QueryExpression("hil_orderchecklistproduct");
                queryDistinctPlancode.ColumnSet = new ColumnSet("hil_plantcode");
                queryDistinctPlancode.Distinct = true;
                queryDistinctPlancode.Criteria = new FilterExpression(LogicalOperator.And);
                queryDistinctPlancode.Criteria.AddCondition(new ConditionExpression("hil_orderchecklistid", ConditionOperator.Equal, lstPO.Id));
                queryDistinctPlancode.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                queryDistinctPlancode.Criteria.AddCondition(new ConditionExpression("hil_poqty", ConditionOperator.GreaterThan, 0));
                EntityCollection entColdelv = service.RetrieveMultiple(queryDistinctPlancode);
                String[] OaSAPMessage = new String[entColdelv.Entities.Count];
                int i = 0;
                Console.WriteLine("total Active Product Retrived count " + entColdelv.Entities.Count);
                foreach (Entity distPlant in entColdelv.Entities)
                {

                    try
                    {
                        List<ITTABLEITEM> iTTABLEITEMLst = new List<ITTABLEITEM>();
                        QueryExpression query = new QueryExpression("hil_orderchecklistproduct");
                        query.ColumnSet = new ColumnSet(true);
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition(new ConditionExpression("hil_orderchecklistid", ConditionOperator.Equal, lstPO.Id));
                        query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                        query.Criteria.AddCondition(new ConditionExpression("hil_plantcode", ConditionOperator.Equal, distPlant.GetAttributeValue<EntityReference>("hil_plantcode").Id));
                        queryDistinctPlancode.Criteria.AddCondition(new ConditionExpression("hil_poqty", ConditionOperator.GreaterThan, 0));
                        EntityCollection entColl = service.RetrieveMultiple(query);
                        foreach (Entity product in entColl.Entities)
                        {
                            ITTABLEITEM iTTABLEITEM = new ITTABLEITEM();

                            iTTABLEITEM.ZTENDERNO = lstPO.Contains("hil_tenderno") ? lstPO.GetAttributeValue<EntityReference>("hil_tenderno").Name :
                                product.GetAttributeValue<EntityReference>("hil_orderchecklistid").Name;


                            iTTABLEITEM.ZMATNR = product.Contains("hil_product") ? product.GetAttributeValue<EntityReference>("hil_product").Name : "";

                            var quantity = product.Contains("hil_poqty") ? product.GetAttributeValue<decimal>("hil_poqty").ToString() : product.GetAttributeValue<decimal>("hil_quantity").ToString();

                            iTTABLEITEM.ZQUANTITY = String.Format("{0:0.00}", quantity);
                            iTTABLEITEM.ZUOM = "M";

                            iTTABLEITEM.ZWERKS = product.Contains("hil_plantcode") ? service.Retrieve(product.GetAttributeValue<EntityReference>("hil_plantcode").LogicalName,
                                product.GetAttributeValue<EntityReference>("hil_plantcode").Id, new ColumnSet("hil_plantname")).GetAttributeValue<string>("hil_plantname") : "";

                            iTTABLEITEM.ZNET_PRICE = product.Contains("hil_porate") ? String.Format("{0:0.00}", product.GetAttributeValue<Money>("hil_porate").Value) :
                                String.Format("{0:0.00}", product.GetAttributeValue<Money>("hil_basicpriceinrsmtr").Value);

                            iTTABLEITEM.ZZFRE = product.Contains("hil_freightcharges") ? String.Format("{0:0.00}", product.GetAttributeValue<Money>("hil_freightcharges").Value) : "";
                            iTTABLEITEM.ZLD_PENALTY = product.Contains("hil_inspectiontype") ? Convert.ToString(service.Retrieve("hil_inspectiontype", product.GetAttributeValue<EntityReference>("hil_inspectiontype").Id, new ColumnSet("hil_code")).GetAttributeValue<string>("hil_code")) : "";
                            iTTABLEITEM.ZINDIVIDUAL = String.Format("{0:0.00}", product.GetAttributeValue<Decimal>("hil_toleranceupperlimit"));
                            iTTABLEITEM.ZOVERALL = String.Format("{0:0.00}", product.GetAttributeValue<Decimal>("hil_tolerancelowerlimit"));
                            iTTABLEITEM.ZAPP_PRICE = product.Contains("hil_hopricespecialconstructions") ? String.Format("{0:0.00}", product.GetAttributeValue<Money>("hil_hopricespecialconstructions").Value) : String.Format("{0:0.00}", product.GetAttributeValue<Money>("hil_basicpriceinrsmtr").Value);

                            string deliverDate = product.Contains("hil_deliverydate") ?
                                 dateFormatter((DateTime)product.GetAttributeValue<DateTime>("hil_deliverydate").AddMinutes(330)).ToString() :
                                 genrateDeliveryDate(service, product.ToEntityReference());
                            DateTime Delv = DateTime.Parse(deliverDate);
                            if (Delv.Date >= DateTime.Now.Date)
                            {
                                iTTABLEITEM.ZDELIVERY_DATE = deliverDate;
                            }
                            else
                            {
                                deliverDate = genrateDeliveryDate(service, product.ToEntityReference());
                                //throw new Exception("Delivery Date is not Valid");
                            }
                            iTTABLEITEM.ZTENDER_LINE = product.GetAttributeValue<string>("hil_name");
                            iTTABLEITEMLst.Add(iTTABLEITEM);
                            Console.WriteLine("Line ID:- " + iTTABLEITEM.ZTENDER_LINE);
                        }

                        #region Integration Hit
                        RCL.IT_TABLE_ITEM = iTTABLEITEMLst;
                        Console.WriteLine("Data Fetched");
                        IntegrationConfig intConfig = Models.IntegrationConfiguration(service, "GetCableOrder");
                        string url = intConfig.uri;
                        string authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));
                        string data = JsonConvert.SerializeObject(RCL);
                        var client = new RestClient(url);
                        client.Timeout = -1;
                        var request = new RestRequest(Method.POST);

                        Entity intigrationTrace = new Entity("hil_integrationtrace");
                        intigrationTrace["hil_entityname"] = lstPO.LogicalName;
                        intigrationTrace["hil_entityid"] = lstPO.Id.ToString();
                        intigrationTrace["hil_request"] = data;
                        intigrationTrace["hil_name"] = lstPO.GetAttributeValue<string>("hil_name");

                        Guid intigrationTraceID = service.Create(intigrationTrace);

                        request.AddHeader("Authorization", "Basic " + authInfo);
                        request.AddHeader("Content-Type", "application/json");

                        request.AddParameter("application/json", data, ParameterType.RequestBody);
                        IRestResponse response = client.Execute(request);

                        Entity intigrationTraceUp = new Entity("hil_integrationtrace");
                        intigrationTraceUp["hil_response"] = response.Content == "" ? response.ErrorMessage : response.Content;
                        intigrationTraceUp.Id = intigrationTraceID;
                        service.Update(intigrationTraceUp);
                        Console.WriteLine("Data Fetched");
                        dynamic returndatas = JsonConvert.DeserializeObject<RootCheckListReturn>(response.Content);
                        Guid headerid = Guid.Empty;
                        OaSAPMessage[i] = returndatas.RETURN;
                        i++;
                        Console.WriteLine("Response Status = " + Convert.ToString(returndatas.RETURN).ToUpper());
                        if (Convert.ToString(returndatas.RETURN).ToUpper() == "Sucess".ToUpper())
                        {
                            Console.WriteLine("ET_TABLE.Count  = " + Convert.ToString(returndatas.ET_TABLE.Count).ToUpper());
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

                                        QueryExpression checkOAnumberExist = new QueryExpression("hil_oaheader");
                                        checkOAnumberExist.ColumnSet = new ColumnSet("hil_oaheaderid");
                                        checkOAnumberExist.NoLock = true;
                                        checkOAnumberExist.Criteria = new FilterExpression(LogicalOperator.And);
                                        checkOAnumberExist.Criteria.AddCondition("hil_name", ConditionOperator.Equal, item.ORDERNUMBER);
                                        EntityCollection entColOA = service.RetrieveMultiple(checkOAnumberExist);
                                        if (entColOA.Entities.Count > 0)
                                        {
                                            headerid = entColOA.Entities[0].Id;
                                        }
                                        else
                                        {
                                            OAHeader["hil_name"] = Convert.ToString(item.ORDERNUMBER);
                                            OANumber = Convert.ToString(item.ORDERNUMBER);
                                            if (tender != null)
                                                OAHeader["hil_tenderid"] = tender.ToEntityReference();

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

                                            //Emailer Setup Mapping
                                            OAHeader["hil_omsemailersetup"] = new EntityReference("hil_tenderemailersetup", new Guid("b37bc5c4-6132-ec11-b6e6-002248d4ce49"));

                                            OAHeader["ownerid"] = tenderOwner;
                                            OAHeader["hil_customercode"] = item.CUST_CODE;
                                            OAHeader["hil_creditlimit"] = Convert.ToDecimal(item.CREDITLIMIT);
                                            OAHeader["hil_creditdayscustomercode"] = item.CREDITDAY_CUST;
                                            OAHeader["hil_creditdays"] = Convert.ToInt32(item.CREDITDAYS);
                                            lstoaheader.Add(_Oaheader);

                                            headerid = service.Create(OAHeader);
                                            Console.WriteLine("OA Header Created");
                                        }
                                    }
                                    QueryExpression QueryProductCode = new QueryExpression("product");
                                    QueryProductCode.ColumnSet = new ColumnSet("name");
                                    QueryProductCode.NoLock = true;
                                    QueryProductCode.Criteria = new FilterExpression(LogicalOperator.And);
                                    QueryProductCode.Criteria.AddCondition("productnumber", ConditionOperator.Equal, item.ITEM_CODE);
                                    EntityCollection entCol5 = service.RetrieveMultiple(QueryProductCode);

                                    QueryExpression checkOAProductExist = new QueryExpression("hil_oaproduct");
                                    checkOAProductExist.ColumnSet = new ColumnSet("hil_oaproductid");
                                    checkOAProductExist.NoLock = true;
                                    checkOAProductExist.Criteria = new FilterExpression(LogicalOperator.And);
                                    checkOAProductExist.Criteria.AddCondition("hil_oaheader", ConditionOperator.Equal, headerid);
                                    checkOAProductExist.Criteria.AddCondition("hil_productcode", ConditionOperator.Equal, entCol5[0].ToEntityReference().Id);
                                    checkOAProductExist.Criteria.AddCondition("hil_name", ConditionOperator.Equal, item.ZTENDER_LINE);

                                    EntityCollection entColOAProduct = service.RetrieveMultiple(checkOAProductExist);
                                    if (entColOAProduct.Entities.Count > 0)
                                    {
                                        Entity OAProduct = new Entity("hil_oaproduct", entColOAProduct[0].Id);
                                        OAProduct["hil_deliverydate"] = Convert.ToDateTime(item.DELIVERY_DATE);
                                        OAProduct["hil_name"] = item.ZTENDER_LINE;
                                        OAProduct["hil_materialcategorydescription"] = Convert.ToString(item.CATEGORY_DISC);
                                        OAProduct["hil_materialgroup"] = item.MATKL;
                                        OAProduct["hil_netvalue"] = new Money(Convert.ToDecimal(item.VALUE));
                                        if (entCol5.Entities.Count > 0)
                                        {
                                            OAProduct["hil_productcode"] = entCol5[0].ToEntityReference();
                                        }
                                        OAProduct["hil_productdescription"] = item.ITEM_DESC;
                                        OAProduct["hil_productionplanningremarks"] = "";
                                        OAProduct["hil_quantity"] = Convert.ToDecimal(item.QTY) * 1000;
                                        OAProduct["hil_stockinhand"] = Convert.ToDecimal(item.STOCK);
                                        if (tender != null)
                                            OAProduct["hil_tenderid"] = tender.ToEntityReference();
                                        OAProduct["hil_oaheader"] = new EntityReference(OAHeader.LogicalName, headerid);
                                        OAProduct["ownerid"] = tenderOwner;

                                        //Emailer Setup Mapping
                                        OAProduct["hil_omsemailersetup"] = new EntityReference("hil_tenderemailersetup", new Guid("b37bc5c4-6132-ec11-b6e6-002248d4ce49"));

                                        service.Update(OAProduct);
                                    }
                                    else
                                    {
                                        Entity OAProduct = new Entity("hil_oaproduct");
                                        OAProduct["hil_deliverydate"] = Convert.ToDateTime(item.DELIVERY_DATE);
                                        OAProduct["hil_name"] = item.ZTENDER_LINE;
                                        OAProduct["hil_materialcategorydescription"] = Convert.ToString(item.CATEGORY_DISC);
                                        OAProduct["hil_materialgroup"] = item.MATKL;
                                        OAProduct["hil_netvalue"] = new Money(Convert.ToDecimal(item.VALUE));
                                        //QueryExpression QueryProductCode = new QueryExpression("product");
                                        //QueryProductCode.ColumnSet = new ColumnSet("name");
                                        //QueryProductCode.NoLock = true;
                                        //QueryProductCode.Criteria = new FilterExpression(LogicalOperator.And);
                                        //QueryProductCode.Criteria.AddCondition("productnumber", ConditionOperator.Equal, item.ITEM_CODE);
                                        //EntityCollection entCol5 = service.RetrieveMultiple(QueryProductCode);
                                        if (entCol5.Entities.Count > 0)
                                        {
                                            OAProduct["hil_productcode"] = entCol5[0].ToEntityReference();
                                        }
                                        OAProduct["hil_productdescription"] = item.ITEM_DESC;
                                        OAProduct["hil_productionplanningremarks"] = "";
                                        OAProduct["hil_quantity"] = Convert.ToDecimal(item.QTY) * 1000;
                                        OAProduct["hil_stockinhand"] = Convert.ToDecimal(item.STOCK);
                                        if (tender != null)
                                            OAProduct["hil_tenderid"] = tender.ToEntityReference();
                                        OAProduct["hil_oaheader"] = new EntityReference(OAHeader.LogicalName, headerid);
                                        OAProduct["ownerid"] = tenderOwner;

                                        //Emailer Setup Mapping
                                        OAProduct["hil_omsemailersetup"] = new EntityReference("hil_tenderemailersetup", new Guid("b37bc5c4-6132-ec11-b6e6-002248d4ce49"));

                                        service.Create(OAProduct);
                                        QueryExpression OCLPrd = new QueryExpression("hil_orderchecklistproduct");
                                        OCLPrd.ColumnSet = new ColumnSet("hil_name");
                                        OCLPrd.NoLock = true;
                                        OCLPrd.Criteria = new FilterExpression(LogicalOperator.And);
                                        OCLPrd.Criteria.AddCondition("hil_name", ConditionOperator.Equal, item.ZTENDER_LINE);
                                        EntityCollection OCLCOLL = service.RetrieveMultiple(OCLPrd);
                                        Entity productitem = new Entity("hil_orderchecklistproduct");
                                        productitem.Id = OCLCOLL[0].Id;
                                        productitem["hil_syncwithsap"] = true;
                                        service.Update(productitem);
                                    }
                                    Console.WriteLine("OA Product Created");
                                    ///
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("OA Header ELSE");
                            Entity orderchecklist1 = new Entity("hil_orderchecklist");
                            orderchecklist1.Id = orderChecklistId;
                            orderchecklist1["hil_synwithsap"] = false;
                            orderchecklist1.Attributes["hil_syncwithsapremark"] = Convert.ToString(returndatas.RETURN);
                            service.Update(orderchecklist1);

                            string bodtText1 = mailBody(CustomerName, ProjectName, OANumber, EnqNo, orderchecklistno, "Failed", Convert.ToString(returndatas.RETURN));
                            sendEmail(service, bodtText1, orderchecklist1.ToEntityReference(), owner);

                        }
                        #endregion tt

                    }
                    catch (Exception ex)
                    {
                        Entity orderchecklist2 = new Entity("hil_orderchecklist");
                        orderchecklist2.Id = lstPO.Id;
                        orderchecklist2.Attributes["hil_syncwithsapremark"] = "Error :" + ex.Message;// Convert.ToString(returndatas.RETURN);
                        service.Update(orderchecklist2);

                        Console.WriteLine("Error :" + ex.Message);
                        RCL.IT_TABLE_ITEM = null;
                        continue;
                    }


                }
                #region plantwise all status
                if (OaSAPMessage.Length > 0)
                {
                    bool checkStatus = false;
                    bool checkErrorStatus = false;
                    StringBuilder responseError = new StringBuilder();
                    for (int j = 0; j < OaSAPMessage.Length; j++)
                    {
                        if (OaSAPMessage[j].ToUpper() == "Sucess".ToUpper())
                        {
                            checkStatus = true;
                        }
                        else
                        {
                            checkErrorStatus = true;
                            responseError.Append(OaSAPMessage[j]);
                            responseError.Append(", ");
                        }
                    }
                    if (checkErrorStatus == false)
                    {
                        Entity orderchecklist = new Entity("hil_orderchecklist");
                        orderchecklist.Id = orderChecklistId;
                        orderchecklist["hil_synwithsap"] = true;
                        orderchecklist.Attributes["hil_syncwithsapremark"] = Convert.ToString("Sucess");
                        service.Update(orderchecklist);

                        string bodtText = mailBody(CustomerName, ProjectName, OANumber, EnqNo, orderchecklistno, "successful", Convert.ToString("Sucess"));
                        sendEmail(service, bodtText, orderchecklist.ToEntityReference(), owner);
                        Console.WriteLine("Mail Send to USER");
                    }
                    else
                    {
                        Entity orderchecklist1 = new Entity("hil_orderchecklist");
                        orderchecklist1.Id = orderChecklistId;
                        orderchecklist1["hil_synwithsap"] = false;
                        orderchecklist1.Attributes["hil_syncwithsapremark"] = Convert.ToString(responseError);
                        service.Update(orderchecklist1);

                        string bodtText1 = mailBody(CustomerName, ProjectName, OANumber, EnqNo, orderchecklistno, "Failed", Convert.ToString(responseError));
                        sendEmail(service, bodtText1, orderchecklist1.ToEntityReference(), owner);
                        Console.WriteLine("Mail Send to USER");
                    }
                }
                #endregion plantwise all status
            }
            Console.WriteLine("OA Sync End");
        }
        public static void changeAutoNumber(string tenderID)
        {

            String fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='hil_tender'>
                                    <attribute name='hil_tenderid' />
                                    <attribute name='hil_name' />
                                    <attribute name='createdon' />
                                    <order attribute='hil_name' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='hil_tenderid' operator='in'>
                                        <value uiname='TEND00003671' uitype='hil_tender'>{A73AB6E9-6D4E-EC11-8F8E-6045BD7332D5}</value>
                                        <value uiname='TEND00003650' uitype='hil_tender'>{8097E0D5-B94D-EC11-8F8E-6045BD733B91}</value>
                                        <value uiname='TEND00003637' uitype='hil_tender'>{9EC5CB9F-174D-EC11-8F8E-6045BD733F58}</value>
                                        <value uiname='TEND00003549' uitype='hil_tender'>{C2EE8E8C-584B-EC11-8F8E-6045BD733B84}</value>
                                        <value uiname='TEND00003540' uitype='hil_tender'>{6947462D-484B-EC11-8F8E-6045BD7333DC}</value>
                                        <value uiname='TEND00003528' uitype='hil_tender'>{691EC13D-3549-EC11-8C62-002248D48AEE}</value>
                                        <value uiname='TEND00003523' uitype='hil_tender'>{B5496745-2849-EC11-8C62-002248D483A3}</value>
                                        <value uiname='TEND00003522' uitype='hil_tender'>{E0BFD335-2749-EC11-8C62-002248D48B6C}</value>
                                        <value uiname='TEND00003518' uitype='hil_tender'>{E35FC62C-1D49-EC11-8C62-002248D48636}</value>
                                        <value uiname='TEND00003511' uitype='hil_tender'>{DFF7E827-0849-EC11-8C62-002248D48615}</value>
                                        <value uiname='TEND00003510' uitype='hil_tender'>{AFD29C13-0749-EC11-8C62-002248D48636}</value>
                                        <value uiname='TEND00003505' uitype='hil_tender'>{9FE0C32B-0349-EC11-8C62-002248D481FC}</value>
                                        <value uiname='TEND00003495' uitype='hil_tender'>{EBF5B5DD-EE48-EC11-8C62-002248D481FC}</value>
                                        <value uiname='TEND00003490' uitype='hil_tender'>{DB72843F-7C48-EC11-8C62-002248D483A3}</value>
                                        <value uiname='TEND00003479' uitype='hil_tender'>{03DB1741-5148-EC11-8C62-002248D483A3}</value>
                                        <value uiname='TEND00003472' uitype='hil_tender'>{0EC64016-3E48-EC11-8C62-0022486E2560}</value>
                                        <value uiname='TEND00003468' uitype='hil_tender'>{F22B835C-3848-EC11-8C62-0022486E2560}</value>
                                        <value uiname='TEND00003437' uitype='hil_tender'>{4F7A30E9-6D47-EC11-8C62-002248D48615}</value>
                                        <value uiname='TEND00003433' uitype='hil_tender'>{4E2E5429-6447-EC11-8C62-002248D483A3}</value>
                                        <value uiname='TEND00003413' uitype='hil_tender'>{EAD6AEB8-A846-EC11-8C62-002248D48AEE}</value>
                                        <value uiname='TEND00003412' uitype='hil_tender'>{4EC5FE9E-A646-EC11-8C62-002248D48AEE}</value>
                                        <value uiname='TEND00003389' uitype='hil_tender'>{85596DB1-0746-EC11-8C62-002248D48615}</value>
                                        <value uiname='TEND00003384' uitype='hil_tender'>{5B9E21C6-F645-EC11-8C62-002248D48B6C}</value>
                                        <value uiname='TEND00003372' uitype='hil_tender'>{A717AE75-E145-EC11-8C62-002248D48636}</value>
                                        <value uiname='TEND00003370' uitype='hil_tender'>{7DEC2EC1-DB45-EC11-8C62-002248D48615}</value>
                                        <value uiname='TEND00003332' uitype='hil_tender'>{C2051AA2-8243-EC11-8C62-002248D4CCFD}</value>
                                        <value uiname='TEND00003263' uitype='hil_tender'>{A5BC3757-5A41-EC11-8C62-002248D4A375}</value>
                                        <value uiname='TEND00003258' uitype='hil_tender'>{5AAD1175-4C41-EC11-8C62-002248D4C8E1}</value>
                                        <value uiname='TEND00003232' uitype='hil_tender'>{D079EC12-8F40-EC11-8C62-002248D4C61A}</value>
                                        <value uiname='TEND00003224' uitype='hil_tender'>{10C848F3-6C40-EC11-8C62-002248D4C61A}</value>
                                        <value uiname='TEND00003220' uitype='hil_tender'>{C83F0502-6740-EC11-8C62-002248D4C635}</value>
                                        <value uiname='TEND00003218' uitype='hil_tender'>{E8772C58-6240-EC11-8C62-002248D4C635}</value>
                                        <value uiname='TEND00003215' uitype='hil_tender'>{A46567FA-5F40-EC11-8C62-002248D4CACF}</value>
                                        <value uiname='TEND00003188' uitype='hil_tender'>{2D9A4690-7F3C-EC11-8C62-6045BD73244E}</value>
                                        <value uiname='TEND00003186' uitype='hil_tender'>{559E47C1-7A3C-EC11-8C62-6045BD73244E}</value>
                                        <value uiname='TEND00003173' uitype='hil_tender'>{31FB7BF4-C63B-EC11-8C62-6045BD732F9A}</value>
                                        <value uiname='TEND00003164' uitype='hil_tender'>{B8A0EF36-A83B-EC11-8C62-6045BD73244E}</value>
                                        <value uiname='TEND00003157' uitype='hil_tender'>{2DF60141-943B-EC11-8C62-6045BD732318}</value>
                                        <value uiname='TEND00003150' uitype='hil_tender'>{3B21DEBF-133B-EC11-8C62-6045BD732F6B}</value>
                                        <value uiname='TEND00003132' uitype='hil_tender'>{85D3D38D-E83A-EC11-8C62-6045BD732F6B}</value>
                                        <value uiname='TEND00003131' uitype='hil_tender'>{A9F90339-E73A-EC11-8C62-6045BD732316}</value>
                                        <value uiname='TEND00003116' uitype='hil_tender'>{8D8A34C5-6A39-EC11-8C64-000D3AF0735F}</value>
                                        <value uiname='TEND00003115' uitype='hil_tender'>{593FE048-6339-EC11-8C64-000D3AF07046}</value>
                                        <value uiname='TEND00003110' uitype='hil_tender'>{43C431CC-5239-EC11-8C64-000D3AF0735F}</value>
                                        <value uiname='TEND00003108' uitype='hil_tender'>{96416DD6-4E39-EC11-8C64-000D3AF07EE3}</value>
                                        <value uiname='TEND00003092' uitype='hil_tender'>{551BBFEE-B138-EC11-8C64-000D3AF07E49}</value>
                                        <value uiname='TEND00003080' uitype='hil_tender'>{0C6B2B3F-8E38-EC11-8C64-000D3AF0735F}</value>
                                        <value uiname='TEND00003079' uitype='hil_tender'>{3D4B5ECA-8C38-EC11-8C64-000D3AF0735F}</value>
                                        <value uiname='TEND00003078' uitype='hil_tender'>{3FCB6ECC-8A38-EC11-8C64-000D3AF0670D}</value>
                                        <value uiname='TEND00003076' uitype='hil_tender'>{063BD92D-8A38-EC11-8C64-000D3AF0735F}</value>
                                        <value uiname='TEND00003074' uitype='hil_tender'>{C98D5744-8638-EC11-8C64-000D3AF0735F}</value>
                                        <value uiname='TEND00003054' uitype='hil_tender'>{90EB0935-E137-EC11-8C64-000D3AF0735F}</value>
                                        <value uiname='TEND00003030' uitype='hil_tender'>{3A562EB1-A737-EC11-8C64-000D3AF079B1}</value>
                                        <value uiname='TEND00002950' uitype='hil_tender'>{1AE15529-4D35-EC11-8C64-000D3AF067EC}</value>
                                        <value uiname='TEND00002848' uitype='hil_tender'>{1AF93D4B-ED30-EC11-B6E6-002248D4CE49}</value>
                                        <value uiname='TEND00002828' uitype='hil_tender'>{889A4DC2-A330-EC11-B6E6-002248D4C6D7}</value>
                                        <value uiname='TEND00002800' uitype='hil_tender'>{42543F0A-D02F-EC11-B6E6-002248D4CAD3}</value>
                                        <value uiname='TEND00002733' uitype='hil_tender'>{DE4D580C-912A-EC11-B6E6-6045BD72FAED}</value>
                                        <value uiname='TEND00002673' uitype='hil_tender'>{3DC7BC55-8D27-EC11-B6E6-6045BD72C90B}</value>
                                        <value uiname='TEND00002665' uitype='hil_tender'>{1D37B3ED-6827-EC11-B6E6-6045BD72CC32}</value>
                                        <value uiname='TEND00002594' uitype='hil_tender'>{353CBAC2-2F25-EC11-B6E6-6045BD72CEAA}</value>
                                        <value uiname='TEND00002514' uitype='hil_tender'>{29C5D038-DA21-EC11-B6E6-6045BD72D8A6}</value>
                                        <value uiname='TEND00002431' uitype='hil_tender'>{AD6A20F4-6F1F-EC11-B6E6-6045BD7311EE}</value>
                                        <value uiname='TEND00002363' uitype='hil_tender'>{3EBA8246-801C-EC11-B6E7-000D3A3E4DB5}</value>
                                        <value uiname='TEND00002210' uitype='hil_tender'>{AA2F8513-9017-EC11-B6E7-6045BD72DDB8}</value>
                                        <value uiname='TEND00002159' uitype='hil_tender'>{B418BAD9-FB15-EC11-B6E7-6045BD72D872}</value>
                                        <value uiname='TEND00002133' uitype='hil_tender'>{3733CC68-3215-EC11-B6E7-6045BD72D872}</value>
                                        <value uiname='TEND00002102' uitype='hil_tender'>{75194672-7D14-EC11-B6E7-6045BD72DDB8}</value>
                                        <value uiname='TEND00002034' uitype='hil_tender'>{03651CC8-4C11-EC11-B6E7-000D3AF0DF47}</value>
                                        <value uiname='TEND00001819' uitype='hil_tender'>{BB7A90DD-8409-EC11-B6E6-6045BD72FAF3}</value>
                                        <value uiname='TEND00001775' uitype='hil_tender'>{ADAF6890-0F07-EC11-B6E7-000D3AF0B248}</value>
                                        <value uiname='TEND00001634' uitype='hil_tender'>{4F649930-8801-EC11-94EF-6045BD72E516}</value>
                                        <value uiname='TEND00001478' uitype='hil_tender'>{F698BBE0-7DFB-EB11-94EF-6045BD72EAD7}</value>
                                        <value uiname='TEND00001458' uitype='hil_tender'>{23275131-A1FA-EB11-94EF-6045BD72E7EE}</value>
                                        <value uiname='TEND00001188' uitype='hil_tender'>{E262CD1A-71EF-EB11-94EF-000D3A3E22AA}</value>
                                        <value uiname='TEND00001183' uitype='hil_tender'>{19E40F5E-69EF-EB11-94EF-000D3A3E5C3D}</value>
                                        <value uiname='TEND00000577' uitype='hil_tender'>{2E0B6382-72D5-EB11-BACC-6045BD725CD8}</value>
                                      </condition>
                                    </filter>
                                  </entity>
                                </fetch>";

            EntityCollection TenderEntCol = service.RetrieveMultiple(new FetchExpression(fetch));
            foreach (Entity tender in TenderEntCol.Entities)
            {

                QueryExpression qsChecklist = new QueryExpression("hil_tenderproduct");
                qsChecklist.ColumnSet = new ColumnSet("hil_name", "hil_tenderid");
                qsChecklist.Criteria = new FilterExpression(LogicalOperator.And);
                qsChecklist.Criteria.AddCondition("hil_tenderid", ConditionOperator.Equal, tender.Id);
                EntityCollection entCol = service.RetrieveMultiple(qsChecklist);
                int count = entCol.Entities.Count;
                Console.WriteLine(count);
                int i = 1;
                foreach (Entity prod in entCol.Entities)
                {
                    string _tend = prod.GetAttributeValue<EntityReference>("hil_tenderid").Name;
                    string name = prod.GetAttributeValue<string>("hil_name");
                    Entity prd = new Entity(prod.LogicalName);
                    prd.Id = prod.Id;
                    Console.WriteLine(_tend + "_" + i.ToString().PadLeft(3, '0'));
                    prd["hil_name"] = _tend + "_" + i.ToString().PadLeft(3, '0');
                    service.Update(prd);
                    i++;
                    QueryExpression qsChecklist1 = new QueryExpression("hil_deliveryschedule");
                    qsChecklist1.ColumnSet = new ColumnSet("hil_name", "hil_tenderproduct");
                    qsChecklist1.Criteria = new FilterExpression(LogicalOperator.And);
                    qsChecklist1.Criteria.AddCondition("hil_tenderproduct", ConditionOperator.Equal, prod.Id);
                    EntityCollection entCol1 = service.RetrieveMultiple(qsChecklist1);
                    foreach (Entity prod1 in entCol1.Entities)
                    {
                        string _tndprd = prod1.GetAttributeValue<EntityReference>("hil_tenderproduct").Name;
                        string name1 = prod1.GetAttributeValue<string>("hil_name");
                        int j = 1;
                        Entity prd1 = new Entity(prod1.LogicalName);
                        prd1.Id = prod1.Id;
                        prd1["hil_name"] = _tndprd + "_" + j.ToString().PadLeft(3, '0');
                        j++;
                        service.Update(prd1);
                    }
                }
            }
        }
        static dynamic orderChecklistId;
        public static string dateFormatter(DateTime date)
        {
            string _enquiryDatetime = string.Empty;
            if (date.Year.ToString().PadLeft(4, '0') != "0001")
                _enquiryDatetime = date.Year.ToString() + "-" + date.Month.ToString().PadLeft(2, '0') + "-" + date.Day.ToString().PadLeft(2, '0');
            return _enquiryDatetime;
        }
        public static IOrganizationService LoadAppSettings()
        {
            try
            {
                _userId = ConfigurationManager.AppSettings["CrmUserId"].ToString();
                _password = ConfigurationManager.AppSettings["CrmUserPassword"].ToString();
                _soapOrganizationServiceUri = ConfigurationManager.AppSettings["CrmSoapOrganizationServiceUri"].ToString();
                ConnectToCRM();

            }
            catch (Exception ex)
            {
                Console.WriteLine("D365AuditLogMigration.Program.Main.LoadAppSettings ::  Error While Loading App Settings:" + ex.Message.ToString());
            }
            return service;
        }
        static void ConnectToCRM()
        {
            try
            {
                ClientCredentials credentials = new ClientCredentials();
                credentials.UserName.UserName = _userId;
                credentials.UserName.Password = _password;
                Uri serviceUri = new Uri(_soapOrganizationServiceUri);
                Console.WriteLine("URL " + serviceUri);
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                OrganizationServiceProxy proxy = new OrganizationServiceProxy(serviceUri, null, credentials, null);
                proxy.EnableProxyTypes();
                service = (IOrganizationService)proxy;
                loginUserGuid = ((WhoAmIResponse)service.Execute(new WhoAmIRequest())).UserId;
            }
            catch (Exception ex)
            {
                Console.WriteLine("SAP_IntegrationForOrderCreation.Program.Main.ConnectToCRM :: Error While Creating Connection with MS CRM Organisation:"
                    + ex.Message.ToString());
            }
        }
        public static void sendEmail(IOrganizationService _service, string mailbody, EntityReference ordercheklistid, EntityReference owner, string status, string orderchecklistno, string CustomerName, string ProjectName)
        {
            Entity entEmail = new Entity("email");
            Entity entFrom = new Entity("activityparty");
            entFrom["partyid"] = new EntityReference("queue", new Guid("a81ac51b-c9df-eb11-bacb-0022486ea537"));
            Entity[] entFromList = { entFrom };
            entEmail["from"] = entFromList;
            EntityReference to = owner;//new EntityReference("systemuser", new Guid("803d62d9-5e34-eb11-bf68-000d3af05a1b"));
            Entity toActivityParty = new Entity("activityparty");
            toActivityParty["partyid"] = to;
            entEmail["to"] = new Entity[] { toActivityParty };
            if (status == "successful")
            {
                entEmail["subject"] = @"OA Successfully Created for Order Check list No :" + orderchecklistno + " Customer Name : " + CustomerName + " Project name :" + ProjectName;
            }
            else
            {
                entEmail["subject"] = @"Opps Order is not created for PO Check List Ref No :" + orderchecklistno + " Customer Name : " + CustomerName + " Project name :" + ProjectName;
            }
            entEmail["description"] = mailbody;
            entEmail["regardingobjectid"] = ordercheklistid;
            Guid emailId = _service.Create(entEmail);
            SendEmailRequest sendEmailReq = new SendEmailRequest()
            {
                EmailId = emailId,
                IssueSend = true
            };
            SendEmailResponse sendEmailRes = (SendEmailResponse)_service.Execute(sendEmailReq);
        }
        public static string mailBody(string CustomerName, string ProjectName, string OANumber, string EnqNo, string orderchecklistno, string status, string returnmessage)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<div><p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;'><span> Hi There </span></p>");
            sb.Append("<p style='margin-bottom:.0001pt;'></p>");
            sb.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;'>");
            sb.Append("</p><p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;' ><span></span></p>");
            if (status == "successful")
            {
                sb.Append("<span> Congratulations Order is successfully created in SAP. The OA Ref No: " + OANumber + " For Order Checklist Request with Ref No: " + orderchecklistno + ", Enq No.: " + EnqNo + " , Customer Name: " + CustomerName + " , Project Name :" + ProjectName + " </span>");
            }
            else
            {
                sb.Append("<span> Order is not created in SAP against PO Checklist No : " + orderchecklistno + ", Customer Name : " + CustomerName + ", Project Name : " + ProjectName);
                sb.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;'></p>");
                sb.Append("<span>Below are the Error Details:</span>");
                sb.Append("<p style='margin-bottom:.0001pt;'></p>");
                sb.Append("<span style='font-weight:bold;font-size:16.0pt'>" + returnmessage + "</span>");

            }

            sb.Append("</p><p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;' ><span></span></p>");
            sb.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;'><span>");
            sb.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;'><span> Regards </span></p>");
            sb.Append("<p style='margin-bottom:.0001pt;font-size:11.0pt;font-family:Calibri,sans-serif;'><span> System </span></p></div> ");

            string resultBody = sb.ToString();
            return resultBody;
        }
        public static void sendEmail(IOrganizationService _service, string mailbody, EntityReference ordercheklistid, EntityReference owner)
        {
            Entity entEmail = new Entity("email");
            Entity entFrom = new Entity("activityparty");
            entFrom["partyid"] = new EntityReference("queue", new Guid("a81ac51b-c9df-eb11-bacb-0022486ea537"));
            Entity[] entFromList = { entFrom };
            entEmail["from"] = entFromList;
            EntityReference to = owner;//new EntityReference("systemuser", new Guid("803d62d9-5e34-eb11-bf68-000d3af05a1b"));
            Entity toActivityParty = new Entity("activityparty");
            toActivityParty["partyid"] = to;
            entEmail["to"] = new Entity[] { toActivityParty };
            entEmail["subject"] = @"OA Generation Status"; ;
            entEmail["description"] = mailbody;
            entEmail["regardingobjectid"] = ordercheklistid;
            Guid emailId = _service.Create(entEmail);
            SendEmailRequest sendEmailReq = new SendEmailRequest()
            {
                EmailId = emailId,
                IssueSend = true
            };
            SendEmailResponse sendEmailRes = (SendEmailResponse)_service.Execute(sendEmailReq);
        }
        private static EntityReference GetOALineRef(IOrganizationService _service, string _tenderLineNo)
        {
            EntityReference efOAProduct = null;
            try
            {
                QueryExpression Query = new QueryExpression("hil_oaproduct");
                Query.ColumnSet = new ColumnSet(false);
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, _tenderLineNo);
                EntityCollection Found = _service.RetrieveMultiple(Query);
                if (Found.Entities.Count > 0)
                {
                    efOAProduct = Found.Entities[0].ToEntityReference();
                }
            }
            catch { }
            return efOAProduct;
        }
        public static void GetReadinessDates(IOrganizationService _service)
        {
            if (_service != null)
            {
                try
                {
                    Console.WriteLine("GetReadinessDates Started");
                    IntegrationConfig intConfig = Models.IntegrationConfiguration(_service, "GetCableReadinessData");
                    string url = intConfig.uri;
                    string authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(intConfig.Auth));

                    var client = new RestClient(string.Format(url, "R"));
                    client.Timeout = -1;

                    var request = new RestRequest(Method.GET);
                    request.AddHeader("authorization", "Basic " + authInfo);
                    request.AddHeader("Content-Type", "application/json");
                    //request.AddParameter("application/json", , ParameterType.RequestBody);
                    IRestResponse response = client.Execute(request);

                    List<LTTABLE> table = new List<LTTABLE>();
                    table = JsonConvert.DeserializeObject<OutputClass>(response.Content).ET_TABLE;
                    Console.WriteLine("Totla Count: " + table.Count);
                    var i = 1;
                    int counter = 0;
                    for (; table.Count > counter;)
                    {
                        LTTABLE row = table[counter];
                        try
                        {
                            if (row.ZTENDER_LINE != "" && row.ZTENDERNO != "")
                            {
                                Entity entOAReadinessDates = new Entity("hil_oareadinessdatestaging");
                                if (row.INSP_DATE != "")
                                    entOAReadinessDates["hil_insp_date"] = Convert.ToDateTime(row.INSP_DATE);// Date
                                if (row.MODIFY_DATE != "")
                                    entOAReadinessDates["hil_modify_date"] = Convert.ToDateTime(row.MODIFY_DATE);// Date
                                entOAReadinessDates["hil_modifyby"] = row.MODIFYBY;
                                entOAReadinessDates["hil_name"] = row.MTIMESTAMP;
                                if (row.READDATE1 != "")
                                    entOAReadinessDates["hil_readdate1"] = Convert.ToDateTime(row.READDATE1);// Date
                                if (row.READDATE2 != "")
                                    entOAReadinessDates["hil_readdate2"] = Convert.ToDateTime(row.READDATE2);// Date
                                if (row.READDATE3 != "")
                                    entOAReadinessDates["hil_readdate3"] = Convert.ToDateTime(row.READDATE3);// Date

                                entOAReadinessDates["hil_ztender_line"] = row.ZTENDER_LINE;
                                entOAReadinessDates["hil_ztenderno"] = row.ZTENDERNO;
                                entOAReadinessDates["hil_oaproductno"] = GetOALineRef(_service, row.ZTENDER_LINE);// Date
                                _service.Create(entOAReadinessDates);
                                Console.WriteLine("OA Readiness dates created for: " + i.ToString());
                                i++;
                            }
                            counter++;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error " + ex.Message);
                            try
                            {
                                Program.loginUserGuid = ((WhoAmIResponse)_service.Execute(new WhoAmIRequest())).UserId;
                            }
                            catch
                            {
                                try
                                {
                                    ClientCredentials credentials = new ClientCredentials();
                                    credentials.UserName.UserName = Program._userId;
                                    credentials.UserName.Password = Program._password;
                                    Uri serviceUri = new Uri(Program._soapOrganizationServiceUri);
                                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                                    ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                                    OrganizationServiceProxy proxy = new OrganizationServiceProxy(serviceUri, null, credentials, null);
                                    proxy.EnableProxyTypes();
                                    _service = (IOrganizationService)proxy;
                                    Program.loginUserGuid = ((WhoAmIResponse)_service.Execute(new WhoAmIRequest())).UserId;
                                }
                                catch
                                {
                                    Console.WriteLine("Service not created....");
                                }
                                counter--;
                            }
                        }
                    }
                    client = new RestClient(string.Format(url, "U"));
                    client.Timeout = -1;
                    request = new RestRequest(Method.GET);
                    request.AddHeader("authorization", "Basic " + authInfo);
                    request.AddHeader("Content-Type", "application/json");
                    response = client.Execute(request);
                    Console.WriteLine("Done");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error  " + ex.Message);
                    Console.ReadLine();
                }
                Console.WriteLine("GetReadinessDates Ended");
            }
        }
        public static string genrateDeliveryDate(IOrganizationService _service, EntityReference product)
        {
            string dayss = null;
            try
            {

                Entity oclPrd = _service.Retrieve(product.LogicalName, product.Id, new ColumnSet("hil_orderchecklistid", "hil_plantcode"));
                EntityReference OCL = oclPrd.GetAttributeValue<EntityReference>("hil_orderchecklistid");
                EntityReference plant = oclPrd.GetAttributeValue<EntityReference>("hil_plantcode");
                Entity oclEntty = _service.Retrieve(OCL.LogicalName, OCL.Id, new ColumnSet("hil_typeoforder"));
                int typeofOrder = oclEntty.GetAttributeValue<OptionSetValue>("hil_typeoforder").Value;

                QueryExpression Query = new QueryExpression("hil_deliveryschedulemaster");
                Query.ColumnSet = new ColumnSet("hil_deliveryday");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_ordertype", ConditionOperator.Equal, typeofOrder);

                LinkEntity EntityB = new LinkEntity("hil_orderchecklistid", "hil_hil_deliveryschedulemaster_hil_plantmas", "hil_deliveryschedulemasterid", "hil_deliveryschedulemasterid", JoinOperator.Inner);
                EntityB.Columns = new ColumnSet(false);
                EntityB.LinkCriteria = new FilterExpression(LogicalOperator.And);
                EntityB.LinkCriteria.AddCondition("hil_plantmasterid", ConditionOperator.Equal, plant.Id);
                EntityB.EntityAlias = "Plant";
                Query.LinkEntities.Add(EntityB);
                EntityCollection Found = _service.RetrieveMultiple(Query);

                if (Found.Entities.Count == 1)
                {
                    Entity _oclPrd = new Entity(product.LogicalName);
                    _oclPrd.Id = product.Id;
                    int days = Found[0].GetAttributeValue<int>("hil_deliveryday");
                    _oclPrd["hil_deliverydate"] = DateTime.Now.AddDays(days);
                    service.Update(_oclPrd);
                    dayss = dateFormatter(DateTime.Now.AddDays(days)).ToString();
                }
                else
                {
                    Query = new QueryExpression("hil_deliveryschedulemaster");
                    Query.ColumnSet = new ColumnSet("hil_deliveryday");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_ordertype", ConditionOperator.Equal, typeofOrder);
                    Found = _service.RetrieveMultiple(Query);
                    if (Found.Entities.Count == 1)
                    {
                        Entity _oclPrd = new Entity(product.LogicalName);
                        _oclPrd.Id = product.Id;
                        int days = Found[0].GetAttributeValue<int>("hil_deliveryday");
                        _oclPrd["hil_deliverydate"] = DateTime.Now.AddDays(days);
                        service.Update(_oclPrd);
                        dayss = dateFormatter(DateTime.Now.AddDays(days)).ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error " + ex.Message);
            }

            return dayss;
        }

    }
}