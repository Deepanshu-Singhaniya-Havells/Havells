using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.ServiceModel.Description;
using System.Text;

namespace Havells.CRM.WebJob.BusinessManagement.OABooking
{
    public class Program
    {
        static dynamic orderChecklistId;
        static void Main(string[] args)
        {
            Console.WriteLine("Program Started");
            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmURL = "https://havellscrmdev1.crm8.dynamics.com"; //ConfigurationManager.AppSettings["CRMUrl"].ToString();
            string finalString = string.Format(connStr, CrmURL);
            IOrganizationService service = HavellsConnection.CreateConnection.createConnection(finalString);
            Console.WriteLine("Connection is Established..");
            //GetReadinessDates(service);
            OABookingPushtoSAP(service, args);
            Console.WriteLine("Program Terminated Secussfully");
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
                }
                Console.WriteLine("GetReadinessDates Ended");
            }
        }


        private static void OABookingPushtoSAP(IOrganizationService _service, string[] args)
        {
            QueryExpression qsChecklist = new QueryExpression("hil_orderchecklist");
            qsChecklist.ColumnSet = new ColumnSet("ownerid", "hil_name", "hil_tenderno", "hil_typeoforder", "hil_nameofclientcustomercode", "hil_paymentterms", "hil_approvedtransporter", "hil_projectname",
                "hil_poloinofooter", "hil_poloino", "hil_podate", "hil_prices", "hil_paymentterms", "hil_clearedformanufacturing", "hil_approveddatasheetgtp",
                "hil_approvedqap", "hil_inspection", "hil_ldterms", "hil_approvedtransporter", "hil_bankguaranteerequired", "hil_valueofbg", "hil_warrantyperiod",
                "hil_clearedon", "hil_drumlengthschedule", "hil_inspectionagency", "hil_usage", "hil_leadcode", "hil_specialinstructions", "hil_typeofdrum", "hil_consigneeaddress", "hil_taxtype");
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
            EntityCollection entCol = _service.RetrieveMultiple(qsChecklist);
            Console.WriteLine("Count of OCL " + entCol.Entities.Count);
            if (entCol.Entities.Count > 0)
            {
                GetOANumberPlantWise(entCol, _service);
            }
        }
        public static void GetOANumberPlantWise(EntityCollection entCol, IOrganizationService service)
        {
            Console.WriteLine("Total Record " + entCol.Entities.Count);
            EntityReference DepartmentRef = new EntityReference();
            string Department = string.Empty;//added for department 03/08/2022
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

                if (lstPO.Contains("hil_taxtype"))
                {
                    string taxType = lstPO.FormattedValues["hil_taxtype"];
                    string[] taxTypeArr = taxType.Split('-');
                    taxType = taxTypeArr[1];
                    iTTABLEHEADER.ZTAXTYPE = taxType;

                }
                if (lstPO.Contains("hil_clearedformanufacturing"))
                {
                    iTTABLEHEADER.ZMANUFCT_CLEAR = lstPO.GetAttributeValue<bool>("hil_clearedformanufacturing") == true ? "YES" : "NO";
                }
                else
                {
                    iTTABLEHEADER.ZMANUFCT_CLEAR = "NO";
                }
                string zusage = ORDERTYPE == "DS" ? "S" : "P";
                if (lstPO.Contains("hil_usage"))
                {
                   EntityReference usages =  lstPO.GetAttributeValue<EntityReference>("hil_usage");
                    zusage = service.Retrieve(usages.LogicalName, usages.Id, new ColumnSet("hil_code")).GetAttributeValue<string>("hil_code");
                }
                iTTABLEHEADER.ZUSAGE = zusage;//usase indicator master
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
                iTTABLEHEADER.ZZLEADCOD = lstPO.Contains("hil_leadcode") ? lstPO.GetAttributeValue<string>("hil_leadcode") : "NA";

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

                QueryExpression query = new QueryExpression("hil_orderchecklistproduct");
                query.ColumnSet = new ColumnSet("hil_plantcode");
                query.Distinct = true;
                query.Criteria = new FilterExpression(LogicalOperator.And);
                query.Criteria.AddCondition(new ConditionExpression("hil_orderchecklistid", ConditionOperator.Equal, lstPO.Id));
                query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                if (lstPO.GetAttributeValue<OptionSetValue>("hil_typeoforder").Value == 1)
                {
                    query.Criteria.AddCondition(new ConditionExpression("hil_quantity", ConditionOperator.GreaterThan, 0));
                }
                else
                {
                    query.Criteria.AddCondition(new ConditionExpression("hil_poqty", ConditionOperator.GreaterThan, 0));
                }

                EntityCollection entColdelv = service.RetrieveMultiple(query);

                String[] OaSAPMessage = new String[entColdelv.Entities.Count];
                int i = 0;
                Console.WriteLine("total Active Product Retrived count " + entColdelv.Entities.Count);
                QueryExpression querylines = null;
                foreach (Entity distPlant in entColdelv.Entities)
                {

                    try
                    {
                        List<ITTABLEITEM> iTTABLEITEMLst = new List<ITTABLEITEM>();

                        querylines = new QueryExpression("hil_orderchecklistproduct");
                        querylines.ColumnSet = new ColumnSet(true);
                        querylines.Criteria = new FilterExpression(LogicalOperator.And);
                        querylines.Criteria.AddCondition(new ConditionExpression("hil_orderchecklistid", ConditionOperator.Equal, lstPO.Id));
                        querylines.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                        querylines.Criteria.AddCondition(new ConditionExpression("hil_plantcode", ConditionOperator.Equal, distPlant.GetAttributeValue<EntityReference>("hil_plantcode").Id));
                        if (lstPO.GetAttributeValue<OptionSetValue>("hil_typeoforder").Value == 1)
                        {
                            querylines.Criteria.AddCondition(new ConditionExpression("hil_quantity", ConditionOperator.GreaterThan, 0));
                        }
                        else
                        {
                            querylines.Criteria.AddCondition(new ConditionExpression("hil_poqty", ConditionOperator.GreaterThan, 0));
                        }

                        EntityCollection entColl = service.RetrieveMultiple(querylines);
                        foreach (Entity product in entColl.Entities)
                        {
                            ITTABLEITEM iTTABLEITEM = new ITTABLEITEM();

                            iTTABLEITEM.ZTENDERNO = lstPO.Contains("hil_tenderno") ? lstPO.GetAttributeValue<EntityReference>("hil_tenderno").Name :
                                product.GetAttributeValue<EntityReference>("hil_orderchecklistid").Name;

                            iTTABLEITEM.ZMATNR = product.Contains("hil_product") ? product.GetAttributeValue<EntityReference>("hil_product").Name : "";

                            var quantity = product.Contains("hil_poqty") ? product.GetAttributeValue<decimal>("hil_poqty").ToString() : product.GetAttributeValue<decimal>("hil_quantity").ToString();

                            iTTABLEITEM.ZQUANTITY = String.Format("{0:0.00}", quantity);


                            DepartmentRef = product.GetAttributeValue<EntityReference>("hil_department");
                            //03/08/22  iTTABLEITEM.ZUOM = "M"; 
                            Department = product.GetAttributeValue<EntityReference>("hil_department").Name.ToLower();//addedd below four line for terms of unit 03/08/2022


                            if (Department.ToLower() == "cable")
                                iTTABLEITEM.ZUOM = "M";
                            else if (Department.ToLower() == "motor")
                                iTTABLEITEM.ZUOM = "NOS";

                            iTTABLEITEM.ZWERKS = product.Contains("hil_plantcode") ? service.Retrieve(product.GetAttributeValue<EntityReference>("hil_plantcode").LogicalName,
                                product.GetAttributeValue<EntityReference>("hil_plantcode").Id, new ColumnSet("hil_plantname")).GetAttributeValue<string>("hil_plantname") : "";

                            iTTABLEITEM.ZNET_PRICE = product.Contains("hil_porate") ? String.Format("{0:0.00}", product.GetAttributeValue<Money>("hil_porate").Value) :
                                String.Format("{0:0.00}", product.GetAttributeValue<Money>("hil_basicpriceinrsmtr").Value);

                            iTTABLEITEM.ZZFRE = product.Contains("hil_freightcharges") ? String.Format("{0:0.00}", product.GetAttributeValue<Money>("hil_freightcharges").Value) : "";

                            if (Department.ToLower() == "cable")
                                iTTABLEITEM.ZLD_PENALTY = product.Contains("hil_inspectiontype") ?
                                    Convert.ToString(service.Retrieve("hil_inspectiontype", product.GetAttributeValue<EntityReference>("hil_inspectiontype").Id,
                                    new ColumnSet("hil_code")).GetAttributeValue<string>("hil_code")) : "";
                            else if (Department.ToLower() == "motor")
                                iTTABLEITEM.ZLD_PENALTY = "";
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
                                iTTABLEITEM.ZDELIVERY_DATE = deliverDate;
                                //throw new Exception("Delivery Date is not Valid");
                            }
                            iTTABLEITEM.ZDELIVERY_DATE = deliverDate;
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
                                        if (item.WERKS != null)
                                        {
                                            OAProduct["hil_plant"] = getPlantRef(item.WERKS, service);
                                        }
                                        if (entCol5.Entities.Count > 0)
                                        {
                                            OAProduct["hil_productcode"] = entCol5[0].ToEntityReference();
                                        }
                                        OAProduct["hil_productdescription"] = item.ITEM_DESC;
                                        OAProduct["hil_productionplanningremarks"] = "";


                                        //commented on 03/08/22 OAProduct["hil_quantity"] = Convert.ToDecimal(item.QTY) * 1000;
                                        //if (Department.ToLower() == "cable") //added below four line for qty department wise precondition
                                        //    OAProduct["hil_quantity"] = Convert.ToDecimal(item.QTY) * 1000;
                                        //else if (Department.ToLower() == "motor")
                                        //    OAProduct["hil_quantity"] = Convert.ToDecimal(item.QTY);

                                        //chnaged by Saurabh for Conversion Factor
                                        int conversionFactor = service.Retrieve(DepartmentRef.LogicalName, DepartmentRef.Id, new ColumnSet("hil_conversionfactor")).GetAttributeValue<int>("hil_conversionfactor");
                                        if (conversionFactor != 0)
                                            OAProduct["hil_quantity"] = Convert.ToDecimal(item.QTY) * conversionFactor;
                                        else
                                            OAProduct["hil_quantity"] = Convert.ToDecimal(item.QTY);
                                        //chnaged end Here by Saurabh for Conversion Factor

                                        Console.WriteLine("Conversion Factor is " + conversionFactor);
                                        Console.WriteLine("########################################################################################");


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
                                        if (item.DELIVERY_DATE != null && item.DELIVERY_DATE != "")
                                        {
                                            OAProduct["hil_deliverydate"] = Convert.ToDateTime(item.DELIVERY_DATE);
                                        }
                                        //OAProduct["hil_deliverydate"] = Convert.ToDateTime(item.DELIVERY_DATE);
                                        OAProduct["hil_name"] = item.ZTENDER_LINE;
                                        OAProduct["hil_materialcategorydescription"] = Convert.ToString(item.CATEGORY_DISC);
                                        OAProduct["hil_materialgroup"] = item.MATKL;
                                        OAProduct["hil_netvalue"] = new Money(Convert.ToDecimal(item.VALUE));
                                        if (item.WERKS != null)
                                        {
                                            OAProduct["hil_plant"] = getPlantRef(item.WERKS, service);
                                        }
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

                                        //commented on 03/08/22 OAProduct["hil_quantity"] = Convert.ToDecimal(item.QTY) * 1000;


                                        //if (Department.ToLower() == "cable") //added below four line for qty department wise precondition
                                        //    OAProduct["hil_quantity"] = Convert.ToDecimal(item.QTY) * 1000;
                                        //else if (Department.ToLower() == "motor")
                                        //    OAProduct["hil_quantity"] = Convert.ToDecimal(item.QTY);

                                        //chnaged by Saurabh for Conversion Factor
                                        int conversionFactor = service.Retrieve(DepartmentRef.LogicalName, DepartmentRef.Id, new ColumnSet("hil_conversionfactor")).GetAttributeValue<int>("hil_conversionfactor");
                                        if (conversionFactor != 0)
                                            OAProduct["hil_quantity"] = Convert.ToDecimal(item.QTY) * conversionFactor;
                                        else
                                            OAProduct["hil_quantity"] = Convert.ToDecimal(item.QTY);
                                        //chnaged end Here by Saurabh for Conversion Factor

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
                        //sendEmail(service, bodtText, orderchecklist.ToEntityReference(), owner);
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
        public static EntityReference getPlantRef(int plantCode, IOrganizationService _service)
        {
            EntityReference panltRef = null;
            try
            {
                QueryExpression Query = new QueryExpression("hil_plantmaster");
                Query.ColumnSet = new ColumnSet(false);
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_plantname", ConditionOperator.Equal, plantCode.ToString());
                EntityCollection Found = _service.RetrieveMultiple(Query);
                if (Found.Entities.Count > 0)
                {
                    panltRef = Found.Entities[0].ToEntityReference();
                }
            }
            catch { }
            return panltRef;
        }
        public static string dateFormatter(DateTime date)
        {
            string _enquiryDatetime = string.Empty;
            if (date.Year.ToString().PadLeft(4, '0') != "0001")
                _enquiryDatetime = date.Year.ToString() + "-" + date.Month.ToString().PadLeft(2, '0') + "-" + date.Day.ToString().PadLeft(2, '0');
            return _enquiryDatetime;
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
        public static string genrateDeliveryDate(IOrganizationService _service, EntityReference product)
        {
            string dayss = null;
            try
            {

                Entity oclPrd = _service.Retrieve(product.LogicalName, product.Id, new ColumnSet("hil_orderchecklistid", "hil_plantcode"));
                EntityReference OCL = oclPrd.GetAttributeValue<EntityReference>("hil_orderchecklistid");
                EntityReference plant = oclPrd.GetAttributeValue<EntityReference>("hil_plantcode");
                Entity oclEntty = _service.Retrieve(OCL.LogicalName, OCL.Id, new ColumnSet("hil_typeoforder", "hil_department")); //added a colum "hil_department"
                int typeofOrder = oclEntty.GetAttributeValue<OptionSetValue>("hil_typeoforder").Value;
                EntityReference Department = oclEntty.GetAttributeValue<EntityReference>("hil_department");//added for get department

                QueryExpression Query = new QueryExpression("hil_deliveryschedulemaster");
                Query.ColumnSet = new ColumnSet("hil_deliveryday");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_ordertype", ConditionOperator.Equal, typeofOrder);
                Query.Criteria.AddCondition("hil_department", ConditionOperator.Equal, Department.Id);//added to get department wise deliveryschedulemaster 

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
                    _service.Update(_oclPrd);
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
                        _service.Update(_oclPrd);
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
