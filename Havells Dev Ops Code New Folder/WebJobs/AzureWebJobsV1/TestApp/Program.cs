using System;
using System.ServiceModel.Description;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System.Configuration;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Collections.Generic;
using System.Text;
using System.Security.Cryptography;
using System.IO;

namespace TestApp
{
    class Program
    {
        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        static string _securityPublicKey = string.Empty;
        #endregion
        static void Main(string[] args)
        {
            string txt = "C 13042320665738 187121";
            string[] sms = txt.Split(' ');
            long JobIdVer = 0;
            int iKKGVer = 0;
            //string KKG = txt.Substring(txt.Length - 4);
            //string JobId = txt.Substring(1, txt.Length - 5);
            string KKG = sms[2];
            string JobId = sms[1];
            bool result = long.TryParse(JobId, out JobIdVer);
            if (result)
            {
                result = int.TryParse(KKG, out iKKGVer);
                if (result)
                {
                    //UpdateWorkOrderSMS(service, SMS, KKG, JobId);
                }
            }

            //int _rowCount = 5001;
            //string _str = "HIL_" + _rowCount.ToString().PadLeft(8, '0');
            //Console.WriteLine(EncryptAES256("a4ead844-adbc-ed11-83ff-6045bdac5013"));
            //Console.WriteLine(EncryptAES256("8285906486"));
            //Console.WriteLine(EncryptAES256("5"));
            //Console.ReadLine();
            //string srno = "1231234_SVC";
            //string AlphaRegex = @"^[a-zA-Z0-9]*$";
            //if (!System.Text.RegularExpressions.Regex.IsMatch(srno.Replace("_SVC", ""), AlphaRegex))
            //    Console.WriteLine("Serial Number is not valid.");

            ///*Console.WriteLine(Reg*/ex.Replace("", @"[^0-9a-zA-Z]+", ""));
            //LoadAppSettings();
            //string _JobId = SecurityUtility.Encrypt("2002218395659", "ConSuMerNpS");
            if (loginUserGuid != Guid.Empty)
            {
                //SFA_ServiceCall _sfaservicecall = new SFA_ServiceCall();

                //SFA_ServiceCallResult _result = _sfaservicecall.SFA_CreateServiceCall(new SFA_ServiceCall()
                //{
                //    Address = "Address Line 1",
                //    AreaCode = "0VZ050",
                //    CustomerEmailId = "kuldeep.khare@havells.com",
                //    CustomerGuid = new Guid("ec0baec7-9967-ec11-8943-6045bda5eeb8"),
                //    CustomerMobileNo = "8285906486",
                //    CustomerName = "Kuldeep Khare",
                //    PINCode = "110048",
                //    Source = "7",
                //    ServiceCallData = new List<SFA_ServiceCallData>() { new SFA_ServiceCallData() 
                //    {
                //        Address="Address Line 1",
                //        AreaCode="0VZ050",
                //        CallType="I",
                //        PINCode="110048",
                //        ModelNumber="GOS24I02WOEL",
                //        Qty=1,
                //        ServiceDate="2022-03-30"
                //    }
                // }
                //}, _service);

                //AddressGuid = new Guid("a2b5c077-34af-ec11-9840-6045bdaa981b"),
                // AddressGuid = new Guid("a2b5c077-34af-ec11-9840-6045bdaa981b"),

                //SFA_ServiceCallResult _result = _sfaservicecall.SFA_CreateServiceCall(new SFA_ServiceCall()
                //{
                //    Address = "Address Line 1",
                //    AreaCode = "0VZ050",
                //    CustomerEmailId = "kuldeep.khare@havells.com",
                //    CustomerGuid = new Guid("ec0baec7-9967-ec11-8943-6045bda5eeb8"),
                //    CustomerMobileNo = "8285906486",
                //    CustomerName = "Kuldeep Khare",
                //    PINCode = "110048",
                //    Source = "7",
                //    ServiceCallData = new List<SFA_ServiceCallData>() { new SFA_ServiceCallData() {
                //        Address="Address Line 1",
                //        AreaCode="0VZ050",
                //        CallType="I",
                //        PINCode="110048",
                //        ModelNumber="GOS24I02WOEL",
                //        Qty=1
                //    }
                //}
                //}, _service);

                //msdyn_workorder _enJobTemp = (msdyn_workorder)_service.Retrieve(msdyn_workorder.EntityLogicalName, new Guid("8494a6bf-a076-ec11-8943-000d3af0ffdb"), new ColumnSet("hil_callsubtype"));
                ////AMC Call
                //if (_enJobTemp != null && _enJobTemp.GetAttributeValue<EntityReference>("hil_callsubtype").Id == new Guid("55a71a52-3c0b-e911-a94e-000d3af06cd4"))
                //{
                //    QueryExpression expJobProduct = new QueryExpression("msdyn_workorderproduct");
                //    expJobProduct.ColumnSet = new ColumnSet("msdyn_product");
                //    expJobProduct.Criteria = new FilterExpression(LogicalOperator.And);
                //    expJobProduct.Criteria.AddCondition("msdyn_workorder", ConditionOperator.Equal, new Guid("8494a6bf-a076-ec11-8943-000d3af0ffdb"));
                //    LinkEntity EntityA = new LinkEntity("msdyn_workorderproduct", "hil_productcatalog", "msdyn_product", "hil_productcode", JoinOperator.LeftOuter);
                //    EntityA.Columns = new ColumnSet("hil_amctandc");
                //    EntityA.EntityAlias = "Prod";
                //    expJobProduct.LinkEntities.Add(EntityA);
                //    LinkEntity EntityB = new LinkEntity("msdyn_workorderproduct", "product", "msdyn_product", "productid", JoinOperator.Inner);
                //    EntityB.Columns = new ColumnSet(false);
                //    EntityB.LinkCriteria = new FilterExpression(LogicalOperator.And);
                //    EntityB.LinkCriteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 910590001);
                //    EntityB.LinkCriteria.AddCondition("msdyn_fieldserviceproducttype", ConditionOperator.Equal, 690970001);
                //    EntityB.EntityAlias = "ProdAMC";
                //    expJobProduct.LinkEntities.Add(EntityB);
                //    EntityCollection entCollJobProd = _service.RetrieveMultiple(expJobProduct);
                //    if (entCollJobProd.Entities.Count == 0)
                //    {
                //        throw new Exception("ERROR !!! AMC Product is required.");
                //    }
                //    else
                //    {
                //        if (!entCollJobProd.Entities[0].Contains("Prod.hil_amctandc"))
                //        {
                //            throw new Exception("Warranty Desc is not defined in Master setup.");
                //        }
                //    }
                //}

                //string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                //                  <entity name='hil_tenderproduct'>
                //                    <attribute name='hil_tenderproductid' />
                //                    <attribute name='hil_product' />
                //                    <attribute name='hil_name' />
                //                    <attribute name='hil_lprsmtr' />
                //                    <order attribute='hil_name' descending='false' />
                //                    <filter type='and'>
                //                      <condition attribute='hil_tenderid' operator='eq' value='{3f483cab-4593-ec11-b400-6045bdaada1b}' />
                //                      <condition attribute='statecode' operator='eq' value='0' />
                //                    </filter>
                //                    <link-entity name='product' from='productid' to='hil_product' visible='false' link-type='outer' alias='prd'>
                //                      <attribute name='hil_amount' />
                //                    </link-entity>
                //                  </entity>
                //                </fetch>";
                //EntityCollection _entColl = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                //if (_entColl.Entities.Count > 0)
                //{
                //    Decimal _tndProdLP;
                //    Decimal _prodLP;
                //    string _rowIds = "LP has been changed for Product# ";
                //    foreach (Entity ent in _entColl.Entities)
                //    {
                //        _tndProdLP = 0;
                //        _prodLP = 0;
                //        _tndProdLP = ent.GetAttributeValue<Money>("hil_lprsmtr").Value;
                //        if (ent.Contains("prd.hil_amount"))
                //        {
                //            _prodLP = Math.Round(((Money)ent.GetAttributeValue<AliasedValue>("prd.hil_amount").Value).Value / 1000, 2);
                //        }
                //        if (_tndProdLP != _prodLP)
                //        {
                //            _rowIds = _rowIds + ent.GetAttributeValue<EntityReference>("hil_product").Name + ",";
                //        }
                //    }
                //}

                //RequestDTO _data = new RequestDTO() { name = "api_send_email", parameters = new ParametersDTO() { message_body = "Dear KHOMLAL VERMA, Your Product Service with ref no 27122112450612 has been assigned to VIPUL KATKIYA his mob no 7990720616. He will fix appointment before visiting - Havells", mobile_number = "8285906486", subject = "Technician Number for the Service Request",templateId= "1107162314614376909" } };
                //CommunicationNPS _obj = new CommunicationNPS();
                //_obj.SendCommunication(_data, _service);

                //Entity ent = _service.Retrieve("hil_tenderproduct", new Guid("0a03b494-5107-ec11-b6e7-000d3af0b5fa"), new ColumnSet("hil_lprsmtr", "hil_hopricespecialconstructions"));
                //if (ent != null && ent.Contains("hil_lprsmtr"))
                //{
                //    Decimal _lpMtr = ent.GetAttributeValue<Money>("hil_lprsmtr").Value;
                //    Decimal _hoPrice = 1140 ;//ent.GetAttributeValue<Money>("hil_hopricespecialconstructions").Value;
                //    Decimal hodiscount = ((1 - Math.Round(_hoPrice, 2) / Math.Round(_lpMtr, 2)) * 100);
                //    Entity entUpdate = new Entity("hil_tenderproduct", ent.Id);
                //    entUpdate["hil_approveddiscount"] = Math.Round(hodiscount,2);
                //    entUpdate["hil_hodiscper"] = Math.Round(hodiscount, 2);
                //    _service.Update(entUpdate);
                //}

                //ent = _service.Retrieve("hil_tenderproduct", new Guid("0a03b494-5107-ec11-b6e7-000d3af0b5fa"), new ColumnSet("hil_hodiscper","hil_lprsmtr", "hil_hopricespecialconstructions"));
                //if (ent != null && ent.Contains("hil_hodiscper"))
                //{
                //    Decimal _hoDiscPer = Math.Round(ent.GetAttributeValue<Decimal>("hil_hodiscper"), 2);
                //    Entity entUpdate = new Entity("hil_tenderproduct", ent.Id);
                //    entUpdate["hil_approveddiscount"] = _hoDiscPer;

                //    Entity ent1 = _service.Retrieve("hil_tenderproduct", ent.Id, new ColumnSet("hil_lprsmtr"));
                //    if (ent1 != null && ent1.Contains("hil_lprsmtr"))
                //    {
                //        Decimal _lpMtr = ent1.GetAttributeValue<Money>("hil_lprsmtr").Value;
                //        Decimal hoPrice = (_lpMtr * (1 - _hoDiscPer / 100));
                //        entUpdate["hil_hopricespecialconstructions"] = new Money(hoPrice);
                //    }
                //    _service.Update(entUpdate);
                //}

                //HomeAdvisory HA = new HomeAdvisory();
                ////GetEnquiry _retObj = HA.GetEnqueryStatus(new EnquiryStatus() { MobileNo = "9818011012", FromDate = "2021-08-19", ToDate = "2021-08-19" }, _service);
                //List<EnquiryDetails> _retObj = HA.GetSalesEnquiry(_service, new EnquiryStatus() { MobileNo = "9818021012", FromDate = "2021-08-01", ToDate = "2021-08-26" });

                //LogicCheck();
                //D365AuditLog objLog = new D365AuditLog();
                //List<D365AuditLogResult> _obj = objLog.GetD365AuditLogData(new D365AuditLog()
                //{
                //    EntityName = "msdyn_workorder",
                //    FieldName = "All",
                //    RecId = "7b243e1b-04f5-eb11-94ef-000d3a3e3f77"
                //}, _service);

                //LogicCheck();
                //D365AuditLog _obj = new D365AuditLog();
                //List<D365AuditLogResult> _retOnj = _obj.GetD365AuditLogData(new D365AuditLog() {
                //EntityName= "msdyn_workorder",
                //FieldName= "Sub-Status",
                //RecId= "d2511689-d5ea-eb11-bacb-000d3af0cd16",
                //FromDate= "20210723",
                //ToDate= "20210803"
                //});

                //try
                //{
                //    List<EnquiryType> _retObf = GetEnquiryTypes("1", _service);
                //    List<EnquiryProductType> _retObf1 = GetEnquiryProductTypes("2", _service);
                //}
                //catch (Exception ex)
                //{
                //    //return new CommonLib() { status = "204", statusRemarks = ex.Message };
                //}

                //FlushUATData _retObj = new FlushUATData();
                //_retObj.FlushData(new FlushUATData() {MobileNumber= "9589896282", EntryType="" }, _service);

            }
        }

        static string EncryptAES256(string plainText)
        {
            string Key = "DklsdvkfsDlkslsdsdnv234djSDAjkd1";
            byte[] key32 = Encoding.UTF8.GetBytes(Key);
            byte[] IV16 = Encoding.UTF8.GetBytes(Key.Substring(0, 16)); if (plainText == null || plainText.Length <= 0)
            {
                throw new ArgumentNullException("plainText");
            }
            byte[] encrypted;
            using (AesManaged aesAlg = new AesManaged())
            {
                aesAlg.KeySize = 256;
                aesAlg.Padding = PaddingMode.PKCS7;
                aesAlg.IV = IV16;
                aesAlg.Key = key32;
                ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
                using (MemoryStream msEncrypt = new MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                        {
                            //Write all data to the stream.
                            swEncrypt.Write(plainText);
                        }
                        encrypted = msEncrypt.ToArray();
                    }
                }
            }
            return Convert.ToBase64String(encrypted);
        }
        static void CustomerAssetImport(IOrganizationService _service) { 
        
        
        }
        static void LogicCheck() {
            #region RepeatRepair Approval
            Guid _entityId = new Guid("89A2847A-C6F2-EB11-94EF-000D3A3E3D78");
            msdyn_workorder _enJob1 = (msdyn_workorder)_service.Retrieve(msdyn_workorder.EntityLogicalName, _entityId, new ColumnSet("msdyn_customerasset","msdyn_timeclosed","hil_typeofassignee", "hil_isgascharged", "hil_laborinwarranty", "hil_callsubtype", "hil_productsubcategory", "hil_customerref", "createdon", "hil_isocr", "hil_claimstatus"));

            DateTime _createdOn = _enJob1.GetAttributeValue<DateTime>("createdon").AddDays(-15);
            DateTime _ClosedOn = _enJob1.GetAttributeValue<DateTime>("msdyn_timeclosed").AddDays(-15); // DateTime.Now.AddDays(-15);
            string _strCreatedOn = _createdOn.Year.ToString() + "-" + _createdOn.Month.ToString() + "-" + _createdOn.Day.ToString();
            string _strClosedOn = _ClosedOn.Year.ToString() + "-" + _ClosedOn.Month.ToString() + "-" + _ClosedOn.Day.ToString();
            EntityCollection entCol;

            //Callsubtype{8D80346B-3C0B-E911-A94E-000D3AF06CD} Dealer Stock Repair
            //JobStatus: {1727FA6C-FA0F-E911-A94E-000D3AF060A1}-Closed,2927FA6C-FA0F-E911-A94E-000D3AF060A1-Workdone,7E85074C-9C54-E911-A951-000D3AF0677F-workdone SMS
            string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='msdyn_workorder'>
                <attribute name='msdyn_name' />
                <attribute name='createdon' />
                <attribute name='hil_productsubcategory' />
                <attribute name='hil_customerref' />
                <attribute name='hil_callsubtype' />
                <attribute name='msdyn_customerasset' />
                <attribute name='msdyn_workorderid' />
                <attribute name='msdyn_timeclosed' />
                <attribute name='msdyn_closedby' />
                <order attribute='createdon' descending='true' />
                <filter type='and'>
                    <condition attribute='hil_isocr' operator='ne' value='1' />
                    <condition attribute='hil_typeofassignee' operator='ne' value='{7D1ECBAB-1208-E911-A94D-000D3AF0694E}' />
                    <condition attribute='msdyn_workorderid' operator='ne' value='" + _entityId + @"' />
                    <condition attribute='hil_customerref' operator='eq' value='" + _enJob1.GetAttributeValue<EntityReference>("hil_customerref").Id + @"' />
                    <condition attribute='hil_callsubtype' operator='eq' value='" + _enJob1.GetAttributeValue<EntityReference>("hil_callsubtype").Id + @"' />
                    <condition attribute='hil_callsubtype' operator='ne' value='{8D80346B-3C0B-E911-A94E-000D3AF06CD4}' />
                    <condition attribute='hil_productsubcategory' operator='eq' value='" + _enJob1.GetAttributeValue<EntityReference>("hil_productsubcategory").Id + @"' />
                    <filter type='or'>
                        <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strCreatedOn + @"' />
                        <condition attribute='msdyn_timeclosed' operator='on-or-after' value='" + _strClosedOn + @"' />
                    </filter>
                    <condition attribute='msdyn_substatus' operator='in'>
                    <value>{1727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                    <value>{2927FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                    <value>{7E85074C-9C54-E911-A951-000D3AF0677F}</value>
                    </condition>
                </filter>
                </entity>
                </fetch>";
            entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (entCol.Entities.Count > 0)
            {
                string _remarks = string.Empty;
                string _remarks1 = string.Empty;
                EntityReference _entref = null;
                foreach (Entity ent in entCol.Entities) {
                    if (ent.GetAttributeValue<EntityReference>("msdyn_customerasset").Id == _enJob1.GetAttributeValue<EntityReference>("msdyn_customerasset").Id)
                    {
                        _remarks += ent.GetAttributeValue<string>("msdyn_name") + ",";
                        _entref = ent.ToEntityReference();
                    }
                    else
                    {
                        _remarks1 += ent.GetAttributeValue<string>("msdyn_name") + ",";
                    }
                }
                _remarks = ((_remarks == string.Empty ? "" : "Repeated Jobs with Same Serial Number: " + _remarks + ":\n") + (_remarks1 == string.Empty ? "" : "Repeated Jobs with Same Product Subcategory: " + _remarks1 + ":")).Replace(",:", "");
                CommonLib obj = new CommonLib();
                CommonLib objReturn = obj.CreateSAWActivity(_enJob1.Id, 0, SAWCategoryConst._RepeatRepair, _service, _remarks, entCol.Entities[0].ToEntityReference());
            }
            #endregion
        }
        static bool CheckIfSAWActivityExist(Guid _jobId, IOrganizationService _service, string _sawCategory)
        {
            try
            {
                QueryExpression qryExp = new QueryExpression("hil_sawactivity");
                qryExp.ColumnSet = new ColumnSet("hil_jobid");
                qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                qryExp.Criteria.AddCondition("hil_jobid", ConditionOperator.Equal, _jobId);
                qryExp.Criteria.AddCondition("hil_sawcategory", ConditionOperator.Equal, new Guid(_sawCategory));
                EntityCollection entCol = _service.RetrieveMultiple(qryExp);
                if (entCol.Entities.Count > 0) { return true; } else { return false; }
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        static void FillMyTeam()
        {
            Guid _ownerId = new Guid("84392f95-ba96-eb11-b1ac-0022486ec132");
            Entity entOwner = _service.Retrieve("systemuser", _ownerId, new ColumnSet("businessunitid", "positionid"));
            if (entOwner.Contains("businessunitid") && entOwner.Contains("positionid"))
            {
                string _positionName = entOwner.GetAttributeValue<EntityReference>("positionid").Name.ToUpper();
                if (_positionName == "FRANCHISE")
                {

                }
                else if (_positionName == "CCO" || _positionName == "BSH")
                {

                }
            }
        }
        static void CreateAdvisoryLine()
        {
            Guid _advisoryHeader = new Guid("84392f95-ba96-eb11-b1ac-0022486ec132");
            Entity entAdvisoryHeader = _service.Retrieve("hil_advisoryenquiry", _advisoryHeader, new ColumnSet("hil_sourceofcreation", "hil_typeofproduct", "hil_typeofenquiry", "hil_pincode"));
            if (entAdvisoryHeader.Contains("hil_typeofproduct"))
            {
                OptionSetValueCollection _prodTypes = entAdvisoryHeader.GetAttributeValue<OptionSetValueCollection>("hil_typeofproduct");
                string _prodTypesName = string.Empty;
                foreach (OptionSetValue obj in _prodTypes)
                {
                    string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_typeofproduct'>
                            <attribute name='hil_name' />
                            <filter type='and'>
                                <condition attribute='hil_index' operator='eq' value='" + obj.Value.ToString() + @"' />
                            </filter>
                            </entity>
                            </fetch>";
                    EntityCollection entColProdType = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entColProdType.Entities.Count > 0)
                    {
                        _prodTypesName += entColProdType.Entities[0].GetAttributeValue<string>("hil_name") + ",";
                    }
                }
                if (_prodTypesName.Length > 0)
                {
                    _prodTypesName = _prodTypesName.Substring(0, _prodTypesName.Length - 1);
                }
                entAdvisoryHeader["hil_productname"] = _prodTypesName;
            }
        }
        #region App Setting Load/CRM Connection
        static void LoadAppSettings()
        {
            try
            {
                _userId = ConfigurationManager.AppSettings["CrmUserId"].ToString();
                _password = ConfigurationManager.AppSettings["CrmUserPassword"].ToString();
                _soapOrganizationServiceUri = ConfigurationManager.AppSettings["CrmSoapOrganizationServiceUri"].ToString();
                _securityPublicKey = ConfigurationManager.AppSettings["SecurityPublicKey"].ToString();
                ConnectToCRM();
            }
            catch (Exception ex)
            {
                Console.WriteLine("D365AuditLogMigration.Program.Main.LoadAppSettings ::  Error While Loading App Settings:" + ex.Message.ToString());
            }
        }
        static void ConnectToCRM()
        {
            try
            {
                ClientCredentials credentials = new ClientCredentials();
                credentials.UserName.UserName = _userId;
                credentials.UserName.Password = _password;
                Uri serviceUri = new Uri(_soapOrganizationServiceUri);
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                OrganizationServiceProxy proxy = new OrganizationServiceProxy(serviceUri, null, credentials, null);

                proxy.EnableProxyTypes();
                _service = (IOrganizationService)proxy;
                loginUserGuid = ((WhoAmIResponse)_service.Execute(new WhoAmIRequest())).UserId;
            }
            catch (Exception ex)
            {
                Console.WriteLine("D365AuditLogMigration.Program.Main.ConnectToCRM :: Error While Creating Connection with MS CRM Organisation:" + ex.Message.ToString());
            }
        }
        #endregion

        static List<EnquiryType> GetEnquiryTypes(string _enquiryCategory, IOrganizationService service)
        {
            List<EnquiryType> lstEnquiryType = new List<EnquiryType>();
            try
            {
                //IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    string _fetchxml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='hil_enquirytype'>
                            <attribute name='hil_enquirytypename' />
                            <attribute name='hil_enquirytypecode' />
                            <order attribute='hil_enquirytypename' descending='false' />
                            <filter type='and'>
                              <condition attribute='statecode' operator='eq' value='0' />
                              <condition attribute='hil_enquirycategory' operator='eq' value='" + _enquiryCategory + @"' />
                            </filter>
                          </entity>
                        </fetch>";
                    EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchxml));
                    if (entCol.Entities.Count > 0)
                    {
                        foreach (Entity ent in entCol.Entities)
                        {
                            lstEnquiryType.Add(new EnquiryType()
                            {
                                EnquiryTypeCode = ent.GetAttributeValue<string>("hil_enquirytypename"),
                                EnquiryTypeName = ent.GetAttributeValue<Int32>("hil_enquirytypecode").ToString(),
                            });
                        }
                    }
                    else
                    {
                        lstEnquiryType.Add(new EnquiryType() { EnquiryTypeCode = "ERROR", EnquiryTypeName = "No Data found." });
                    }
                    return lstEnquiryType;
                }
                else
                {
                    lstEnquiryType.Add(new EnquiryType() { EnquiryTypeCode = "ERROR", EnquiryTypeName = "D365 Service Unavailable" });
                    return lstEnquiryType;
                }
            }
            catch (Exception ex)
            {
                lstEnquiryType.Add(new EnquiryType() { EnquiryTypeCode = "ERROR", EnquiryTypeName = "D365 Internal Server Error : " + ex.Message });
                return lstEnquiryType;
            }
        }

        static List<EnquiryProductType> GetEnquiryProductTypes(string _enquiryCategory, IOrganizationService service)
        {
            List<EnquiryProductType> lstDivisionCallType = new List<EnquiryProductType>();
            try
            {
                //IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    string _fetchxml = string.Empty;

                    if (_enquiryCategory == "1")
                    {
                        _fetchxml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='hil_typeofproduct'>
                            <attribute name='hil_typeofproductid' />
                            <attribute name='hil_name' />
                            <attribute name='hil_typeofenquiry' />
                            <attribute name='hil_index' />
                            <attribute name='hil_code' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                              <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            <link-entity name='hil_enquirytype' from='hil_enquirytypeid' to='hil_typeofenquiry' visible='false' link-type='outer' alias='eq'>
                              <attribute name='hil_enquirytypename' />
                              <attribute name='hil_enquirytypecode' />
                            </link-entity>
                          </entity>
                        </fetch>";
                        EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchxml));
                        if (entCol.Entities.Count > 0)
                        {
                            foreach (Entity ent in entCol.Entities)
                            {
                                lstDivisionCallType.Add(new EnquiryProductType()
                                {
                                    EnquiryProductCode = ent.GetAttributeValue<string>("hil_code"),
                                    EnquiryProductName = ent.GetAttributeValue<string>("hil_name"),
                                    EnquiryTypeCode = ent.GetAttributeValue<AliasedValue>("eq.hil_enquirytypename").Value.ToString(),
                                    EnquiryTypeName = ent.GetAttributeValue<AliasedValue>("eq.hil_enquirytypecode").Value.ToString(),
                                });
                            }
                        }
                        else
                        {
                            lstDivisionCallType.Add(new EnquiryProductType() { EnquiryProductCode = "ERROR", EnquiryProductName = "No Data found." });
                        }
                    }
                    else if (_enquiryCategory == "2")
                    {
                        _fetchxml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='product'>
                            <attribute name='name' />
                            <attribute name='productnumber' />
                            <attribute name='hil_sapcode' />
                            <order attribute='name' descending='false' />
                            <filter type='and'>
                              <condition attribute='hil_hierarchylevel' operator='eq' value='2' />
                              <condition attribute='producttypecode' operator='eq' value='1' />
                            </filter>
                          </entity>
                        </fetch>";
                        EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchxml));
                        if (entCol.Entities.Count > 0)
                        {
                            foreach (Entity ent in entCol.Entities)
                            {
                                lstDivisionCallType.Add(new EnquiryProductType()
                                {
                                    EnquiryProductCode = ent.GetAttributeValue<string>("hil_sapcode"),
                                    EnquiryProductName = ent.GetAttributeValue<string>("name"),
                                    EnquiryTypeCode = "",
                                    EnquiryTypeName = "",
                                });
                            }
                        }
                        else
                        {
                            lstDivisionCallType.Add(new EnquiryProductType() { EnquiryProductCode = "ERROR", EnquiryProductName = "No Data found." });
                        }

                    }
                    return lstDivisionCallType;
                }
                else
                {
                    lstDivisionCallType.Add(new EnquiryProductType() { EnquiryProductCode = "ERROR", EnquiryProductName = "D365 Service Unavailable" });
                    return lstDivisionCallType;
                }
            }
            catch (Exception ex)
            {
                lstDivisionCallType.Add(new EnquiryProductType() { EnquiryProductCode = "ERROR", EnquiryProductName = "D365 Internal Server Error : " + ex.Message });
                return lstDivisionCallType;
            }
        }
    }
    public class EnquiryType
    {
        public string EnquiryTypeName { get; set; }
        public string EnquiryTypeCode { get; set; }
    }

    public class EnquiryProductType
    {
        public string EnquiryTypeName { get; set; }
        public string EnquiryTypeCode { get; set; }
        public string EnquiryProductCode { get; set; }
        public string EnquiryProductName { get; set; }

    }
}

