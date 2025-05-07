using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Text;
using System.Security.Cryptography;
using Havells_Plugin.SAWActivity;

namespace Havells_Plugin.WorkOrder
{
    public  class Common
    {
        public static void SetFranchisee(Entity entity, IOrganizationService service)
        {
            try
            {
                msdyn_workorder WorkOrder = entity.ToEntity<msdyn_workorder>();
                if (WorkOrder.OwnerId.Id != null)
                {
                    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                    {
                        WorkOrder =(msdyn_workorder)service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_owneraccount"));
                        if (WorkOrder.hil_OwnerAccount != null)
                        {
                            var obj =from _Account in orgContext.CreateQuery<Account>()
                                     where _Account.AccountId.Value==WorkOrder.hil_OwnerAccount.Id
                                      select new {
                                          _Account.CustomerTypeCode
                                      };
                            foreach (var iobj in obj)
                            {

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrder.PreCreate.SetFranchisee: " + ex.Message);
            }
        }

        public static string Base64Encode(string plainText)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return System.Convert.ToBase64String(plainTextBytes);
        }

        public static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }

        public static string GetEncryptedValue(string data, IOrganizationService service)
        {
            string retValue = string.Empty;
            string key = string.Empty;
            try
            {
                QueryExpression queryExp = new QueryExpression("hil_kkgcodeconfiguration");
                queryExp.ColumnSet = new ColumnSet("hil_name");
                queryExp.Criteria = new FilterExpression(LogicalOperator.And);
                queryExp.Criteria.AddCondition(new ConditionExpression("createdon", ConditionOperator.Today));
                EntityCollection entCol = service.RetrieveMultiple(queryExp);
                if (entCol.Entities.Count > 0)
                {
                    key = entCol.Entities[0].GetAttributeValue<string>("hil_name");
                }
                else
                {
                    #region Generating 8 Digits Private Key
                    Random rnd = new Random();
                    string tempKey;
                    tempKey = rnd.Next(10000000, 99999999).ToString();
                    #endregion
                    #region Stamping Private Key for Future use for the same day
                    Entity ent = new Entity("hil_kkgcodeconfiguration");
                    ent["hil_name"] = key;
                    Guid privateKeyId = service.Create(ent);
                    if (privateKeyId != Guid.Empty)
                    {
                        key = tempKey;
                    }
                    #endregion
                }
                if (key != string.Empty)
                {
                    retValue = AESEncryption(data, key);
                }
            }
            catch {
                return null;
            }
            return retValue;
        }

        public static string AESEncryption(string textData, string Encryptionkey)

        {

            RijndaelManaged objrij = new RijndaelManaged();

            //set the mode for operation of the algorithm

            objrij.Mode = CipherMode.CBC;

            //set the padding mode used in the algorithm.

            objrij.Padding = PaddingMode.PKCS7;

            //set the size, in bits, for the secret key.

            objrij.KeySize = 0x80;

            //set the block size in bits for the cryptographic operation.

            objrij.BlockSize = 0x80;

            //set the symmetric key that is used for encryption & decryption.

            byte[] passBytes = Encoding.UTF8.GetBytes(Encryptionkey);

            //set the initialization vector (IV) for the symmetric algorithm

            byte[] EncryptionkeyBytes = new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };

            int len = passBytes.Length;

            if (len > EncryptionkeyBytes.Length)

            {

                len = EncryptionkeyBytes.Length;

            }

            Array.Copy(passBytes, EncryptionkeyBytes, len);

            objrij.Key = EncryptionkeyBytes;

            objrij.IV = EncryptionkeyBytes;

            //Creates symmetric AES object with the current key and initialization vector IV.

            ICryptoTransform objtransform = objrij.CreateEncryptor();

            byte[] textDataByte = Encoding.UTF8.GetBytes(textData);

            //Final transform the test string.

            return Convert.ToBase64String(objtransform.TransformFinalBlock(textDataByte, 0, textDataByte.Length));

        }

        #region Added by Kuldeep Khare on 25/Nov/2019 to get TAT Category by TAT Hrs
        public static EntityReference TATCategory(IOrganizationService service, double TAThr)
        {
            EntityReference erTemp = null;
            try
            {
                QueryExpression Query = new QueryExpression("hil_jobtatcategory");
                Query.ColumnSet = new ColumnSet("hil_jobtatcategoryid");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("hil_tatcategoryrangetohr", ConditionOperator.GreaterEqual, Convert.ToDecimal(TAThr));
                Query.Criteria.AddCondition("hil_tatcategoryrangefromhr", ConditionOperator.LessThan, Convert.ToDecimal(TAThr));
                EntityCollection entCol = service.RetrieveMultiple(Query);
                if (entCol.Entities.Count > 0)
                {
                    erTemp = entCol.Entities[0].ToEntityReference();
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WorkOrder.PostUpdate.Execute.TATCategory: " + ex.Message);
            }
            return erTemp;
        }
        #endregion

        public const string AMCCallSubTypeGUID = "55A71A52-3C0B-E911-A94E-000D3AF06CD4";

        #region Added by Saurabh Tripathi on 13/Jan/2021 Validate AMC Receipt Amount
        public static void ValidateAMCReceiptAmount(IOrganizationService service, Guid jobGuid, Decimal ReceiptAmount)
        {
            string _fetchXML = string.Empty;
            DateTime _invoiceDate;
            _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='msdyn_workorderincident'>
                        <attribute name='msdyn_name' />
                        <filter type='and'>
                          <condition attribute='msdyn_workorder' operator='eq' value='" + jobGuid + @"' />
                          <condition attribute='statecode' operator='eq' value='0' />
                        </filter>
                        <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='msdyn_customerasset' visible='false' link-type='outer' alias='ca'>
                          <attribute name='hil_invoicedate' />
                        </link-entity>
                        <link-entity name='msdyn_workorder' from='msdyn_workorderid' to='msdyn_workorder' visible='false' link-type='outer' alias='wo'>
                          <attribute name='createdon' />
                          <attribute name='hil_actualcharges' />
                          <attribute name='hil_callsubtype' />
                          <attribute name='hil_productcategory' />
                        </link-entity>
                      </entity>
                    </fetch>";
            EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (entCol.Entities.Count > 0)
            {
                if (!entCol.Entities[0].Attributes.Contains("ca.hil_invoicedate"))
                {
                    _invoiceDate = new DateTime(1900, 1, 1);
                }
                else
                {
                    _invoiceDate = (DateTime)(entCol.Entities[0].GetAttributeValue<AliasedValue>("ca.hil_invoicedate").Value);
                }
                EntityReference entTemp = (EntityReference)entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.hil_callsubtype").Value;
                EntityReference entProdCatg = (EntityReference)entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.hil_productcategory").Value;

                if (entTemp.Id == new Guid(JobCallSubType.AMCCall)) //AMC Call SubType
                {
                    decimal _payableAmount = ((Money)(entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.hil_actualcharges").Value)).Value;
                    DateTime _jobDate = (DateTime)entCol.Entities[0].GetAttributeValue<AliasedValue>("wo.createdon").Value;

                    //_asOn Definition :: AMC Job Create date is concidered for Applying Discount rate becoz Product ageing also calculated from AMC Job Create Date
                    string _asOn = _jobDate.Year.ToString() + "-" + _jobDate.Month.ToString().PadLeft(2, '0') + "-" + _jobDate.Day.ToString().PadLeft(2, '0');
                    int _dayDiff = Convert.ToInt32(Math.Round((_jobDate - _invoiceDate).TotalDays, 0));
                    decimal _stdDiscPer = 0;
                    decimal _spcDiscPer = 0;
                    decimal _stdDiscAmount = 0;
                    decimal _spcDiscAmount = 0;

                    _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_amcdiscountmatrix'>
                    <attribute name='hil_amcdiscountmatrixid' />
                    <attribute name='hil_name' />
                    <attribute name='createdon' />
                    <attribute name='hil_discounttype' />
                    <attribute name='hil_discper' />
                    <order attribute='hil_name' descending='false' />
                    <filter type='and'>
                        <condition attribute='statecode' operator='eq' value='0' />
                        <condition attribute='hil_appliedto' operator='eq' value='{03B5A2D6-CC64-ED11-9562-6045BDAC526A}' />
                        <condition attribute='hil_productaegingstart' operator='le' value='" + _dayDiff.ToString() + @"' />
                        <condition attribute='hil_productageingend' operator='ge' value='" + _dayDiff.ToString() + @"' />
                        <condition attribute='hil_validfrom' operator='on-or-before' value='" + _asOn + @"' />
                        <condition attribute='hil_validto' operator='on-or-after' value='" + _asOn + @"' />
                        <condition attribute='hil_productcategory' operator='eq' value='{" + entProdCatg.Id + @"}' />
                    </filter>
                    </entity>
                    </fetch>";
                    EntityCollection entCol1 = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entCol1.Entities.Count > 0)
                    {
                        foreach (Entity ent in entCol1.Entities)
                        {
                            if (ent.GetAttributeValue<OptionSetValue>("hil_discounttype").Value == 1)
                            {
                                _stdDiscPer = ent.GetAttributeValue<decimal>("hil_discper");
                            }
                            else
                            {
                                _spcDiscPer = ent.GetAttributeValue<decimal>("hil_discper");
                            }
                        }
                        _stdDiscAmount = Math.Round((_payableAmount - (_payableAmount * _stdDiscPer) / 100), 2); //Max Limit (90)
                        _spcDiscAmount = Math.Round(_payableAmount - (_payableAmount * (_stdDiscPer + _spcDiscPer)) / 100, 2); //Min Limit (85)
                        decimal _additionaldisc = Math.Round(_stdDiscAmount - ReceiptAmount, 2);                                                                                                       //ReceiptAmount 80
                        if (ReceiptAmount >= _stdDiscAmount) // ReceiptAmount = 90 
                        {
                            //no need to do any thing
                        }
                        if (ReceiptAmount < _stdDiscAmount && ReceiptAmount >= _spcDiscAmount) // ReceiptAmount = 89
                        {
                            //Create Approval
                            StringBuilder strRemarks = new StringBuilder();
                            strRemarks.AppendLine("AMC Price : " + Math.Round(_payableAmount,2).ToString());
                            strRemarks.AppendLine("After Standard Discount : " + _stdDiscAmount.ToString());
                            strRemarks.AppendLine("Price offered : " + Math.Round(ReceiptAmount,2).ToString());
                            strRemarks.AppendLine("Approval difference = " + _additionaldisc.ToString());

                            CommonLib obj = new CommonLib();
                            CommonLib objReturn = obj.CreateSAWActivity(jobGuid, _additionaldisc, SAWCategoryConst._AMCSpecialDiscount, service, strRemarks.ToString(), null);
                        }
                        if (ReceiptAmount < _spcDiscAmount)// ReceiptAmount = 84
                        {
                            throw new InvalidPluginExecutionException("As Per AMC Discount Policy minimum allowed receipt amount is " + _spcDiscAmount.ToString());
                        }
                    }
                }
            }
            else
            {
                throw new InvalidPluginExecutionException("No Work Order Incident found.");
            }
        }
        #endregion
        public static void PreJobClosureValidations(IOrganizationService service, Guid jobGuid)
        {
            string _fetchXML = string.Empty;
            _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='msdyn_workorder'>
                <attribute name='hil_callsubtype' />
                <attribute name='hil_salesoffice' />
                <attribute name='hil_owneraccount' />
                <attribute name='msdyn_customerasset' />
                <attribute name='hil_productcategory' />
                <filter type='and'>
                  <filter type='or'>
                    <condition attribute='hil_callsubtype' operator='null' />
                    <condition attribute='hil_salesoffice' operator='null' />
                    <condition attribute='hil_owneraccount' operator='null' />
                    <condition attribute='msdyn_customerasset' operator='null' />
                    <condition attribute='hil_productcategory' operator='null' />
                  </filter>
                  <condition attribute='msdyn_workorderid' operator='eq' value='{" + jobGuid + @"}' />
                </filter>
              </entity>
            </fetch>";
            EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
            string _errorMsg = string.Empty;

            if (entCol.Entities.Count > 0)
            {
                if (!entCol.Entities[0].Attributes.Contains("hil_callsubtype"))
                {
                    _errorMsg = _errorMsg + " ,Call SubType";
                }
                if (!entCol.Entities[0].Attributes.Contains("hil_salesoffice"))
                {
                    _errorMsg = _errorMsg + " ,Sales Office";
                }
                if (!entCol.Entities[0].Attributes.Contains("hil_owneraccount"))
                {
                    _errorMsg = _errorMsg + " ,Franchisee";
                }
                if (!entCol.Entities[0].Attributes.Contains("msdyn_customerasset"))
                {
                    _errorMsg = _errorMsg + " ,Customer Asset";
                }
                if (!entCol.Entities[0].Attributes.Contains("hil_productcategory"))
                {
                    _errorMsg = _errorMsg + " ,Product Category";
                }
                _errorMsg = _errorMsg + " are required.";
                throw new InvalidPluginExecutionException(" ***" + _errorMsg + "*** ");
            }
            else
            {
                Guid _callsubType = service.Retrieve("msdyn_workorder", jobGuid, new ColumnSet("hil_callsubtype")).GetAttributeValue<EntityReference>("hil_callsubtype").Id;
                if (_callsubType == new Guid(JobCallSubType.AMCCall))
                {
                    _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='msdyn_workorderproduct'>
                        <attribute name='createdon' />
                        <filter type='and'>
                          <condition attribute='hil_markused' operator='eq' value='1' />
                          <condition attribute='statecode' operator='eq' value='0' />
                          <condition attribute='msdyn_workorder' operator='eq' value='{" + jobGuid + @"}' />
                        </filter>
                        <link-entity name='product' from='productid' to='hil_replacedpart' link-type='inner' alias='ac'>
                          <filter type='and'>
                            <condition attribute='hil_hierarchylevel' operator='eq' value='910590001' />
                          </filter>
                        </link-entity>
                      </entity>
                    </fetch>";
                    entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entCol.Entities.Count == 0)
                    {
                        throw new InvalidPluginExecutionException(" ***Atleast one AMC Product is required to close this Job.*** ");
                    }
                }
            }
        }

        public static bool ValidateOnlinePayment(IOrganizationService service, Guid jobGuid)
        {
            bool _retValue = false;

            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='hil_paymentstatus'>
                        <attribute name='hil_paymentstatusid' />
                        <attribute name='hil_paymentstatus' />
                        <filter type='and'>
                          <condition attribute='hil_job' operator='eq' value='{jobGuid.ToString()}' />
                        </filter>
                      </entity>
                    </fetch>";
            EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (_entCol.Entities.Count > 0)
            {
                string _status = _entCol.Entities[0].GetAttributeValue<string>("hil_paymentstatus");
                if (_status == "success")
                {
                    _retValue = true;
                }
                else
                {
                    OrganizationRequest req = new OrganizationRequest("hil_NewGetPaymentStatusAction");
                    req["EntityID"] = jobGuid.ToString().Replace("{", "").Replace("}", "");
                    req["EntityName"] = "msdyn_workorder";
                    OrganizationResponse response = service.Execute(req);
                    string Status = response.Results["Status"].ToString();
                    string Message = response.Results["Message"].ToString();
                    if (Message == "Payment received successfully.")
                    {
                        _retValue = true;
                    }
                }
            }
            else
            {
                throw new InvalidPluginExecutionException("Payment for this AMC is pending.");
            }
            return _retValue;
        }
    }
    public static class JobCallSubType {
        public const string PMS = "E2129D79-3C0B-E911-A94E-000D3AF06CD4";
        public const string Breakdown = "6560565A-3C0B-E911-A94E-000D3AF06CD4";
        public const string DealerStockRepair = "8D80346B-3C0B-E911-A94E-000D3AF06CD4";
        public const string AMCCall = "55a71a52-3c0b-e911-a94e-000d3af06cd4";
        public const string Installation = "e3129d79-3c0b-e911-a94e-000d3af06cd4";
        public const string PDI = "ce45f586-3c0b-e911-a94e-000d3af06cd4";
    }
    public static class JobProductCategory
    {
        public const string LLOYDAIRCONDITIONER = "D51EDD9D-16FA-E811-A94C-000D3AF0694E";
        public const string LLOYDREFRIGERATORS = "2DD99DA1-16FA-E811-A94C-000D3AF06091";
    }
    public static class JobSubStatus
    {
        public const string Canceled = "1527FA6C-FA0F-E911-A94E-000D3AF060A1";
        public const string Closed = "1727FA6C-FA0F-E911-A94E-000D3AF060A1";
        public const string KKGAuditFailed = "6C8F2123-5106-EA11-A811-000D3AF057DD";
        public const string PartPOCreated = "1B27FA6C-FA0F-E911-A94E-000D3AF060A1";
        public const string WorkDone = "2927FA6C-FA0F-E911-A94E-000D3AF060A1";
        public const string WorkDoneSMS = "7E85074C-9C54-E911-A951-000D3AF0677F";
        public const string WorkInitiated = "2B27FA6C-FA0F-E911-A94E-000D3AF060A1";
    }
}
