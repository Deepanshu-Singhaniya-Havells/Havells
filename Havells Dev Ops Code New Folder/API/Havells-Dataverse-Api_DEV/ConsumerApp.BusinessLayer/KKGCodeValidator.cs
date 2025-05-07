using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;


namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class KKGCodeValidator
    {
        [DataMember]
        public string JobID { get; set; }
        [DataMember]
        public string KKGCode { get; set; }
        [DataMember]
        public string Source { get; set; }
        [DataMember]
        public string UserId { get; set; }

        public ValidateKKGCodeResult ValidateKKGCode(KKGCodeValidator jobData) {
            ValidateKKGCodeResult responseObj = null;
            QueryExpression query;
            EntityCollection entcoll;
            string createdOn = string.Empty;
            string primaryKey = string.Empty;
            string kkgOTP = string.Empty;
            string kkgOTPDecrypted = string.Empty;
            Entity entTemp = null;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    query = new QueryExpression(msdyn_workorder.EntityLogicalName);
                    query.ColumnSet = new ColumnSet("createdon", "hil_kkgotp", "ownerid", "msdyn_workorderid");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, jobData.JobID);
                    entcoll = service.RetrieveMultiple(query);
                    if (entcoll.Entities.Count > 0)
                    {
                        entTemp = entcoll.Entities[0];
                        createdOn = entTemp.GetAttributeValue<DateTime>("createdon").ToString("MM/dd/yyyy");
                        kkgOTP = entTemp.GetAttributeValue<string>("hil_kkgotp"); //2qzRtOxT7zcIFEoKYtPNZg==

                        string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
                        "<entity name='hil_kkgcodeconfiguration'>" +
                        "<attribute name='hil_name' />" +
                        "<filter type='and'>" +
                        "<condition attribute='createdon' operator='on' value='" + createdOn + @"' />" +
                        "</filter>" +
                        "</entity>" +
                        "</fetch>";
                        entcoll = service.RetrieveMultiple(new FetchExpression(fetchXML));
                        if (entcoll.Entities.Count > 0)
                        {
                            primaryKey = entcoll.Entities[0].GetAttributeValue<string>("hil_name");
                            kkgOTPDecrypted = AESDecryption(kkgOTP, primaryKey);
                            if (kkgOTPDecrypted == jobData.KKGCode)
                            {
                                Entity ent = new Entity("hil_kkgcodeattempts");
                                ent["hil_jobid"] = entTemp.ToEntityReference();
                                ent["hil_kkgcode"] = jobData.KKGCode;
                                ent["hil_source"] = new OptionSetValue(Convert.ToInt32(jobData.Source));
                                ent["hil_status"] = true;
                                ent["hil_userid"] = new EntityReference("systemuser", new Guid(jobData.UserId));
                                service.Create(ent);
                                responseObj = new ValidateKKGCodeResult()
                                {
                                    StatusCode = "200",
                                    StatusDescription = "OK"
                                };
                            }
                            else {
                                Entity ent = new Entity("hil_kkgcodeattempts");
                                ent["hil_jobid"] = entTemp.ToEntityReference();
                                ent["hil_kkgcode"] = jobData.KKGCode;
                                ent["hil_source"] = new OptionSetValue(Convert.ToInt32(jobData.Source));
                                ent["hil_status"] = false;
                                ent["hil_userid"] = new EntityReference("systemuser", new Guid(jobData.UserId));
                                service.Create(ent);
                                responseObj = new ValidateKKGCodeResult()
                                {
                                    StatusCode = "204",
                                    StatusDescription = "Invalid KKG Code."
                                };
                            }
                        }
                    }
                    else
                    {
                        responseObj = new ValidateKKGCodeResult()
                        {
                            StatusCode = "204",
                            StatusDescription = "Job Id does not exist."
                        };
                    }
                }
                else
                {
                    responseObj = new ValidateKKGCodeResult { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                responseObj = new ValidateKKGCodeResult()
                {
                    StatusCode = "500",
                    StatusDescription = "D365 Internal Server Error : " + ex.Message
                };
            }
            return responseObj;
        }

        private static string AESDecryption(string EncryptedText, string Encryptionkey)
        {
            RijndaelManaged objrij = new RijndaelManaged();
            objrij.Mode = CipherMode.CBC;
            objrij.Padding = PaddingMode.PKCS7;
            objrij.KeySize = 0x80;
            objrij.BlockSize = 0x80;
            byte[] encryptedTextByte = Convert.FromBase64String(EncryptedText);
            byte[] passBytes = Encoding.UTF8.GetBytes(Encryptionkey);
            byte[] EncryptionkeyBytes = new byte[0x10];
            int len = passBytes.Length;
            if (len > EncryptionkeyBytes.Length)
            {
                len = EncryptionkeyBytes.Length;
            }
            Array.Copy(passBytes, EncryptionkeyBytes, len);
            objrij.Key = EncryptionkeyBytes;
            objrij.IV = EncryptionkeyBytes;
            byte[] TextByte = objrij.CreateDecryptor().TransformFinalBlock(encryptedTextByte, 0, encryptedTextByte.Length);
            return Encoding.UTF8.GetString(TextByte);  //it will return readable string
        }
        public ValidateKKGCodeResult KKGCodeVerification(ValidateKKGCodeInput validateKKGCodeInput)
        {
            bool? hash = null;
            string statusMessage = "";
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgServiceRebuild();
                if (service != null)
                {
                    QueryExpression query = new QueryExpression("hil_jobsauth");
                    query.ColumnSet = new ColumnSet("hil_checksum", "hil_hash", "hil_salt");
                    query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, validateKKGCodeInput._jobId));
                    query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1));
                    EntityCollection entColl = service.RetrieveMultiple(query);
                    if (entColl.Entities.Count == 1)
                    {
                        string Checksum = entColl[0].GetAttributeValue<string>("hil_checksum");
                        string _HashOld = entColl[0].GetAttributeValue<string>("hil_hash");
                        string _salt = entColl[0].GetAttributeValue<string>("hil_salt");
                        string _NewHash = getHash(_salt + validateKKGCodeInput._KKGCode);
                        if (_NewHash == _HashOld)
                        {
                            hash = true;
                            statusMessage = "Success";
                        }
                        else
                        {
                            hash = false;
                            statusMessage = "Invalid KKG Code";
                        }
                    }
                    else
                    {
                        query = new QueryExpression("msdyn_workorder");
                        query.ColumnSet = new ColumnSet("hil_kkgotp");
                        query.Criteria.AddCondition(new ConditionExpression("msdyn_name", ConditionOperator.Equal, validateKKGCodeInput._jobId));
                        entColl = service.RetrieveMultiple(query);
                        if (entColl.Entities.Count == 1)
                        {
                            string kkgOtp = entColl[0].GetAttributeValue<string>("hil_kkgotp");
                            var base64Bytes = System.Convert.FromBase64String(kkgOtp);
                            kkgOtp = Encoding.UTF8.GetString(base64Bytes);
                            if (kkgOtp == validateKKGCodeInput._KKGCode)
                            {
                                hash = true;
                                statusMessage = "Success";
                            }
                            else
                            {
                                hash = false;
                                statusMessage = "Invalid KKG Code";
                            }
                        }
                        else
                        {
                            return new ValidateKKGCodeResult { IsVarified = false, StatusCode = "503", StatusDescription = "Invalid Job Id." };
                        }
                    }
                }
                else
                {
                    return new ValidateKKGCodeResult { IsVarified = hash, StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                return new ValidateKKGCodeResult()
                {
                    IsVarified = hash,
                    StatusCode = "500",
                    StatusDescription = "D365 Internal Server Error : " + ex.Message
                };
            }
            return new ValidateKKGCodeResult { IsVarified = hash, StatusCode = "200", StatusDescription = statusMessage };
        }
        private static string getHash(string text)
        {
            using (var sha512 = SHA512.Create())
            {
                var hashedBytes = sha512.ComputeHash(Encoding.UTF8.GetBytes(text));
                return BitConverter.ToString(hashedBytes).Replace("-", "").ToLower();
            }
        }

    }
    [DataContract]
    public class ValidateKKGCodeResult
    {
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public bool? IsVarified { get; set; }

        [DataMember]
        public string StatusDescription { get; set; }
    }
    [DataContract]
    public class ValidateKKGCodeInput
    {
        [DataMember]
        public string _jobId { get; set; }

        [DataMember]
        public string _KKGCode { get; set; }
    }

}
