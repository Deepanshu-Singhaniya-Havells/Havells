using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Havells_Plugin
{
    public class KKGCodeHashing
    {
        private static Random generator = new Random();
        private static string encryptKey = "AEs0PeR @T!0N";
        public static bool? GenerateKKGCodeHashVerification(string _jobId, String _KKGCode, IOrganizationService service)
        {
            bool? hash = null;
            QueryExpression query = new QueryExpression("hil_jobsauth");
            query.ColumnSet = new ColumnSet("hil_checksum", "hil_hash", "hil_salt");
            query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, _jobId));
            query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1));
            EntityCollection entColl = service.RetrieveMultiple(query);
            if (entColl.Entities.Count == 1)
            {
                string Checksum = entColl[0].GetAttributeValue<string>("hil_checksum");
                string _HashOld = entColl[0].GetAttributeValue<string>("hil_hash");
                string _salt = entColl[0].GetAttributeValue<string>("hil_salt");
                string _NewHash = getHash(_salt + _KKGCode);
                if (_NewHash == _HashOld)
                {
                    hash = true;
                }
                else
                    hash = false;
            }
            else
                hash = null;
            return hash;
        }
        public static string GenerateKKGCodeHash(string _jobId, IOrganizationService service)
        {
            string _KKGCode = null;
            QueryExpression query = new QueryExpression("hil_jobsauth");
            query.ColumnSet = new ColumnSet(false);
            query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, _jobId));
            query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1));
            EntityCollection entColl = service.RetrieveMultiple(query);

            foreach (Entity ent in entColl.Entities)
            {
                SetStateRequest deactivateRequest = new SetStateRequest
                {
                    EntityMoniker = ent.ToEntityReference(),
                    State = new OptionSetValue(1),
                    Status = new OptionSetValue(2)
                };
                service.Execute(deactivateRequest);
            }
            string Checksum = getChecksum(_jobId).ToString();
            _KKGCode = GenerateRandomNo();
            Entity entJobAuth = new Entity("hil_jobsauth");
            string _salt = getSalt(Checksum);
            entJobAuth["hil_checksum"] = Checksum;
            entJobAuth["hil_salt"] = _salt;
            entJobAuth["hil_hash"] = getHash(_salt + _KKGCode);
            entJobAuth["hil_name"] = _jobId;
            service.Create(entJobAuth);
            return _KKGCode;
        }
        public static string GetKKGCode(string _jobId, IOrganizationService service)
        {
            string fetch = $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                                  <entity name=""hil_jobsextension"">
                                   <attribute name=""hil_jobsextensionid"" />
                                    <attribute name=""hil_name"" />
                                    <attribute name=""hil_kkgencription"" />
                                    <order attribute=""hil_name"" descending=""false"" />
                                    <link-entity name=""msdyn_workorder"" from=""msdyn_workorderid"" to=""hil_jobs"" link-type=""inner"" alias=""ab"">
                                      <filter type=""and"">
                                        <condition attribute=""msdyn_name"" operator=""eq"" value=""{_jobId}"" />
                                      </filter>
                                    </link-entity>
                                  </entity>
                                </fetch>";
            EntityCollection entColl = service.RetrieveMultiple(new FetchExpression(fetch));
            if (entColl.Entities.Count >= 1)
            {
                if (entColl[0].Contains("hil_kkgencription"))
                    if (entColl[0].GetAttributeValue<bool>("hil_kkgencription"))
                        return GenerateKKGCodeHash(_jobId, service);
                    else
                        return GetKKGOTP(_jobId, service);
                else
                    return GetKKGOTP(_jobId, service);
            }
            else
            {
                return GetKKGOTP(_jobId, service);
            }
        }
        public static string GetKKGOTP(string workOrderId, IOrganizationService service)
        {
            string _kkgOPT = string.Empty;
            try
            {
                QueryExpression Query = new QueryExpression(msdyn_workorder.EntityLogicalName);
                Query.ColumnSet = new ColumnSet("hil_kkgotp");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("msdyn_name", ConditionOperator.Equal, workOrderId));
                EntityCollection enCol = service.RetrieveMultiple(Query);
                if (enCol.Entities.Count > 0)
                {
                    _kkgOPT = enCol.Entities[0].GetAttributeValue<string>("hil_kkgotp");
                    _kkgOPT = Havells_Plugin.WorkOrder.Common.Base64Decode(_kkgOPT);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.SMS_PreOp_Create.Execute.GetKKGOTP" + ex.Message);
            }
            return _kkgOPT;
        }
        private static string GenerateRandomNo()
        {
            return generator.Next(0, 999999).ToString("D6");
        }
        private static int getChecksum(string _jobId)
        {
            int _checkSum = 0;
            for (int i = 0; i < 3; i++)
            {
                int _randomNum = generator.Next(1, _jobId.Length);
                _checkSum += Convert.ToInt32(_jobId.Substring(_randomNum - 1, 1));
            }
            return _checkSum;
        }
        private static string getSalt(string Checksum)
        {
            byte[] bytes = new byte[256 / 8]; // 32bit
            using (var keyGenerator = RandomNumberGenerator.Create())
            {
                keyGenerator.GetBytes(bytes);
                return BitConverter.ToString(bytes).Replace("-", "").ToLower() + Checksum;
            }
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
}
