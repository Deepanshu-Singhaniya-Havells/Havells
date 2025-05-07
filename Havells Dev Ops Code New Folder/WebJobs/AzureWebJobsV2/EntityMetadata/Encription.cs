using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;

namespace EntityMetadata
{
    public class Encription
    {
        private static Random generator = new Random();
        public static string GenerateKKGCodeHash(string _jobId, IOrganizationService service)
        {
            string _KKGCode = null;
            QueryExpression query = new QueryExpression("hil_jobsauth");
            query.ColumnSet = new ColumnSet(false);
            query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, _jobId));
            query.Criteria.AddCondition(new ConditionExpression("statuscode", ConditionOperator.Equal, 1));
            EntityCollection entColl = service.RetrieveMultiple(query);
            if (entColl.Entities.Count == 0)
            {
                string Checksum = getChecksum(_jobId).ToString();
                _KKGCode = GenerateRandomNo();
                Entity entJobAuth = new Entity("hil_jobsauth");
                string _salt = getSalt(Checksum);
                entJobAuth["hil_checksum"] = Checksum;
                entJobAuth["hil_salt"] = _salt;
                entJobAuth["hil_hash"] = getHash(_salt + _KKGCode);
                entJobAuth["hil_name"] = _jobId;
                service.Create(entJobAuth);
            }
            else
            {
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
            }
            return _KKGCode;
        }
        public static void KKGCodeVerification(string _jobId, string _KKGCode, IOrganizationService service)
        {
            bool? hash = null;
            try
            {

                if (service != null)
                {
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
                    {
                        query = new QueryExpression("msdyn_workorder");
                        query.ColumnSet = new ColumnSet("hil_kkgotp");
                        query.Criteria.AddCondition(new ConditionExpression("msdyn_name", ConditionOperator.Equal, _jobId));
                        entColl = service.RetrieveMultiple(query);
                        if (entColl.Entities.Count == 1)
                        {
                            string kkgOtp = entColl[0].GetAttributeValue<string>("hil_kkgotp");
                            var base64Bytes = System.Convert.FromBase64String(kkgOtp);
                            kkgOtp = Encoding.UTF8.GetString(base64Bytes);
                            if (kkgOtp == _KKGCode)
                            {
                                hash = true;
                            }
                            else
                                hash = false;
                        }
                    }
                }
                else
                {

                }
            }
            catch (Exception ex)
            {

            }
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
