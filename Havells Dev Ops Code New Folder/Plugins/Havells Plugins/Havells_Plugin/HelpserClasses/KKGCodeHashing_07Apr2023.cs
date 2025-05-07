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

        private static void GenerateKKGCodeHash(string _jobId, IOrganizationService service)
        {
            QueryExpression query = new QueryExpression("hil_jobsauth");
            query.ColumnSet = new ColumnSet(false);
            query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, _jobId));
            EntityCollection entColl = service.RetrieveMultiple(query);
            if (entColl.Entities.Count == 0)
            {
                string Checksum = getChecksum(_jobId).ToString();
                string _KKGCode = GenerateRandomNo();
                Entity entJobAuth = new Entity("hil_jobsauth");
                string _salt = getSalt(Checksum);
                entJobAuth["hil_Checksum"] = Checksum;
                entJobAuth["hil_Salt"] = _salt;
                entJobAuth["hil_Hash"] = getHash(_salt + _KKGCode);
                entJobAuth["hil_name"] = _jobId;
                service.Create(entJobAuth);
            }
        }
        private static string GenerateRandomNo()
        {
            return generator.Next(0, 9999).ToString("D4");
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
