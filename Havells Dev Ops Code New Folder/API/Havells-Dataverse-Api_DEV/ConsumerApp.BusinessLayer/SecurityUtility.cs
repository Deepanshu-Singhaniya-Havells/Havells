using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    public class SecurityUtility
    {
        public static RijndaelManaged GetRijndaelManaged(String SecretKey)
        {
            var keyBytes = new byte[16];
            var secretKeyBytes = Encoding.UTF8.GetBytes(SecretKey);
            Array.Copy(secretKeyBytes, keyBytes, Math.Min(keyBytes.Length, secretKeyBytes.Length));
            return new RijndaelManaged
            {
                Mode = CipherMode.CBC,
                Padding = PaddingMode.PKCS7,
                KeySize = 128,
                BlockSize = 128,
                Key = keyBytes,
                IV = keyBytes
            };
        }

        public static byte[] Encrypt(byte[] plainBytes, RijndaelManaged rijndaelManaged)
        {
            return rijndaelManaged.CreateEncryptor()
                .TransformFinalBlock(plainBytes, 0, plainBytes.Length);
        }

        public static byte[] Decrypt(byte[] encryptedData, RijndaelManaged rijndaelManaged)
        {
            return rijndaelManaged.CreateDecryptor()
                .TransformFinalBlock(encryptedData, 0, encryptedData.Length);
        }

        /// <summary>
        /// Encrypts plaintext using AES 128bit key and a Chain Block Cipher and returns a base64 encoded string
        /// </summary>
        /// <param name="plainText">Plain text to encrypt</param>
        /// <param name="key">Secret key</param>
        /// <returns>Base64 encoded string</returns>
        public static String Encrypt(String plainText, String secretKey)
        {
            var plainBytes = Encoding.UTF8.GetBytes(plainText);
            return Convert.ToBase64String(Encrypt(plainBytes, GetRijndaelManaged(secretKey)));
        }

        /// <summary>
        /// Decrypts a base64 encoded string using the given key (AES 128bit key and a Chain Block Cipher)
        /// </summary>
        /// <param name="encryptedText">Base64 Encoded String</param>
        /// <param name="key">Secret Key</param>
        /// <returns>Decrypted String</returns>
        public static String Decrypt(String encryptedText, String secretKey)
        {
            var encryptedBytes = Convert.FromBase64String(encryptedText);
            return Encoding.UTF8.GetString(Decrypt(encryptedBytes, GetRijndaelManaged(secretKey)));
        }
        private static string EncryptAES256(string plainText)
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
                            swEncrypt.Write(plainText);
                        }
                        encrypted = msEncrypt.ToArray();
                    }
                }
            }
            return Convert.ToBase64String(encrypted);
        }

        public AuthModel AuthenticateSamparkAppLogin(AuthModel request_param)
        {
            AuthModel _retObj = new AuthModel() { user_id = request_param.user_id, status_code = "200", status_description = "OK" };
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    Entity entTemp = service.Retrieve("systemuser", new Guid(request_param.user_id), new ColumnSet("hil_employeecode"));
                    if (entTemp != null)
                    {
                        _retObj.user_id = entTemp.GetAttributeValue<string>("hil_employeecode");
                        _retObj.access_token = entTemp.GetAttributeValue<string>("hil_employeecode");
                        if (!string.IsNullOrEmpty(_retObj.access_token))
                        {
                            _retObj.access_token = EncryptAES256(_retObj.access_token);
                        }
                    }
                    else
                    {
                        _retObj.status_code = "204";
                        _retObj.status_description = "User does not exist";
                    }
                }
                else
                {
                    _retObj.status_code = "503";
                    _retObj.status_description = "D365 Service Unavailable";
                }
            }
            catch (Exception ex)
            {
                _retObj.status_code = "500";
                _retObj.status_description = "D365 Internal Server Error : " + ex.Message;
            }
            return _retObj;
        }
    }
    public class AuthModel
    {
        public string user_id { get; set; }
        public string access_token { get; set; }
        public string status_code { get; set; }
        public string status_description { get; set; }
    }
}
