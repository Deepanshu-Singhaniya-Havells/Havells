using HavellsSync_ModelData.ICommon;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_ModelData.Common
{
    public class AES256 : IAES256
    {
        //private static string _Key= "T6bpUDhl9e6MEy9w3Ck85BgSiC1Nw56U";
        //private static string _iv = "okBrFHCIAA8It4qf";
        private static string _Key = "";
        private static string _iv = "";
        public AES256(IConfiguration configuration)
        {
            Check.Argument.IsNotNull(nameof(configuration), configuration);
            _Key = configuration.GetSection(ConfigKeys.ASE256Details + ":" + ConfigKeys.EncryptionKeyAES256).Value;
            _iv = configuration.GetSection(ConfigKeys.ASE256Details + ":" + ConfigKeys.IV).Value;
        }
        public string EncryptAES256(string plainText)
        {
            try
            {
                byte[] key32 = Encoding.UTF8.GetBytes(_Key);
                byte[] IV16 = Encoding.UTF8.GetBytes(_iv);
                // Check arguments.
                if (plainText == null || plainText.Length <= 0)
                    throw new ArgumentNullException("plainText");

                byte[] encrypted;
                // Create an AesManaged object
                // with the specified key and IV.
                using (AesManaged aesAlg = new AesManaged())
                {
                    aesAlg.KeySize = 256;
                    aesAlg.Padding = PaddingMode.PKCS7;
                    aesAlg.Key = key32;
                    aesAlg.IV = IV16;
                    // Create an encryptor to perform the stream transform.
                    ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
                    // Create the streams used for encryption.
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
                // Return the encrypted base64 from the memory stream.
                return Convert.ToBase64String(encrypted);
            }
            catch (Exception)
            {
                return null;
            }
        }

        public string DecryptAES256(string encryptedText)
        {
            try
            {
                byte[] key32 = Encoding.UTF8.GetBytes(_Key);
                byte[] IV16 = Encoding.UTF8.GetBytes(_iv);

                // Check arguments.
                if (encryptedText == null || encryptedText.Length <= 0)
                    throw new ArgumentNullException("encryptedText");

                byte[] cipherText = Convert.FromBase64String(encryptedText);
                // Declare the string used to hold
                // the decrypted text.
                string plaintext = null;

                // Create an AesManaged object
                // with the specified key and IV.
                using (AesManaged aesAlg = new AesManaged())
                {
                    aesAlg.KeySize = 256;
                    aesAlg.Padding = PaddingMode.PKCS7;
                    aesAlg.Key = key32;
                    aesAlg.IV = IV16;

                    // Create a decryptor to perform the stream transform.
                    ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);
                    // Create the streams used for decryption.
                    using (MemoryStream msDecrypt = new MemoryStream(cipherText))
                    {
                        using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                        {
                            using (StreamReader srDecrypt = new StreamReader(csDecrypt))
                            {
                                // Read the decrypted bytes from the decrypting stream
                                // and place them in a string.
                                plaintext = srDecrypt.ReadToEnd();
                            }
                        }
                    }
                }
                return plaintext;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public string EncryptAES256AMC(string plainText, string Aes256kEY)
        {
            try
            {
                byte[] encrypted = null;
                if (!string.IsNullOrEmpty(plainText) && !string.IsNullOrEmpty(Aes256kEY))
                {
                    var KeyData = Aes256kEY.ToString().Split('.')[1];
                    var _Key = KeyData.Substring(KeyData.Length - 32);
                    var _iv = _Key.Substring(_Key.Length - 16);
                    byte[] key32 = Encoding.UTF8.GetBytes(_Key);
                    byte[] IV16 = Encoding.UTF8.GetBytes(_iv);
                    // Check arguments.
                    if (plainText == null || plainText.Length <= 0)
                        throw new ArgumentNullException("plainText");


                    // Create an AesManaged object
                    // with the specified key and IV.
                    using (AesManaged aesAlg = new AesManaged())
                    {
                        aesAlg.KeySize = 256;
                        aesAlg.Padding = PaddingMode.PKCS7;
                        aesAlg.Key = key32;
                        aesAlg.IV = IV16;
                        // Create an encryptor to perform the stream transform.
                        ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
                        // Create the streams used for encryption.
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

                    // Return the encrypted base64 from the memory stream.

                }
                return Convert.ToBase64String(encrypted);
            }
            catch (Exception)
            {
                return null;
            }
        }

        public string DecryptAES256AMC(string encryptedText, string Aes256kEY)
        {
            try
            {
                string plaintext = null;
                if (!string.IsNullOrEmpty(encryptedText) && !string.IsNullOrEmpty(Aes256kEY))
                {
                    var KeyData = Aes256kEY.ToString().Split('.')[1];
                    var _Key = KeyData.Substring(KeyData.Length - 32);
                    var _iv = _Key.Substring(_Key.Length - 16);
                    byte[] key32 = Encoding.UTF8.GetBytes(_Key);
                    byte[] IV16 = Encoding.UTF8.GetBytes(_iv);

                    // Check arguments.
                    if (encryptedText == null || encryptedText.Length <= 0)
                        throw new ArgumentNullException("encryptedText");

                    byte[] cipherText = Convert.FromBase64String(encryptedText);
                    // Declare the string used to hold
                    // the decrypted text.

                    // Create an AesManaged object
                    // with the specified key and IV.
                    using (AesManaged aesAlg = new AesManaged())
                    {
                        aesAlg.KeySize = 256;
                        aesAlg.Padding = PaddingMode.PKCS7;
                        aesAlg.Key = key32;
                        aesAlg.IV = IV16;

                        // Create a decryptor to perform the stream transform.
                        ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);
                        // Create the streams used for decryption.
                        using (MemoryStream msDecrypt = new MemoryStream(cipherText))
                        {
                            using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                            {
                                using (StreamReader srDecrypt = new StreamReader(csDecrypt))
                                {
                                    // Read the decrypted bytes from the decrypting stream
                                    // and place them in a string.
                                    plaintext = srDecrypt.ReadToEnd();
                                }
                            }
                        }
                    }
                }
                return plaintext;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
    public class XmlTextSerializer
    {/// <summary>
     /// Method to get the Application Configuration Settings value using key name of settings
     /// </summary>
     /// <param name="key">string Key</param>
     /// <returns>returns application settings value</returns>
        public static string GetAppSettings(string key)
        {
            string value = string.Empty;
            try
            {
                value = ConfigurationManager.AppSettings[key];
            }
            catch (Exception ex)
            {
                value = string.Empty;
                throw ex;

            }
            return value;
        }
        public string EncryptAES256AMC(string plainText, string Aes256kEY)
        {
            try
            {
                byte[] encrypted = null;
                if (!string.IsNullOrEmpty(plainText) && !string.IsNullOrEmpty(Aes256kEY))
                {
                    var KeyData = Aes256kEY.ToString().Split('.')[1];
                    var _Key = KeyData.Substring(KeyData.Length - 32);
                    var _iv = _Key.Substring(_Key.Length - 16);
                    byte[] key32 = Encoding.UTF8.GetBytes(_Key);
                    byte[] IV16 = Encoding.UTF8.GetBytes(_iv);
                    // Check arguments.
                    if (plainText == null || plainText.Length <= 0)
                        throw new ArgumentNullException("plainText");


                    // Create an AesManaged object
                    // with the specified key and IV.
                    using (AesManaged aesAlg = new AesManaged())
                    {
                        aesAlg.KeySize = 256;
                        aesAlg.Padding = PaddingMode.PKCS7;
                        aesAlg.Key = key32;
                        aesAlg.IV = IV16;
                        // Create an encryptor to perform the stream transform.
                        ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
                        // Create the streams used for encryption.
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

                    // Return the encrypted base64 from the memory stream.

                }
                return Convert.ToBase64String(encrypted);
            }
            catch (Exception)
            {
                return null;
            }
        }

        public string DecryptAES256AMC(string encryptedText, string Aes256kEY)
        {
            try
            {
                string plaintext = null;
                if (!string.IsNullOrEmpty(encryptedText) && !string.IsNullOrEmpty(Aes256kEY))
                {
                    var KeyData = Aes256kEY.ToString().Split('.')[1];
                    var _Key = KeyData.Substring(KeyData.Length - 32);
                    var _iv = _Key.Substring(_Key.Length - 16);
                    byte[] key32 = Encoding.UTF8.GetBytes(_Key);
                    byte[] IV16 = Encoding.UTF8.GetBytes(_iv);

                    // Check arguments.
                    if (encryptedText == null || encryptedText.Length <= 0)
                        throw new ArgumentNullException("encryptedText");

                    byte[] cipherText = Convert.FromBase64String(encryptedText);
                    // Declare the string used to hold
                    // the decrypted text.
                    // Create an AesManaged object
                    // with the specified key and IV.
                    using (AesManaged aesAlg = new AesManaged())
                    {
                        aesAlg.KeySize = 256;
                        aesAlg.Padding = PaddingMode.PKCS7;
                        aesAlg.Key = key32;
                        aesAlg.IV = IV16;
                        // Create a decryptor to perform the stream transform.
                        ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);
                        // Create the streams used for decryption.
                        using (MemoryStream msDecrypt = new MemoryStream(cipherText))
                        {
                            using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                            {
                                using (StreamReader srDecrypt = new StreamReader(csDecrypt))
                                {
                                    // Read the decrypted bytes from the decrypting stream
                                    // and place them in a string.
                                    plaintext = srDecrypt.ReadToEnd();
                                }
                            }
                        }
                    }
                }
                return plaintext;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Method to get the Application Configuration Settings value using key name of settings
        /// </summary>
        /// <param name="key">config key name enum</param>
        /// <returns>returns application settings value</returns>
        public static string GetAppSettings(ConfigKeys key)
        {
            string value = string.Empty;
            try
            {
                value = ConfigurationManager.AppSettings[key.ToString()];
            }
            catch (Exception ex)
            {
                value = string.Empty;
                throw ex;

            }
            return value;
        }
    }
}
