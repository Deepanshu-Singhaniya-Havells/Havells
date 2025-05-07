using System;
using System.Configuration;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace SOPaymentReceipt
{
	public class ClsEncryptDecrypt
    {
        private static readonly string _Key = ConfigurationManager.AppSettings["EncryptionKeyAES256"];
        private static readonly string _iv = ConfigurationManager.AppSettings["IV"];
        public string EncryptAES256URL(string plainText)
        {
            try
            {
                if (!string.IsNullOrEmpty(plainText))
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
                    // return WebConfigurationManager.AppSettings["Docurl"] + Convert.ToBase64String(encrypted);
                    return Convert.ToBase64String(encrypted);
                }
                else
                {
                    return null;
                }
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
                if (!string.IsNullOrEmpty(encryptedText))
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
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

    }
}
