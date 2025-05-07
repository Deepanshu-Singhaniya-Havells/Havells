using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_ModelData.ICommon
{
    public interface IAES256
    {
        string EncryptAES256AMC(string plainText, string Aes256kEY);
        string DecryptAES256AMC(string plainText, string Aes256kEY);
        string EncryptAES256(string plainText);
        string DecryptAES256(string plainText);
    }
}
