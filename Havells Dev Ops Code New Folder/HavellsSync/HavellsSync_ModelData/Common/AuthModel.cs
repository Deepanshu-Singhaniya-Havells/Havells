using Newtonsoft.Json;
using Newtonsoft.Json.Schema;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_ModelData.Common
{
    public class AuthModel
    {
        public string LoginUserId { get; set; }
        public string SourceType { get; set; }
    }
    #region CommomData
    public class CommonReqRes
    {
        public object data { get; set; }
    }
    
    public class CommonReq<T>
    {
        public T data { get; set; } 
    }
    #endregion
    public class Param
    {
        public string objParam { get; set; } = string.Empty;
    }
    public class RequestStatus
    {
        public int StatusCode { get; set; }
        public string Message { get; set; }
    }
    public class AuthResponse : TokenExpires
    {
        public string LoginUserId { get; set; }
        public string SourceType { get; set; }
        public string AccessToken { get; set; }
        public int StatusCode { get; set; }
        public string Message { get; set; }
    }
    public class TokenExpires
    {
        public string TokenExpiresAt { get; set; }
    }
    public class ReqSourceInfo
    {
        public string SourceCode { get; set; }
        public string SourceType { get; set; }
        public string AMCSellingSource { get; set; }
    }
    public class ValidateSession
    {
        public bool KeepSessionLive { get; set; } = true;
        public string SessionId { get; set; }

    }
    public class validatesessionResponse : TokenExpires
    {
        public string JWTToken { get; set; }
        public string MobileNumber { get; set; }
        public string SessionId { get; set; }
        public string SourceCode { get; set; }
        public string SourceType { get; set; }
        public int StatusCode { get; set; }
        public string Message { get; set; }
    }
    public class AppSettings
    {
        public string Secret { get; set; }
    }

}
