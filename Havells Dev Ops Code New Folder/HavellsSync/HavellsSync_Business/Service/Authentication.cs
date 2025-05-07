using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using HavellsSync_ModelData.Common;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Extensions.Configuration;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Security.Cryptography;
using HavellsSync_Data.IManager;
using HavellsSync_Busines.IService;
using System.Reflection.PortableExecutable;

namespace HavellsSync_Busines.Service
{
    public class Authentication : IAuthentication
    {
        private readonly IAuthenticationManager _AuthManager;
        public Authentication(IAuthenticationManager AuthManager)
        {
            _AuthManager = AuthManager;
        }

        public AuthResponse AuthenticateUser(AuthModel objAuth)
        {
            return _AuthManager.AuthenticateUser(objAuth);
        }
        public validatesessionResponse ValidateSessionDetails(ValidateSession requestParam, string LoginUserId)
        {
            return _AuthManager.ValidateSessionDetails(requestParam, LoginUserId);
        }
        
    }
}
