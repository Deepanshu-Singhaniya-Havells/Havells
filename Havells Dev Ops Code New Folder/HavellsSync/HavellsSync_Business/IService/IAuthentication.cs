using HavellsSync_ModelData.Common;
using Microsoft.Crm.Sdk.Messages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_Busines.IService 
{
    public interface IAuthentication
    {
        AuthResponse AuthenticateUser(AuthModel objAuth);
        validatesessionResponse ValidateSessionDetails(ValidateSession requestParam,string LoginUserId);
    }
  
}
