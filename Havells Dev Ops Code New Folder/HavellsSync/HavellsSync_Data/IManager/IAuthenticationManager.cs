using HavellsSync_ModelData.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsSync_Data.IManager
{
    public interface IAuthenticationManager
    {
        AuthResponse AuthenticateUser(AuthModel objAuth);
        validatesessionResponse ValidateSessionDetails(ValidateSession requestParam, string LoginUserId);
    }
}
