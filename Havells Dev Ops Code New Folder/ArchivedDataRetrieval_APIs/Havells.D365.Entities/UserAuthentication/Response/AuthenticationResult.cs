using Havells.D365.Entities.Common.Response;

namespace Havells.D365.Entities.UserAuthentication.Response
{
    public class AuthenticationResult:CommonResponse
    {
        public string AccessToken { get; set; }
    }
}
