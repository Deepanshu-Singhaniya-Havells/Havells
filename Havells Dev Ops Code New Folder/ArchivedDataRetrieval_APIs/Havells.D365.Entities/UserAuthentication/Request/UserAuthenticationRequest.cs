using System.ComponentModel.DataAnnotations;

namespace Havells.D365.Entities.UserAuthentication.Request
{
    public class UserAuthenticationRequest
    {
        [Required(ErrorMessage = "Please Enter UserId")]
        public string UserId { get; set; }

        [Required(ErrorMessage = "Please Enter Password")]
        public string Password { get; set; }
    }
}
