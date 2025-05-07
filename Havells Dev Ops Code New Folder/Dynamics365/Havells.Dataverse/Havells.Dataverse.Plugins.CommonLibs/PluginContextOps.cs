using Microsoft.Xrm.Sdk;

namespace Havells.Dataverse.Plugins.CommonLibs
{
    public class PluginContextOps
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="context">Execution Context (run-time environment)</param>
        /// <param name="_messageName">Name of the Web service message</param>
        /// <param name="_primaryEntityName">Name of the primary entity</param>
        /// <param name="_depth">Current depth of execution in the call stac</param>
        /// <returns></returns>
        public static bool ValidateContext(IPluginExecutionContext context, string _messageName, string _primaryEntityName, int _depth)
        {
            bool _retVal = false;
            if (
                    context.InputParameters.Contains("Target") &&
                    context.InputParameters["Target"] is Entity &&
                    context.PrimaryEntityName.ToUpper() == _primaryEntityName.ToUpper() &&
                    context.MessageName.ToUpper() == _messageName.ToUpper() &&
                    context.Depth == _depth
               )
                _retVal = true;

            return _retVal;
        }
    }
    public class PluginMessages {
        public const string Create = "Create";
    }
    public class MDM
    {
        public const string Address = "hil_address";
        public const string Account = "account";
        public const string Contact = "contact";
    }
}
