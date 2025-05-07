using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.AccessControl;

namespace HavellsNewPlugin.Case
{
    public class CaseAssignment : IPlugin
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {

           #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.PrimaryEntityName.ToLower() == "incident" && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    CaseAssignmentEngine _caseAssignmentEngine = new CaseAssignmentEngine();
                    _caseAssignmentEngine.AssignCase(service, entity);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error : " + ex.Message);
            }
        }
    }
}
