using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.Plugin.Approval
{
    public class ApprovalAction : IPlugin
    {
        public static ITracingService tracingService = null;
        public string _primaryField = string.Empty;
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
                    entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));
                    tracingService.Trace(entity.LogicalName);
                   HelperClass.ActionOnApprovalofOCLPrice(service, entity, tracingService);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells.CRM.Plugin.Approval.ApprovalAction Error:- " + ex.Message);// "HavellsNewPlugin.TenderModule.OrderCheckList.OrderCheckListPostCreate.Execute Error " + ex.Message);
            }
        }
    }
}
