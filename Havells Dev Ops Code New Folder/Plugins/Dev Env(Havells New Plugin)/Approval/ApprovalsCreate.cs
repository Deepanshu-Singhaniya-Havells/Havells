using System;
using Microsoft.Xrm.Sdk;

namespace HavellsNewPlugin.Approval
{
    public class ApprovalsCreate : IPlugin
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
                tracingService.Trace("ApprovalsCreate plugin started  ");
                var entityName = context.InputParameters["EntityName"].ToString();
                var entityId = context.InputParameters["EntityID"].ToString();
                var purpose = context.InputParameters["Purpose"].ToString();
                tracingService.Trace("entityName      " + entityName);
                tracingService.Trace("purpose      " + purpose);
                ApprovalHelper.GetPrimaryIdFieldName(entityName, service, out _primaryField);
                ApprovalHelper.createApproval(entityName, entityId, purpose, _primaryField, service, tracingService);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("ApprovalsCreate Error:- " + ex.Message);// "HavellsNewPlugin.TenderModule.OrderCheckList.OrderCheckListPostCreate.Execute Error " + ex.Message);
            }
        }
    }
}
