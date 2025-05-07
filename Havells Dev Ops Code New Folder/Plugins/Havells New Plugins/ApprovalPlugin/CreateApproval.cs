using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.IdentityModel.Metadata;
using System.Linq;
using System.Runtime.Remoting.Services;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.Plugin.Approval
{
    public class CreateApproval : IPlugin
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
                var purpose = (int)context.InputParameters["Purpose"];
                var EntobjType = (int)context.InputParameters["EntityObject"];
                tracingService.Trace("entityName      " + entityName);
                tracingService.Trace("purpose      " + purpose);
                tracingService.Trace("EntobjType      " + EntobjType);
                tracingService.Trace("entityId      " + entityId);
                HelperClass.GetPrimaryIdFieldName(entityName, service, out _primaryField);
                HelperClass.CreateApprovals(service, tracingService, entityName, entityId, purpose, EntobjType);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells.CRM.Plugin.Approval.CreateApproval Error:- " + ex.Message);// "HavellsNewPlugin.TenderModule.OrderCheckList.OrderCheckListPostCreate.Execute Error " + ex.Message);
            }
        }
    }
}
