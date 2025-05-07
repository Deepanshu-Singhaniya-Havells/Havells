using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace HavellsNewPlugin.Case
{
    public class ActionCaseReassignment : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is EntityReference)
                {
                    EntityReference ent = (EntityReference)context.InputParameters["Target"];

                    if (ent == null)
                    {
                        context.OutputParameters["StatusMessage"] = "Case Id is required.";
                    }
                    else
                    {
                        CaseAssignmentEngine _caseAssignment = new CaseAssignmentEngine();
                        Entity _ent = service.Retrieve(ent.LogicalName, ent.Id, new ColumnSet(true));
                        context.OutputParameters["StatusMessage"] =_caseAssignment.AssignCase(service, _ent);
                    }
                }
            }
            catch (Exception ex)
            {
                context.OutputParameters["StatusMessage"] = "ERROR! " + ex.Message;
            }
        }
    }
}
