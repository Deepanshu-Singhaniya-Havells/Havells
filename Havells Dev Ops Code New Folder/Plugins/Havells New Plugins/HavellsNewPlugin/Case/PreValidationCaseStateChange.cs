using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel;

namespace HavellsNewPlugin.Case
{
    public class PreValidationCaseStateChange : IPlugin
    {
        private static readonly Guid SamparkDepartmentId = new Guid("7bf1705a-3764-ee11-8df0-6045bdaa91c3");//Sampark Call center
        private static readonly Guid JobSubStatus = new Guid("1727FA6C-FA0F-E911-A94E-000D3AF060A1"); //Closed
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {

            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            //try
            //{
            if (context.InputParameters.Contains("EntityMoniker") &&
                context.InputParameters["EntityMoniker"] is EntityReference)
            {
                var entityRef = (EntityReference)context.InputParameters["EntityMoniker"];
                var state = (OptionSetValue)context.InputParameters["State"];
                var status = (OptionSetValue)context.InputParameters["Status"];

                Entity Case = service.Retrieve("incident", entityRef.Id, new ColumnSet("hil_casedepartment", "hil_job"));
                Guid DepartmentId = Case.GetAttributeValue<EntityReference>("hil_casedepartment").Id;
                if (DepartmentId != SamparkDepartmentId)
                {
                    Guid JobId = Guid.Empty;
                    if (Case.Contains("hil_job"))
                    {
                        JobId = Case.GetAttributeValue<EntityReference>("hil_job").Id;

                        Entity Job = service.Retrieve("msdyn_workorder", JobId, new ColumnSet("msdyn_substatus", "msdyn_name"));
                        if (Job.Contains("msdyn_substatus"))
                        {
                            Guid susStatusId = Job.GetAttributeValue<EntityReference>("msdyn_substatus").Id;
                            string jobNumber = Job.Contains("msdyn_name") ? Job.GetAttributeValue<string>("msdyn_name") : "No job number found";
                            if (susStatusId != JobSubStatus)
                            {
                                throw new InvalidPluginExecutionException($"Corresponding Job({jobNumber}) should be closed in order to close the case");
                            }
                        }
                    }
                }
            }
        }
    }
}