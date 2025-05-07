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
    public class GrievanceHandlingActivityPostUpdate : IPlugin
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
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                        && context.PrimaryEntityName.ToLower() == "hil_grievancehandlingactivity" && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.Contains("statecode"))
                    {
                        if (entity.GetAttributeValue<OptionSetValue>("statecode").Value == 1)//Completed
                        {
                            Entity entGHA = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("actualstart"));
                            DateTime _actualStart = entGHA.GetAttributeValue<DateTime>("actualstart").AddMinutes(330);
                            DateTime _actualEnd = DateTime.Now.AddMinutes(330);
                            TimeSpan ts = _actualEnd - _actualStart;
                            Entity entUpdateGHA = new Entity("hil_grievancehandlingactivity", entity.Id);
                            entUpdateGHA["actualend"] = _actualEnd;
                            entUpdateGHA["actualdurationminutes"] = ts.TotalMinutes;
                            service.Update(entUpdateGHA);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error : " + ex.Message);
            }
        }
    }
}
