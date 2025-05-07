using System;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;

namespace Havells_Plugin.Estimate_Header
{
    public class PostUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "hil_estimate" && context.MessageName.ToUpper() == "UPDATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if(entity.Attributes.Contains("hil_branchheadapprovalstatus"))
                    {
                        OptionSetValue iApproval = (OptionSetValue)entity["hil_branchheadapprovalstatus"];
                        if (iApproval.Value == 1)
                        {
                            hil_estimate iEstimate = (hil_estimate)service.Retrieve(hil_estimate.EntityLogicalName, entity.Id, new ColumnSet("hil_job", "hil_totalcharges"));
                            if(iEstimate.hil_Job != null)
                            {
                                msdyn_workorder iJob = new msdyn_workorder();
                                iJob.Id = iEstimate.hil_Job.Id;
                                iJob.hil_EstimateChargesTotal = new Money(Convert.ToDecimal(iEstimate.hil_totalCharges));
                                iJob.hil_EstimatedChargeDecimal = iEstimate.hil_totalCharges;
                                service.Update(iJob);
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Estimate_Header.PostUpdate.Execute " + ex.Message.ToUpper());
            }
        }
    }
}
