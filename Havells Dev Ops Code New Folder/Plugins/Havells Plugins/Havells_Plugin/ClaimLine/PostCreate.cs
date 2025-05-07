using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
// Microsoft Dynamics CRM namespace(s)
using System.Net;
using System.Net.Http;
using System.Runtime.Serialization.Json;
using System.IO;
using Havells_Plugin.HelperIntegration;
using System.Linq;
using System;
using System.Linq;
using System.Text;
using Havells_Plugin;
using Microsoft.Xrm.Sdk;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;

namespace Havells_Plugin.ClaimLine
{
    public class PostCreate : IPlugin
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
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == hil_ClaimLines.EntityLogicalName.ToLower()
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("1");
                    TotalCharges(entity, service);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ClaimLine.PreCreate.Execute" + ex.Message);
            }
            #endregion
        }
        public static void TotalCharges(Entity Claim, IOrganizationService service)
        {
            hil_ClaimLines iClaimLine = Claim.ToEntity<hil_ClaimLines>();
            if(iClaimLine.hil_ClaimHeader != null)
            {
                decimal temp = 0;
                hil_ClaimHeader Header = (hil_ClaimHeader)service.Retrieve(hil_ClaimHeader.EntityLogicalName, iClaimLine.hil_ClaimHeader.Id, new ColumnSet(true));
                if(iClaimLine.hil_IncentiveforMobileAppClosure != null)
                {
                    temp = temp + Convert.ToDecimal(iClaimLine.hil_IncentiveforMobileAppClosure);
                }
                if(iClaimLine.hil_TatBreachedPenalty != null)
                {
                    temp = temp - Convert.ToDecimal(iClaimLine.hil_TatBreachedPenalty);
                }
                if(Header.hil_TotalClaimValue != null)
                {
                    Header.hil_TotalClaimValue = Convert.ToDecimal(Header.hil_TotalClaimValue) + temp;
                }
                else
                {
                    Header.hil_TotalClaimValue = temp;
                }
                service.Update(Header);
            }
        }
    }
}
