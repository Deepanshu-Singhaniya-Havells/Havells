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
    public class PostUpdate : IPlugin
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
                    && context.MessageName.ToUpper() == "UPDATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("1");
                    //EvaluateCharhes(entity, service);
                    if(entity.Contains("hil_iswrongjobclosure") && entity.Attributes.Contains("hil_iswrongjobclosure"))
                    {
                        OptionSetValue opWrongJob = entity.GetAttributeValue<OptionSetValue>("hil_iswrongjobclosure");
                        if(opWrongJob.Value == 1)
                        {
                            OnWrongJobClosureInClaimLine(service, entity.Id);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ClaimLine.PreCreate.Execute" + ex.Message);
            }
            #endregion
        }
        #region Evaluate Charges
        public static void EvaluateCharhes(Entity entity, IOrganizationService service)
        {
            decimal temp = 0;
            hil_ClaimLines iClaim = (hil_ClaimLines)service.Retrieve(hil_ClaimLines.EntityLogicalName, entity.Id, new ColumnSet(true));
            if(entity.Attributes.Contains("hil_charges"))
            {
                temp = Convert.ToDecimal(iClaim.hil_Charges);
                hil_ClaimHeader Header = (hil_ClaimHeader)service.Retrieve(hil_ClaimHeader.EntityLogicalName, iClaim.hil_ClaimHeader.Id, new ColumnSet("hil_totalclaimvalue"));
                if (Header.hil_TotalClaimValue != null)
                {
                    decimal iTemp = 0;
                    iTemp = Convert.ToDecimal(Header.hil_TotalClaimValue);
                    Header.hil_TotalClaimValue = Convert.ToDecimal(iTemp + temp);
                }
                else
                {
                    Header.hil_TotalClaimValue = Convert.ToDecimal(temp);
                }
                service.Update(Header);
            }
        }
        #endregion
        #region On Claim Line Wrong Job Closure
        public static void OnWrongJobClosureInClaimLine(IOrganizationService service, Guid ClaimLineId)
        {
            hil_ClaimLines enClaim = new hil_ClaimLines();
            enClaim.Id = ClaimLineId;
            Int32 WrongJobPenalty = GetWrongJobClosurePenalty(service);
            enClaim.hil_WrongCallClosurePenalty = WrongJobPenalty;
            enClaim.hil_TatIncentive = Convert.ToDecimal(0);
            enClaim.hil_TatBreachedPenalty = 0;
            enClaim.hil_IncentiveforMobileAppClosure = 0;
            enClaim.hil_Charges = Convert.ToDecimal(0);
            service.Update(enClaim);
        }
        #region Get Wrong Job Closure Penalty
        public static Int32 GetWrongJobClosurePenalty(IOrganizationService service)
        {
            Int32 dePenalty = 0;
            QueryExpression qe = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
            qe.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "ClaimJobWrongClosurePenalty");
            //qe.Criteria.AddCondition("hil_mobileappclosure", ConditionOperator.Equal, 1);
            qe.ColumnSet = new ColumnSet("hil_priceforpenalty");
            EntityCollection enColl = service.RetrieveMultiple(qe);
            int count = enColl.Entities.Count;
            foreach (Entity en in enColl.Entities)
            {
                if (en.Contains("hil_priceforpenalty"))
                {
                    dePenalty = en.GetAttributeValue<Int32>("hil_priceforpenalty");
                }
            }
            return dePenalty;
        }
        #endregion
        #endregion
    }
}
