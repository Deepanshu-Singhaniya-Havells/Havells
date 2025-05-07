using Havells_Plugin.SAWActivity;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

namespace Havells_Plugin.Tender
{
    public class TenderLinePostUpdate : IPlugin
    {
        static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            Guid tenderID;
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target")
                && context.InputParameters["Target"] is Entity
                && context.PrimaryEntityName.ToLower() == "hil_tenderproduct"
                && context.Depth == 1)
                {
                    tracingService.Trace("1");
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tenderID = ((EntityReference)service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_tenderid"))["hil_tenderid"]).Id;
                    string _tenderProductfetch = @"<fetch version='1.0' mapping='logical' distinct='false' aggregate='true'>
                    <entity name='hil_tenderproduct'>
                    <attribute name='hil_tenderid' alias='tenderNo' groupby='true' />
                    <attribute name='hil_totalvalueinrs' alias='amount' aggregate='sum' />
                    <filter type='and'>
                    <condition attribute='hil_tenderid' operator='eq' value='" + tenderID + @"' />
                    <condition attribute='statecode' operator='eq' value='0' />
                    </filter>
                    </entity>
                    </fetch>";
                    EntityCollection _tenderproductColl = service.RetrieveMultiple(new FetchExpression(_tenderProductfetch));
                    Money FinalAmount = ((Money)((AliasedValue)_tenderproductColl[0]["amount"]).Value);
                    Entity tender = new Entity("hil_tender");
                    tender["hil_tendercost"] = FinalAmount;
                    tender.Id = tenderID;
                    service.Update(tender);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("HavellsNewPlugin.TenderModule.TenderProductPostUpdate_UpdateFinalOfferCost.Execute Error " + ex.Message);
            }
        }
    }
}
