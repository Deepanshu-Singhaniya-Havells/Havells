using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;


namespace HavellsNewPlugin.TenderModule.Attchments
{
    public class AttchmentDeletion : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
           ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                     && context.PrimaryEntityName.ToLower() == "hil_attachment" && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.Contains("hil_isdeleted"))
                    {
                        entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_documenttype"));
                        Entity documentType = service.Retrieve(((EntityReference)entity["hil_documenttype"]).LogicalName, ((EntityReference)entity["hil_documenttype"]).Id, new ColumnSet("hil_allowdelete"));
                        if (documentType.Contains("hil_allowdelete"))
                        {
                            if (!documentType.GetAttributeValue<bool>("hil_allowdelete"))
                            {
                                throw new InvalidPluginExecutionException("Access Denied!!");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error " + ex.Message);
            }
        }
    }
}
