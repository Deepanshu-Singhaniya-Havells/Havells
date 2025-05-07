using System;
using Havells_Plugin;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin.Characterstics
{
    public class PreCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "msdyn_workorder"
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    SetAutoNumber(service, entity);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Characterstic.Characterstic_Pre_Create.SetAutoNumber: " + ex.Message);
            }
            #endregion
        }
        public void SetAutoNumber(IOrganizationService service,Entity Ety)
        {
            String sPreFix = WorkOrder.PreCreate.getAutoNumberPreFix();
            int Sequenium = new int();
            QueryByAttribute Qry = new QueryByAttribute();
            Qry.EntityName = hil_integrationconfiguration.EntityLogicalName;
            ColumnSet Col = new ColumnSet("hil_url");
            Qry.AddAttributeValue("hil_name", "Characteristics AutoNumber");
            Qry.ColumnSet = Col;
            EntityCollection Found = service.RetrieveMultiple(Qry);
            if(Found.Entities.Count >= 1)
            {
                foreach(hil_integrationconfiguration inte in Found.Entities)
                {
                    Characteristic Chrt = (Characteristic)service.Retrieve(Characteristic.EntityLogicalName, Ety.Id, new ColumnSet(true));
                    string AutoNum = string.Empty;
                    AutoNum = sPreFix + inte.hil_Url;
                    Chrt.Name = AutoNum;
                    Sequenium = Convert.ToInt16(inte.hil_Url);
                    Sequenium = Sequenium + 1;
                    inte.hil_Url = Convert.ToString(Sequenium);
                    service.Update(inte);
                }
            }
            


        }
    }
}
