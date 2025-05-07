using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.Integration_Staging
{
    public class PostCreate : IPlugin
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
                    && context.InputParameters["Target"] is Entity && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity Staging = (Entity)context.InputParameters["Target"];
                    if(Staging.Attributes.Contains("hil_productcode") && Staging.Attributes.Contains("hil_uomcode"))
                    {
                        string PdtCode = (string)Staging["hil_productcode"];
                        string Uom = (string)Staging["hil_uomcode"];
                        UpdateRecords(service, PdtCode, Uom, Staging);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Integration_Staging.Integration_Staging_PostCreate.Execute: " + ex.Message);
            }
        }
        public static void UpdateRecords(IOrganizationService service, string PdtCode, string Uom, Entity Staging)
        {
            Guid UnitId = new Guid();
            QueryByAttribute Query = new QueryByAttribute();
            Query.EntityName = UoM.EntityLogicalName;
            Query.AddAttributeValue("name", Uom);
            ColumnSet Col = new ColumnSet(true);
            Query.ColumnSet = Col;
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count >= 1)
            {
                foreach (UoM Unit in Found.Entities)
                {
                    UnitId = (Guid)Unit.UoMId;
                    Staging["hil_uom"] = new EntityReference(UoM.EntityLogicalName, UnitId);
                }
            }
            QueryByAttribute Qry = new QueryByAttribute();
            Qry.EntityName = Product.EntityLogicalName;
            Qry.AddAttributeValue("hil_uniquekey", PdtCode);
            ColumnSet Col1 = new ColumnSet(true);
            Qry.ColumnSet = Col1;
            EntityCollection Found1 = service.RetrieveMultiple(Qry);
            if (Found1.Entities.Count >= 1)
            {
                foreach (Product Pdt in Found1.Entities)
                {
                    Guid ProdId = new Guid();
                    ProdId = Pdt.Id;
                    Pdt.DefaultUoMId = new EntityReference(UoM.EntityLogicalName, UnitId);
                    Staging["hil_product"] = new EntityReference(Product.EntityLogicalName, ProdId);
                    service.Update(Pdt);
                }
            }
            service.Update(Staging);
        }
    }
}
