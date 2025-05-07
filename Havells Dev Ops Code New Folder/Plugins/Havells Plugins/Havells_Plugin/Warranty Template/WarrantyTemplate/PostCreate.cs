using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin.WarrantyTemplate
{
    public class PostCreate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == hil_warrantytemplate.EntityLogicalName
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    hil_warrantytemplate _WrtyTplt = (hil_warrantytemplate)service.Retrieve(hil_warrantytemplate.EntityLogicalName, entity.Id,new ColumnSet(true));
                    LoadPMSConfigurations(service, _WrtyTplt);
                    InitiateProcessForParts(service, _WrtyTplt);
                    InitiateProcessForLabors(service, _WrtyTplt);
                }
            }
            catch (Exception ex)
            {
                //Namespace.class.Method
                throw new InvalidPluginExecutionException("Havells_Plugin.WarrantyTemplate.PostCreate.Execute" + ex.Message);
            }
            #endregion
        }

        public static void LoadPMSConfigurations(IOrganizationService service, hil_warrantytemplate _wrtyTplt)
        {
            try
            {
                if (_wrtyTplt.hil_Product != null)
                {
                    Entity ent = service.Retrieve(Product.EntityLogicalName, _wrtyTplt.hil_Product.Id, new ColumnSet("hil_division"));
                    if (ent != null)
                    {
                        QueryExpression qryExp = new QueryExpression("hil_pmsscheduleconfiguration");
                        qryExp.ColumnSet = new ColumnSet("hil_pmscount");
                        qryExp.Criteria = new FilterExpression(LogicalOperator.And);
                        qryExp.Criteria.AddCondition("hil_division", ConditionOperator.Equal, ent.GetAttributeValue<EntityReference>("hil_division").Id);
                        EntityCollection entCol = service.RetrieveMultiple(qryExp);
                        if (entCol.Entities.Count > 0)
                        {
                            if (entCol.Entities[0].Attributes.Contains("hil_pmscount"))
                            {
                                int pmscount = entCol.Entities[0].GetAttributeValue<int>("hil_pmscount");
                                if (pmscount > 0)
                                {
                                    int wrtyTempPeriod = Convert.ToInt32(_wrtyTplt.hil_WarrantyPeriod);
                                    int pmsConfigLines = Convert.ToInt32((wrtyTempPeriod * 1.0 / 12) * pmscount);
                                    int optValue = 910590000;
                                    if (pmsConfigLines > 0)
                                    {
                                        for (int i = 1; i <= pmsConfigLines; i++)
                                        {
                                            Entity entPmsconfiguration = new Entity("hil_pmsconfiguration");
                                            entPmsconfiguration["hil_warrantytemplate"] = _wrtyTplt.ToEntityReference();
                                            entPmsconfiguration["hil_pmscount"] = new OptionSetValue(optValue);
                                            entPmsconfiguration["hil_dayofpms"] = 0;
                                            service.Create(entPmsconfiguration);
                                            optValue += 1;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.WarrantyTemplate.PostCreate.LoadPMSConfigurations" + ex.Message);
            }
        }

        public static void InitiateProcessForParts(IOrganizationService service, hil_warrantytemplate _WrtyTplt)
        {
            EntityReference ProductCategory = new EntityReference(Product.EntityLogicalName);
            EntityReference Part = new EntityReference(Product.EntityLogicalName);
            if (_WrtyTplt.Attributes.Contains("hil_product"))
            {
                ProductCategory = (EntityReference)_WrtyTplt.hil_Product;
                QueryByAttribute Query = new QueryByAttribute(hil_servicebom.EntityLogicalName);
                ColumnSet Col = new ColumnSet(true);
                Query.ColumnSet = Col;
                Query.AddAttributeValue("hil_productcategory", ProductCategory.Id);
                EntityCollection Found = new EntityCollection();
                Found = service.RetrieveMultiple(Query);
                foreach(hil_servicebom _SrvBm in Found.Entities)
                {
                    Part = (EntityReference)_SrvBm.hil_Product;
                    CreatePartRecordsinWarrantyTemplate(service, Part.Id, _WrtyTplt.Id);
                }
            }
        }
        public static void CreatePartRecordsinWarrantyTemplate(IOrganizationService service, Guid Part, Guid WrtyTplt)
        {
            hil_part _Prt = new hil_part();
            _Prt.hil_WarrantyTemplateId = new EntityReference(hil_warrantytemplate.EntityLogicalName, WrtyTplt);
            _Prt.hil_PartCode = new EntityReference(Product.EntityLogicalName, Part);
            _Prt.hil_IncludedinWarranty = new OptionSetValue(1);
            service.Create(_Prt);
        }
        public static void InitiateProcessForLabors(IOrganizationService service, hil_warrantytemplate _WrtyTplt)
        {
            EntityReference ProductCategory = new EntityReference(Product.EntityLogicalName);
            Guid Labor = new Guid();
            if (_WrtyTplt.Attributes.Contains("hil_product"))
            {
                ProductCategory = (EntityReference)_WrtyTplt.hil_Product;
                QueryByAttribute Query = new QueryByAttribute(Product.EntityLogicalName);
                ColumnSet Col = new ColumnSet(true);
                Query.ColumnSet = Col;
                Query.AddAttributeValue("hil_materialgroup", ProductCategory.Id);
                Query.AddAttributeValue("producttypecode", 4);
                EntityCollection Found = new EntityCollection();
                Found = service.RetrieveMultiple(Query);
                foreach (Product _Pdt in Found.Entities)
                {
                    Labor = (Guid)_Pdt.ProductId;
                    CreateLaborRecordsinWarrantyTemplate(service, Labor, _WrtyTplt.Id);
                }
            }
        }
        public static void CreateLaborRecordsinWarrantyTemplate(IOrganizationService service, Guid Labor, Guid WrtyTplt)
        {
            hil_labor _Lbr = new hil_labor();
            _Lbr.hil_WarrantyTemplateId = new EntityReference(hil_warrantytemplate.EntityLogicalName, WrtyTplt);
            _Lbr.hil_Labor = new EntityReference(Product.EntityLogicalName, Labor);
            _Lbr.hil_includedinwarranty = new OptionSetValue(1);
            service.Create(_Lbr);
        }
    }
}