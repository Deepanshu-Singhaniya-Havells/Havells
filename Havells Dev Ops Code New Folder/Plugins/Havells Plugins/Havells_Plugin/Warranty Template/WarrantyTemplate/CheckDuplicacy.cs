using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin.Warranty_Template
{
    public class CheckDuplicacy : IPlugin
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
                    hil_warrantytemplate _WrtyTplt = entity.ToEntity<hil_warrantytemplate>();

                    if (_WrtyTplt.Attributes.Contains("hil_validto") && _WrtyTplt.Attributes.Contains("hil_validfrom") &&
                        _WrtyTplt.hil_Product != null && _WrtyTplt.hil_type != null && _WrtyTplt.Attributes.Contains("hil_applicableon"))
                    {
                        DateTime iValidFrom = (DateTime)_WrtyTplt["hil_validfrom"];
                        DateTime iValidTo = (DateTime)_WrtyTplt["hil_validto"];
                        Guid iSubCategory = _WrtyTplt.hil_Product.Id;
                        int iType = _WrtyTplt.hil_type.Value;
                        int iApplicableon = _WrtyTplt.GetAttributeValue<OptionSetValue>("hil_applicableon").Value;
                        CheckIfExists(iValidFrom, iValidTo, iSubCategory, iType, iApplicableon,service);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            #endregion
        }
        #region Check Warranty Template Duplicate
        public static void CheckIfExists(DateTime iValidFrom, DateTime iValidTo, Guid iSubCategory, int iType, int iApplicableon, IOrganizationService service)
        {
            QueryExpression Query = new QueryExpression(hil_warrantytemplate.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("hil_validto", "hil_validfrom");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_product", ConditionOperator.Equal, iSubCategory);
            Query.Criteria.AddCondition("hil_type", ConditionOperator.Equal, iType);
            Query.Criteria.AddCondition("hil_applicableon", ConditionOperator.Equal, iApplicableon);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                foreach (hil_warrantytemplate iTemp in Found.Entities)
                {
                    if (iTemp.Attributes.Contains("hil_validto") && iTemp.Attributes.Contains("hil_validfrom"))
                    {
                        DateTime uValidTo = (DateTime)iTemp["hil_validto"];
                        DateTime uValidFrom = (DateTime)iTemp["hil_validfrom"];
                        if ((uValidTo > iValidFrom && uValidTo < iValidTo))
                        {
                            hil_warrantytemplate uTemp = new hil_warrantytemplate();
                            uTemp.Id = iTemp.Id;
                            uTemp["hil_validto"] = iValidFrom.AddDays(-1);
                            service.Update(uTemp);
                        }
                    }
                }
            }
        }
        #endregion
    }
}