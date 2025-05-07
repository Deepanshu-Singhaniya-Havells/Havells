using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.TenderProduct
{
    public class TenderLinePostUpdateProdType : IPlugin
    {
        static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            //Guid tenderID;
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
                    if (entity.Contains("hil_selectproduct"))
                    {
                        bool _prodType = entity.GetAttributeValue<bool>("hil_selectproduct");
                        if (_prodType) //"Write-In"
                            CreateRmCostSheet(service, entity);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        private void CreateRmCostSheet(IOrganizationService service, Entity entity)
        {
            tracingService.Trace("Starting point");

            Entity _entityTenderProd = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_name", "hil_tenderid", "hil_selectproduct"));

            if (_entityTenderProd != null)
            {
                string _tenderProdId = _entityTenderProd.GetAttributeValue<string>("hil_name");

                QueryExpression query = new QueryExpression("hil_rmcostsheet")
                {
                    ColumnSet = new ColumnSet(false),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };
                query.Criteria.AddCondition("hil_tenderproductid", ConditionOperator.Equal, entity.Id);
                EntityCollection existingRMCostSheets = service.RetrieveMultiple(query);
                if (existingRMCostSheets.Entities.Count == 0)
                {
                    Entity rmCostSheet = new Entity("hil_rmcostsheet")
                    {
                        ["hil_tenderproductid"] = _entityTenderProd.ToEntityReference(),
                        ["hil_name"] = $"CS_{_tenderProdId}",
                        ["hil_tenderid"] = _entityTenderProd.GetAttributeValue<EntityReference>("hil_tenderid")
                    };
                    Guid rmCostSheetId = service.Create(rmCostSheet);

                    Entity updatedTenderProduct = new Entity("hil_tenderproduct", entity.Id)
                    {
                        ["hil_rmcostsheet"] = new EntityReference("hil_rmcostsheet", rmCostSheetId)
                    };
                    service.Update(updatedTenderProduct);
                    CreateRmCostSheetDefaultLine(service, rmCostSheetId);
                }
            }
        }

        private void CreateRmCostSheetDefaultLine(IOrganizationService service, Guid rmCostSheetId)
        {
            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='product'>
                                <attribute name='name'/>
                                <attribute name='productnumber'/>
                                <attribute name='description'/>
                                <attribute name='productid'/>
                                <order attribute='productnumber' descending='false'/>
                                <filter type='and'>
                                <condition attribute='productnumber' operator='eq' value='PKG1001'/>
                                </filter>
                                </entity>
                                </fetch>";
            EntityCollection _ebtCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
            if (_ebtCol.Entities.Count > 0)
            {
                Entity rmcostsheetline = new Entity("hil_rmcostsheetline");
                rmcostsheetline["hil_rmcode"] = new EntityReference("product", _ebtCol.Entities[0].Id);
                rmcostsheetline["hil_rmcostsheet"] = new EntityReference("hil_rmcostsheet", rmCostSheetId);
                rmcostsheetline["hil_rmtype"] = new OptionSetValue(1);//Existing
                rmcostsheetline["hil_quantity"] = new decimal(0.00);
                rmcostsheetline["hil_rmdescription"] = _ebtCol.Entities[0].GetAttributeValue<string>("name");
                tracingService.Trace("Last step");
                service.Create(rmcostsheetline);
            }
        }
    }
}

