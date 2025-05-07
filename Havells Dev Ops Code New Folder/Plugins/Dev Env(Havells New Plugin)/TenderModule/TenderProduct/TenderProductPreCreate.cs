using System;
using HavellsNewPlugin.TenderModule.OrderCheckListProduct;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.TenderProduct
{
    public class TenderProductPreCreate : IPlugin
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
            try
            {
                if (
                        context.InputParameters.Contains("Target")
                        && context.InputParameters["Target"] is Entity
                        && context.PrimaryEntityName.ToLower() == "hil_tenderproduct"
                        && context.MessageName.ToUpper() == "CREATE"
                    )
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];

                    EntityReference _entRefDepartment = null;
                    Entity _entTender = service.Retrieve("hil_tender", entity.GetAttributeValue<EntityReference>("hil_tenderid").Id, new ColumnSet("hil_pricelist", "hil_name", "hil_department"));

                    if (_entTender.Contains("hil_department"))
                    {
                        _entRefDepartment = _entTender.GetAttributeValue<EntityReference>("hil_department");
                    }
                    tracingService.Trace("1");
                    QueryExpression query = new QueryExpression("hil_tenderproduct");
                    query.ColumnSet = new ColumnSet("hil_name");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition(new ConditionExpression("hil_tenderid", ConditionOperator.Equal,
                        entity.GetAttributeValue<EntityReference>("hil_tenderid").Id));
                    query.Orders.Add(new OrderExpression("createdon", OrderType.Descending));
                    query.TopCount = 1;
                    EntityCollection entColl = service.RetrieveMultiple(query);

                    int _prodCount = 1;
                    string _tendPrefix = _entTender.GetAttributeValue<string>("hil_name");
                    if (entColl.Entities.Count > 0)
                        _prodCount = int.Parse(entColl[0].GetAttributeValue<string>("hil_name").Split('_')[1]) + 1;

                    entity["hil_name"] = _tendPrefix + "_" + _prodCount.ToString().PadLeft(3, '0');

                    if (entity.Contains("hil_product")) // Excel Import
                    {
                        Guid _productId = entity.GetAttributeValue<EntityReference>("hil_product").Id;
                        string _fetchXML = string.Empty;
                        EntityCollection entCol = null;
                        if (_entTender.Contains("hil_pricelist"))
                        {
                            _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='productpricelevel'>
                                <attribute name='productid' />
                                <attribute name='uomid' />
                                <attribute name='amount' />
                                <attribute name='hil_cogs' />
                                <filter type='and'>
                                  <condition attribute='amount' operator='not-null' />
                                  <condition attribute='pricelevelid' operator='eq' value='{" + _entTender.GetAttributeValue<EntityReference>("hil_pricelist").Id.ToString() + @"}' />
                                  <condition attribute='productid' operator='eq' value='{" + _productId.ToString() + @"}' />
                                </filter>
                                <link-entity name='product' from='productid' to='productid' visible='false' link-type='outer' alias='prd'>
                                  <attribute name='description' />
                                </link-entity>
                              </entity>
                            </fetch>";
                            entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (entCol.Entities.Count > 0)
                            {
                                entity["hil_pricelist"] = new EntityReference("pricelevel", _entTender.GetAttributeValue<EntityReference>("hil_pricelist").Id);
                                if (entCol.Entities[0].Contains("uomid"))
                                    entity["hil_unit"] = entCol.Entities[0].GetAttributeValue<EntityReference>("uomid");
                                if (entCol.Entities[0].Contains("amount"))
                                    entity["hil_lprsmtr"] = new Money(entCol.Entities[0].GetAttributeValue<Money>("amount").Value);
                                if (entCol.Entities[0].Contains("prd.description"))
                                    entity["hil_productdescription"] = entCol.Entities[0].GetAttributeValue<AliasedValue>("prd.description").Value.ToString();
                                if (entCol.Entities[0].Contains("hil_cogs"))
                                    entity["hil_cogs"] = new Money(entCol.Entities[0].GetAttributeValue<Money>("hil_cogs").Value);
                            }
                        }
                        EntityReference _hsnCode;
                        decimal taxValue = OrderCheckListProductPreCreate.getHSNValueBasedOnProduct(service, entity.GetAttributeValue<EntityReference>("hil_product"), out _hsnCode);
                        entity["hil_tax"] = taxValue;
                        entity["hil_hsncode"] = _hsnCode;
                    }
                    else if (entity.Contains("hil_hsncode"))
                    {
                        decimal taxValue = OrderCheckListProductPreCreate.getHSNValueBasedOnHSN(service, entity.GetAttributeValue<EntityReference>("hil_hsncode"));
                        entity["hil_tax"] = taxValue;
                    }

                    #region Update Cost Sheet Expenses on Tender ProductLine
                    string _fetchXMLExp = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='hil_departmentrmcostsetup'>
                                <attribute name='hil_departmentrmcostsetupid' />
                                <attribute name='hil_sdvariableexpenseper' />
                                <attribute name='hil_sdfixedexpenseper' />
                                <attribute name='hil_mfgvariableexpenseper' />
                                <attribute name='hil_mfgfixedexpenseper' />
                                <attribute name='hil_packagingexpenseper' />
                                <order attribute='hil_packagingexpenseper' descending='false' />
                                <filter type='and'>
                                  <condition attribute='hil_department' operator='eq' value='{_entRefDepartment.Id}' />
                                </filter>
                              </entity>
                            </fetch>";
                    EntityCollection _entColDepartment = service.RetrieveMultiple(new FetchExpression(_fetchXMLExp));
                    if (_entColDepartment.Entities.Count > 0)
                    {
                        entity["hil_sdfixedexpenseper"] = _entColDepartment.Entities[0].Contains("hil_sdfixedexpenseper") ? _entColDepartment.Entities[0].GetAttributeValue<decimal>("hil_sdfixedexpenseper") : new decimal(0.00);
                        entity["hil_sdvariableexpenseper"] = _entColDepartment.Entities[0].Contains("hil_sdvariableexpenseper") ? _entColDepartment.Entities[0].GetAttributeValue<decimal>("hil_sdvariableexpenseper") : new decimal(0.00);
                    }
                    #endregion
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Error " + ex.Message);
                throw new InvalidPluginExecutionException("HavellsNewPlugin.TenderModule.TenderProduct.Execute Error " + ex.Message);
            }
        }
    }
}