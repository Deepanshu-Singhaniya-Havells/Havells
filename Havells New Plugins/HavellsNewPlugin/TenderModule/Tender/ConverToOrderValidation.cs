using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.TenderModule.Tender
{
    public class ConverToOrderValidation : IPlugin
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
                if (context.InputParameters.Contains("Target")
                && context.InputParameters["Target"] is Entity
                && context.PrimaryEntityName.ToLower() == "hil_tender"
                && context.Depth == 1)
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    string _rowIds = string.Empty;
                    if (entity.Contains("hil_tenderstatusnew")) // Tender Status
                    {
                        if (entity.GetAttributeValue<OptionSetValue>("hil_tenderstatusnew").Value == 2) //Converted To Order
                        {
                            entity = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_stakeholder"));
                            if (entity.GetAttributeValue<OptionSetValue>("hil_stakeholder").Value != 1)
                            {
                                throw new InvalidPluginExecutionException("Access is Denied!!"); //If Stake Holder is not Branch
                            }
                            string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_tenderproduct'>
                                    <attribute name='hil_tenderproductid' />
                                    <attribute name='hil_name' />
                                    <attribute name='hil_quantity' />
                                    <attribute name='hil_product' />
                                    <attribute name='createdon' />
                                <order attribute='hil_name' descending='false' />
                                <filter type='and'>
                                    <condition attribute='hil_tenderid' operator='eq' value='{" + entity.Id + @"}' />
                                    <filter type='or'>
                                        <condition attribute='hil_quantity' operator='null' />
                                        <condition attribute='hil_quantity' operator='eq' value='0' />
                                        <condition attribute='hil_product' operator='null' />
                                        <condition attribute='hil_lprsmtr' operator='null' />
                                        <condition attribute='hil_lprsmtr'  operator='eq' value='0' />
                                    </filter>
                                    <condition attribute='statecode' operator='eq' value='0' />
                                </filter>
                                </entity>
                                </fetch>";
                            EntityCollection _entColl = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (_entColl.Entities.Count > 0)
                            {
                                _rowIds = "Qty/Product/LP is misisng at Line# ";
                                foreach (Entity ent in _entColl.Entities)
                                {
                                    _rowIds = _rowIds + ent.GetAttributeValue<string>("hil_name") + ",";
                                }
                                throw new InvalidPluginExecutionException(_rowIds);
                            }
                            else
                            {
                                QueryExpression query = new QueryExpression("hil_tenderproduct");
                                query.ColumnSet = new ColumnSet("hil_tenderproductid", "hil_product", "hil_name", "hil_lprsmtr", "hil_unit");
                                query.Distinct = true;
                                query.Criteria.AddCondition("hil_tenderid", ConditionOperator.Equal, entity.Id);
                                query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                                LinkEntity EntityP = new LinkEntity("hil_tenderproduct", "productpricelevel", "hil_product", "productid", JoinOperator.LeftOuter);
                                EntityP.Columns = new ColumnSet("amount");
                                EntityP.EntityAlias = "prd";
                                query.LinkEntities.Add(EntityP);
                                _entColl = service.RetrieveMultiple(query);
                                tracingService.Trace("_entColl.Entities.Count " + _entColl.Entities.Count);
                                if (_entColl.Entities.Count > 0)
                                {
                                    Decimal _tndProdLP;
                                    Decimal _prodLP;
                                    decimal newLP;
                                    bool _flg = false;
                                    _rowIds = "LP has been changed for Product# ";
                                    foreach (Entity ent in _entColl.Entities)
                                    {
                                        _tndProdLP = 0;
                                        _prodLP = 0;
                                        _tndProdLP = ent.GetAttributeValue<Money>("hil_lprsmtr").Value;
                                        tracingService.Trace("_tndProdLP " + _tndProdLP);
                                        if (ent.Contains("prd.amount"))
                                        {
                                            newLP = ((Money)ent.GetAttributeValue<AliasedValue>("prd.amount").Value).Value;
                                            tracingService.Trace("newLP " + newLP);
                                            _prodLP = Math.Round(newLP, 2, MidpointRounding.AwayFromZero);
                                        }
                                        if (_tndProdLP != _prodLP)
                                        {
                                            tracingService.Trace("_tndProdLP " + _tndProdLP);
                                            tracingService.Trace("_prodLP " + _prodLP);
                                            _rowIds = _rowIds + ent.GetAttributeValue<EntityReference>("hil_product").Name + ",";
                                            _flg = true;
                                        }
                                    }
                                    if (_flg)
                                        throw new InvalidPluginExecutionException(_rowIds);
                                }
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
