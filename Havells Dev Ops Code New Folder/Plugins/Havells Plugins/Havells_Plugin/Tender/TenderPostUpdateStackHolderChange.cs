using Havells_Plugin.SAWActivity;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

namespace Havells_Plugin.Tender
{
    public class TenderPostUpdateStackHolderChange : IPlugin
    {
        static ITracingService tracingService = null;
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
                    tracingService.Trace("1");
                    Entity entity = (Entity)context.InputParameters["Target"];
                    string _rowIds = string.Empty;
                    if (entity.Contains("hil_tenderstatusnew"))
                    {
                        if (entity.GetAttributeValue<OptionSetValue>("hil_tenderstatusnew").Value == 2) //Converted To Order
                        {
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
                                _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='hil_tenderproduct'>
                                    <attribute name='hil_tenderproductid' />
                                    <attribute name='hil_product' />
                                    <attribute name='hil_name' />
                                    <attribute name='hil_lprsmtr' />
                                    <order attribute='hil_name' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='hil_tenderid' operator='eq' value='{" + entity.Id + @"}' />
                                      <condition attribute='statecode' operator='eq' value='0' />
                                    </filter>
                                    <link-entity name='product' from='productid' to='hil_product' visible='false' link-type='outer' alias='prd'>
                                      <attribute name='hil_amount' />
                                    </link-entity>
                                  </entity>
                                </fetch>";
                                _entColl = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                if (_entColl.Entities.Count > 0)
                                {
                                    Decimal _tndProdLP;
                                    Decimal _prodLP;
                                    bool _flg = false;
                                    _rowIds = "LP has been changed for Product# ";
                                    foreach (Entity ent in _entColl.Entities)
                                    {
                                        _tndProdLP = 0;
                                        _prodLP = 0;
                                        _tndProdLP = ent.GetAttributeValue<Money>("hil_lprsmtr").Value;
                                        if (ent.Contains("prd.hil_amount"))
                                        {
                                            _prodLP = Math.Round(((Money)ent.GetAttributeValue<AliasedValue>("prd.hil_amount").Value).Value / 1000, 2);
                                        }
                                        if (_tndProdLP != _prodLP)
                                        {
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

                    if (entity.Contains("hil_refreshlp")) {
                        if (entity.GetAttributeValue<bool>("hil_refreshlp")) //Refresh Token = Yes
                        {
                            EntityCollection _entColl;
                            string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='hil_tenderproduct'>
                                    <attribute name='hil_lprsmtr' />
                                    <filter type='and'>
                                      <condition attribute='hil_tenderid' operator='eq' value='{" + entity.Id + @"}' />
                                      <condition attribute='statecode' operator='eq' value='0' />
                                    </filter>
                                    <link-entity name='product' from='productid' to='hil_product' visible='false' link-type='outer' alias='prd'>
                                      <attribute name='hil_amount' />
                                    </link-entity>
                                  </entity>
                                </fetch>";
                            _entColl = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            if (_entColl.Entities.Count > 0)
                            {
                                Decimal _tndProdLP;
                                Decimal _prodLP;
                                foreach (Entity ent in _entColl.Entities)
                                {
                                    _tndProdLP = 0;
                                    _prodLP = 0;
                                    _tndProdLP = ent.GetAttributeValue<Money>("hil_lprsmtr").Value;
                                    if (ent.Contains("prd.hil_amount"))
                                    {
                                        _prodLP = ((Money)ent.GetAttributeValue<AliasedValue>("prd.hil_amount").Value).Value/1000;
                                    }
                                    if (_tndProdLP != _prodLP)
                                    {
                                        decimal _value = Math.Round(((Money)ent.GetAttributeValue<AliasedValue>("prd.hil_amount").Value).Value / 1000,2);
                                        ent["hil_lprsmtr"] = new Money(_value);
                                        service.Update(ent);
                                    }
                                }
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("HavellsNewPlugin.TenderModule.TenderProductPostUpdate_UpdateFinalOfferCost.Execute Error " + ex.Message);
            }
        }
    }
}
