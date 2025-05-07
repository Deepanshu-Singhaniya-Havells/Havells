using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace RMCostSheet
{
    public class Action_CreateRMCostSheet : IPlugin
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
            //try
            //{
            if (context.InputParameters.Contains("TenderProductId") && context.InputParameters["TenderProductId"] is string && context.Depth == 1)
            {
                var TenderProductId = context.InputParameters["TenderProductId"].ToString();

                if (string.IsNullOrWhiteSpace(TenderProductId))
                {
                    context.OutputParameters["Status"] = false;
                    context.OutputParameters["Message"] = "Tender Product Id is required.";
                }

                string[] _tenderProductIds = TenderProductId.Split(';');
                string _returnMessage = string.Empty;
                string _fetchXML = string.Empty;
                EntityCollection _entCol = null;
                foreach (string _tenderProdId in _tenderProductIds)
                {
                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='hil_tenderproduct'>
                                <attribute name='hil_tenderproductid' />
                                <filter type='and'>
                                  <condition attribute='hil_tenderproductid' operator='eq' value='{_tenderProdId}' />
                                  <condition attribute='hil_selectproduct' operator='eq' value='0' />
                                </filter>
                              </entity>
                            </fetch>";
                    _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (_entCol.Entities.Count == 0)
                    {
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_rmcostsheet'>
                                <attribute name='hil_rmcostsheetid' />
                                <attribute name='hil_tenderproductid' />
                                <attribute name='hil_name' />
                                <filter type='and'>
                                    <condition attribute='hil_tenderproductid' operator='eq' value='{_tenderProdId}' />
                                    <condition attribute='statecode' operator='eq' value='0' />
                                </filter>
                                </entity>
                                </fetch>";
                        _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (_entCol.Entities.Count > 0)
                        {
                            //context.OutputParameters["Status"] = true;
                            //context.OutputParameters["Message"] = "ALERT!! RM Costsheet is already created for Tender Product: " + _entCol.Entities[0].GetAttributeValue<EntityReference>("hil_tenderproductid").Name;
                            //return;
                            _returnMessage += _entCol.Entities[0].GetAttributeValue<string>("hil_name") + ",";
                        }
                        else
                        {
                            Entity _entTenderProduct = service.Retrieve("hil_tenderproduct", new Guid(_tenderProdId), new ColumnSet("hil_tenderid", "hil_name"));
                            EntityReference _entRefTender = _entTenderProduct.GetAttributeValue<EntityReference>("hil_tenderid");

                            Entity _entTender = service.Retrieve("hil_tender", _entRefTender.Id, new ColumnSet("hil_designteam"));
                            if (!_entTender.Contains("hil_designteam"))
                            {
                                context.OutputParameters["Status"] = false;
                                context.OutputParameters["Message"] = "ERROR!!! Design Team is not set on this Tender.";
                                return;
                            }
                            Entity _entRMCostSheet = new Entity("hil_rmcostsheet");
                            _entRMCostSheet["hil_tenderproductid"] = _entTenderProduct.ToEntityReference();
                            _entRMCostSheet["hil_tenderid"] = _entRefTender;
                            _entRMCostSheet["hil_name"] = "CS_" + _entTenderProduct.GetAttributeValue<string>("hil_name");
                            _entRMCostSheet["ownerid"] = _entTender.GetAttributeValue<EntityReference>("hil_designteam");
                            Guid _rnmCostSheetId = service.Create(_entRMCostSheet);

                            Entity _entUpdateTenderProd = new Entity("hil_tenderproduct", _entTenderProduct.Id);
                            _entUpdateTenderProd["hil_rmcostsheet"] = new EntityReference("hil_rmcostsheet", _rnmCostSheetId);
                            service.Update(_entUpdateTenderProd);
                        }
                    }
                }
                if (_returnMessage != string.Empty)
                {
                    context.OutputParameters["Status"] = true;
                    context.OutputParameters["Message"] = "Success!!! Cost Sheet already created for following Lines: " + _returnMessage.Substring(0, _returnMessage.Length - 1);
                }
                else
                {
                    context.OutputParameters["Status"] = true;
                    context.OutputParameters["Message"] = "Success!!! Cost Sheet has been created.";
                }
            }
            //}
            //catch (Exception ex)
            //{
            //    context.OutputParameters["Status"] = false;
            //    context.OutputParameters["Message"] = "ERROR!!! " + ex.Message;
            //}
        }
    }
}
