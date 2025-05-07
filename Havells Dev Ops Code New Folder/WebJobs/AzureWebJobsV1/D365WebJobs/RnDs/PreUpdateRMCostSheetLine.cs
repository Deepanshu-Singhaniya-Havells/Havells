using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace D365WebJobs
{
    public class PreUpdateRMCostSheetLine
    {
        public void Execute(IOrganizationService service)
        {
            try
            {
                Entity entity = new Entity("hil_rmcostsheetline", new Guid("33290a88-0606-ef11-9f89-6045bd7261cd"));
                    EntityReference _rmCode = null;
                    OptionSetValue _rmType = null;

                    Entity _entRMCostSheetLine = service.Retrieve("hil_rmcostsheetline", entity.Id, new ColumnSet("hil_rmtype", "hil_rmcode"));
                    if (_entRMCostSheetLine != null)
                    {
                        if (_entRMCostSheetLine.Contains("hil_rmtype")) { _rmType = _entRMCostSheetLine.GetAttributeValue<OptionSetValue>("hil_rmtype"); }
                        if (_entRMCostSheetLine.Contains("hil_rmcode")) { _rmCode = _entRMCostSheetLine.GetAttributeValue<EntityReference>("hil_rmcode"); }
                    }
                    if (entity.Contains("hil_rmtype"))
                    {
                        _rmType = entity.GetAttributeValue<OptionSetValue>("hil_rmtype");
                    }
                    if (entity.Contains("hil_rmcode"))
                    {
                        _rmCode = entity.GetAttributeValue<EntityReference>("hil_rmcode");
                    }

                    if (_rmType.Value == 1 && _rmCode == null)
                    {
                        throw new InvalidPluginExecutionException("RM Code is required for Existing RM Type.");
                    }
                    if (_rmType.Value == 2 && _rmCode != null)
                    {
                        throw new InvalidPluginExecutionException("RM Code is Not required for WriteIn RM Type.");
                    }

                    if (entity.Contains("hil_rmcode"))
                    {
                        #region Fetching RM Price and Description 
                        RMInfo _rmInfo = CodeBaseLib.GetRMInfo(service, _rmCode);
                        if (_rmInfo != null)
                        {
                            entity["hil_rmdescription"] = _rmInfo.rmDescription;
                            entity["hil_rate"] = _rmInfo.rmPrice;
                        }
                        #endregion
                    }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("RMCostSheet.PreUpdate_RMCostSheet.Execute: " + ex.Message);
            }
        }

        private void CalculateRMCost(IOrganizationService _service, Guid _rmCostSheetId)
        {
            try
            {
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' aggregate='true'>
                    <entity name='hil_rmcostsheetline'>
                    <attribute name='hil_rmcostsheet' groupby='true' alias='header'/>
                    <attribute name='hil_cost' aggregate='sum' alias='totalRMCost'/>
                    <filter type='and'>
                        <condition attribute='hil_rmcostsheet' operator='eq' value='{_rmCostSheetId}' />
                        <condition attribute='statecode' operator='eq' value='0' />
                    </filter>
                    </entity>
                </fetch>";
                EntityCollection _entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                if (_entCol.Entities.Count > 0)
                {
                    Money _totalExpense = new Money(new decimal(0.00));
                    Money _totalRMCost = (Money)(_entCol.Entities[0].GetAttributeValue<AliasedValue>("totalRMCost").Value);
                    Entity _entRMCostSheetHeader = _service.Retrieve("hil_rmcostsheet", _rmCostSheetId, new ColumnSet("hil_totalcost"));
                    if (_entRMCostSheetHeader.Contains("hil_totalcost")) {
                        _totalExpense = _entRMCostSheetHeader.GetAttributeValue<Money>("hil_totalcost");
                    }

                    Entity _entRMCostSheet = new Entity("hil_rmcostsheet", _rmCostSheetId);
                    _entRMCostSheet["hil_rmcost"] = _totalRMCost;
                    _entRMCostSheet["hil_cogs"] = _totalExpense.Value + _totalRMCost.Value;
                    _service.Update(_entRMCostSheet);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
