using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HavellsNewPlugin.AMC_OmniChannel
{
    public class PreUpdateValidateReceiptAmount : IPlugin
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
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "invoice" && context.MessageName.ToUpper() == "UPDATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    Entity preImageEntity = ((Entity)context.PreEntityImages["image"]);

                    if (preImageEntity.Contains("hil_salestype"))
                    {
                        if (preImageEntity.GetAttributeValue<OptionSetValue>("hil_salestype").Value == 3)
                            return;
                    }

                    string _fetchXML = string.Empty;
                    decimal _payableAmount = 0;
                    decimal _receiptAmount = 0;
                    decimal _discountAmount = 0;
                    if (entity.Attributes.Contains("hil_receiptamount"))
                    {
                        _receiptAmount = entity.GetAttributeValue<Money>("hil_receiptamount").Value;
                    }
                    if (preImageEntity.Attributes.Contains("totallineitemamount"))
                    {
                        _payableAmount = preImageEntity.GetAttributeValue<Money>("totallineitemamount").Value;
                    }
                    _discountAmount = _payableAmount - _receiptAmount;

                    string _asOn = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString().PadLeft(2, '0') + "-" + DateTime.Now.Day.ToString().PadLeft(2, '0');
                    decimal _stdDiscPer = 0;
                    decimal _spcDiscPer = 0;
                    decimal _stdDiscAmount = 0;
                    decimal _spcDiscAmount = 0;
                    _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_amcdiscountmatrix'>
                        <attribute name='hil_amcdiscountmatrixid' />
                        <attribute name='hil_name' />
                        <attribute name='createdon' />
                        <attribute name='hil_discounttype' />
                        <attribute name='hil_discper' />
                        <order attribute='hil_name' descending='false' />
                        <filter type='and'>
                            <condition attribute='statecode' operator='eq' value='0' />
                            <condition attribute='hil_appliedto' operator='eq' value='{6f60836d-cd64-ed11-9562-6045bdac526a}' />
                            <condition attribute='hil_validfrom' operator='on-or-before' value='" + _asOn + @"' />
                            <condition attribute='hil_validto' operator='on-or-after' value='" + _asOn + @"' />
                        </filter>
                        </entity>
                        </fetch>";

                    EntityCollection entCol1 = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                    if (entCol1.Entities.Count > 0)
                    {
                        foreach (Entity ent in entCol1.Entities)
                        {
                            if (ent.GetAttributeValue<OptionSetValue>("hil_discounttype").Value == 1)
                            {
                                _stdDiscPer = ent.GetAttributeValue<decimal>("hil_discper");
                            }
                            else
                            {
                                _spcDiscPer = ent.GetAttributeValue<decimal>("hil_discper");
                            }
                        }
                        _stdDiscAmount = Math.Round((_payableAmount - (_payableAmount * _stdDiscPer) / 100), 2); //Max Limit (90)
                        _spcDiscAmount = Math.Round(_payableAmount - (_payableAmount * (_stdDiscPer + _spcDiscPer)) / 100, 2); //Min Limit (85)
                        entity["discountamount"] = new Money(_discountAmount);
                        if (_receiptAmount < _spcDiscAmount)
                        {
                            throw new InvalidPluginExecutionException("As per Discount Policy, you are allowed to collect minimum Rs. " + _stdDiscAmount.ToString() + ".");
                        }
                    }
                    else
                    {
                        if (_payableAmount != _receiptAmount)
                        {
                            throw new InvalidPluginExecutionException("No Discount Policy is defined in System !!! Receipt Amount can't be less than Payable Amount.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
}
