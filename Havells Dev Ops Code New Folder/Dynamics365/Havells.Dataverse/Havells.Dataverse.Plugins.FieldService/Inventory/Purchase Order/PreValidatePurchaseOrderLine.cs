using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreValidatePurchaseOrderLine : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService _tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            #endregion


            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                && context.PrimaryEntityName.ToLower() == "hil_inventorypurchaseorderline" && (context.MessageName.ToUpper() == "CREATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (!entity.Contains("hil_partcode"))
                        throw new InvalidPluginExecutionException("Part Code is required.");

                    if (!entity.Contains("hil_orderquantity"))
                        throw new InvalidPluginExecutionException("Part Order Quantity is required.");
                    else
                    {
                        int _ordQty = entity.GetAttributeValue<int>("hil_orderquantity");
                        if (_ordQty == 0)
                            throw new InvalidPluginExecutionException("Order Quantity can't be Zero.");
                    }

                    if (!entity.Contains("hil_ponumber"))
                        throw new InvalidPluginExecutionException("Order Number is required.");


                    Entity entityHeader = service.Retrieve("hil_inventorypurchaseorder", entity.GetAttributeValue<EntityReference>("hil_ponumber").Id, new ColumnSet("hil_postatus"));
                    OptionSetValue _orderStatus = entityHeader.GetAttributeValue<OptionSetValue>("hil_postatus");

                    if (_orderStatus.Value == 3 || _orderStatus.Value == 4) //Approved OR Posted
                    {
                        throw new Exception("Access Denied!!!.\nYou can't add Product in Approved Order.");
                    }

                    //string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    //    <entity name='product'>
                    //    <attribute name='productid' />
                    //    <order attribute='productnumber' descending='false' />
                    //    <filter type='and'>
                    //        <condition attribute='productid' operator='eq' value='{entity.GetAttributeValue<EntityReference>("hil_partcode").Id}' />
                    //    </filter>
                    //    <link-entity name='product' from='productid' to='hil_division' visible='false' link-type='outer' alias='div'>
                    //        <attribute name='hil_brandidentifier' />
                    //    </link-entity>
                    //    </entity>
                    //    </fetch>";

                    //EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));

                    //if (_entCol.Entities.Count > 0)
                    //{
                    //    OptionSetValue _diviionBrand = (OptionSetValue)_entCol.Entities[0].GetAttributeValue<AliasedValue>("div.hil_brandidentifier").Value;
                    //    if(_partBrand.Value!= _diviionBrand.Value)
                    //        throw new InvalidPluginExecutionException("Brand Validation!!!\n Part Code's Brand doesn't match with PO Brand.\n ");
                    //}
                    //else
                    //{
                    //    throw new InvalidPluginExecutionException("Brand is not defined in Part Code Division.");
                    //}
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
    }
}
