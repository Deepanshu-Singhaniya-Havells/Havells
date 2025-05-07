using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreUpdatePurchaseReceiptLine : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "hil_inventorypurchaseorderreceiptline" && (context.MessageName.ToUpper() == "UPDATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.Contains("hil_billedquantity"))
                    {
                        throw new InvalidPluginExecutionException("Access Denined! You can't change Billed Quantity on Receipt Line.");
                    }
                    if (entity.Contains("hil_partcode"))
                    {
                        throw new InvalidPluginExecutionException("Access Denined! You can't change Spare Part on Receipt Line.");
                    }
                    if (entity.Contains("hil_ordernumber"))
                    {
                        throw new InvalidPluginExecutionException("Access Denined! You can't change Order Number on Receipt Line.");
                    }
                    if (entity.Contains("hil_purchaseorderline"))
                    {
                        throw new InvalidPluginExecutionException("Access Denined! You can't change Order Line Number on Receipt Line.");
                    }
                    if (entity.Contains("hil_warehouse"))
                    {
                        throw new InvalidPluginExecutionException("Access Denined! You can't change Warehouse on Receipt Line.");
                    }
                    if (entity.Contains("hil_jobid"))
                    {
                        throw new InvalidPluginExecutionException("Access Denined! You can't change Job ID on Receipt Line.");
                    }
                    if (entity.Contains("hil_jobproduct"))
                    {
                        throw new InvalidPluginExecutionException("Access Denined! You can't change Job Product ID on Receipt Line.");
                    }
                    if (entity.Contains("statecode") && entity.GetAttributeValue<OptionSetValue>("statecode").Value == 1)
                    {
                        throw new InvalidPluginExecutionException("Access Denined! You can't change Receipt Line Status.");
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
    }
}
