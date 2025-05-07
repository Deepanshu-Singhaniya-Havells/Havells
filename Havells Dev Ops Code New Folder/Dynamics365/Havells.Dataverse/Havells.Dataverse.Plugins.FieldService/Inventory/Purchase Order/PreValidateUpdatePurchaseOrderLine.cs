using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreValidateUpdatePurchaseOrderLine : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "hil_inventorypurchaseorderline" && (context.MessageName.ToUpper() == "UPDATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    Entity entityHeader = service.Retrieve("hil_inventorypurchaseorder", entity.GetAttributeValue<EntityReference>("hil_ponumber").Id, new ColumnSet("hil_postatus","hil_approver"));
                    OptionSetValue _orderStatus = entityHeader.GetAttributeValue<OptionSetValue>("hil_postatus");
                    EntityReference _entRefApprover = null;
                    if (entityHeader.Contains("hil_approver"))
                        _entRefApprover = entityHeader.GetAttributeValue<EntityReference>("hil_approver");

                    if (entity.Contains("hil_partcode"))
                    {
                        if (_orderStatus.Value != 1 && _orderStatus.Value != 3)
                        {
                            throw new InvalidPluginExecutionException("You are not allowed to change the Part Code.");
                        }
                        else if (_orderStatus.Value == 3 && _entRefApprover.Id != context.InitiatingUserId)
                        {
                            throw new InvalidPluginExecutionException("You are not allowed to change the Part Code.");
                        }
                    }

                    if (entity.Contains("hil_orderquantity"))
                    {
                        if (_orderStatus.Value != 1 && _orderStatus.Value != 3)
                        {
                            throw new InvalidPluginExecutionException("You are not allowed to change the Part Quantity.");
                        }
                        else if (_orderStatus.Value == 3 && _entRefApprover.Id != context.InitiatingUserId)
                        {
                            throw new InvalidPluginExecutionException("You are not allowed to change the Part Quantity.");
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
}
