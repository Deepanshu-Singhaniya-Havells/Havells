using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Diagnostics.SymbolStore;
using System.Runtime.Remoting.Contexts;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreCreatePurchaseOrder : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "hil_inventorypurchaseorder" && (context.MessageName.ToUpper() == "CREATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];

                    if (entity.Contains("hil_franchise")) {
                        EntityReference _entRefChannelPartner = entity.GetAttributeValue<EntityReference>("hil_franchise");
                        Entity _entChannelPartner = service.Retrieve(_entRefChannelPartner.LogicalName, _entRefChannelPartner.Id, new ColumnSet("accountnumber"));
                        entity["hil_channelpartnercode"] = _entChannelPartner.GetAttributeValue<string>("accountnumber");
                    }
                    if (entity.Contains("hil_ordertype") && entity.Contains("hil_jobid"))
                    {
                        int orderType = entity.GetAttributeValue<OptionSetValue>("hil_ordertype").Value;
                        if (orderType != 1 && orderType != 4)//not equal emergency and FG Replacement
                        {
                            throw new InvalidPluginExecutionException("Invalid Order Type. You can't create Non Emergency or FG Replacement Order against Job.");
                        }
                    }
                    //Add the logic here to change the Owner and owning unit of the Entity to Franchise's Owner and it's owening unit if Context.loginid doesn't equalt to Franchise's Owner
                    if (entity.Contains("hil_franchise"))
                    {
                        EntityReference _entRefChannelPartner = entity.GetAttributeValue<EntityReference>("hil_franchise");
                        Entity _entChannelPartner = service.Retrieve(_entRefChannelPartner.LogicalName, _entRefChannelPartner.Id, new ColumnSet("ownerid", "owningbusinessunit"));
                        if (_entChannelPartner.Contains("ownerid") && _entChannelPartner.Contains("owningbusinessunit"))
                        {
                            if (_entChannelPartner.GetAttributeValue<EntityReference>("ownerid").Id != context.UserId)
                            {
                                entity["ownerid"] = _entChannelPartner.GetAttributeValue<EntityReference>("ownerid");
                                entity["owningbusinessunit"] = _entChannelPartner.GetAttributeValue<EntityReference>("owningbusinessunit");
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
}