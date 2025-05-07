using Havells.Dataverse.Plugins.CommonLibs;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Runtime.Remoting.Services;
using System.Text;

namespace Havells.Dataverse.Plugins.FieldService.Inventory
{
    public class PreCreateRMALine : IPlugin
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
                && context.PrimaryEntityName.ToLower() == "hil_inventoryrmaline" && (context.MessageName.ToUpper() == "CREATE"))
            {
                try
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    #region Data Validations
                    if (!entity.Contains("hil_rma"))
                    {
                        throw new InvalidPluginExecutionException("RMA Number is requird.");
                    }
                    if (!entity.Contains("hil_product"))
                    {
                        throw new InvalidPluginExecutionException("Spare Part is requird.");
                    }
                    if (!entity.Contains("hil_quantity"))
                    {
                        throw new InvalidPluginExecutionException("Spare Part Quantity is requird.");
                    }
                    else {
                        int _quantity = entity.GetAttributeValue<int>("hil_quantity");
                        if (_quantity == 0)
                            throw new InvalidPluginExecutionException("Spare Part Quantity is requird.");
                    }
                    if (entity.Contains("hil_rma"))
                    {
                        EntityReference _rmaHeader = entity.GetAttributeValue<EntityReference>("hil_rma");
                        Entity entHeader = service.Retrieve(_rmaHeader.LogicalName, _rmaHeader.Id, new ColumnSet("hil_inspectionnumber", "hil_returntype"));
                        if (entHeader.Contains("hil_inspectionnumber"))
                        {
                            string _inspectionNum = entHeader.GetAttributeValue<string>("hil_inspectionnumber");
                            if (!string.IsNullOrWhiteSpace(_inspectionNum))
                            {
                                throw new InvalidPluginExecutionException("You can't create RMA Line item as Inspection is already submitted!!");
                            }
                        }
                        if (entHeader.Contains("hil_returntype"))
                        {
                            string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1'>
                            <entity name='hil_inventorysettings'>
                            <attribute name='hil_inventorysettingsid' />
                            <attribute name='hil_amcdefectivermatype' />
                            <attribute name='hil_warrantydefectivermatype' />
                            <attribute name='modifiedon' />
                            <filter type='and'>
                                <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            </entity>
                            </fetch>";
                            EntityCollection _entColSettings = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                            EntityReference _warrantyDefectiveChallan = null;
                            EntityReference _amcDefectiveChallan = null;

                            if (_entColSettings.Entities.Count > 0)
                            {
                                _warrantyDefectiveChallan = _entColSettings.Entities.First().GetAttributeValue<EntityReference>("hil_warrantydefectivermatype");
                                _amcDefectiveChallan = _entColSettings.Entities.First().GetAttributeValue<EntityReference>("hil_amcdefectivermatype");
                                EntityReference _returntype = entHeader.GetAttributeValue<EntityReference>("hil_returntype");
                                if (_returntype.Id == _warrantyDefectiveChallan.Id || _returntype.Id == _amcDefectiveChallan.Id)
                                {
                                    if (!entity.Contains("hil_job"))
                                    {
                                        throw new InvalidPluginExecutionException("Job Id is requird.");
                                    }
                                    if (!entity.Contains("hil_jobproduct"))
                                    {
                                        throw new InvalidPluginExecutionException("Job Product is requird.");
                                    }
                                }
                            }
                        }
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
    }
}
