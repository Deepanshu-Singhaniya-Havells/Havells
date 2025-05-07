using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.HealthIndicator
{
    public class Precreate : IPlugin
    {
        private static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            try
            {
                String sSystemAutoNumber = string.Empty;
                EntityReference salesOffice = null;
                EntityReference Branch = null;
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "hil_healthindicatorheader"
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity _entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("1");
                    DateTime iSysTime = DateTime.Now.AddMinutes(330);
                    EntityReference owner = _entity.GetAttributeValue<EntityReference>("ownerid");
                    tracingService.Trace("OwnerID " + owner.Id);

                    Entity _userEntity = service.Retrieve("systemuser", owner.Id, new ColumnSet("hil_account"));

                    if (!_userEntity.Contains("hil_account"))
                    {
                        throw new InvalidPluginExecutionException("Channel Partner is not maped with user. Please contact administrator for more information.");
                    }
                    else
                    {
                        Entity _accountEntity = service.Retrieve(_userEntity.GetAttributeValue<EntityReference>("hil_account").LogicalName,
                            _userEntity.GetAttributeValue<EntityReference>("hil_account").Id,
                            new ColumnSet("hil_salesoffice", "hil_branch", "accountnumber"));
                        String Date = iSysTime.Day.ToString().PadLeft(2, '0') +
                            iSysTime.Month.ToString().PadLeft(2, '0') +
                            (iSysTime.Year % 100).ToString();
                        string number = _accountEntity.GetAttributeValue<string>("accountnumber");
                        if (_accountEntity.Contains("hil_salesoffice") && _accountEntity["hil_salesoffice"] != null)
                        {
                            salesOffice = (EntityReference)_accountEntity["hil_salesoffice"];
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException("Sales Office is not set on account.");
                        }
                        if (_accountEntity.Contains("hil_branch") && _accountEntity["hil_branch"] != null)
                        {
                            Branch = (EntityReference)_accountEntity["hil_branch"];
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException("Branch is not set on account.");
                        }
                        tracingService.Trace("Acc Number  " + number);
                        String trackerNumber = Date + number;
                        String fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='hil_healthindicatorheader'>
                                            <attribute name='hil_healthindicatorheaderid' />
                                            <attribute name='hil_name' />
                                            <attribute name='createdon' />
                                            <attribute name='hil_branch' />
                                            <attribute name='hil_salesoffice' />
                                            <order attribute='hil_name' descending='false' />
                                            <filter type='and'>
                                              <condition attribute='hil_name' operator='eq' value='" + trackerNumber + @"' />
                                            </filter>
                                          </entity>
                                        </fetch>";
                        EntityCollection header = service.RetrieveMultiple(new FetchExpression(fetch));
                        tracingService.Trace("Header Count " + header.Entities.Count);
                        if (header.Entities.Count == 0)
                        {
                            tracingService.Trace("Tracker Number " + trackerNumber);
                            _entity["hil_name"] = trackerNumber;
                            _entity["hil_branch"] = new EntityReference("hil_branch", Branch.Id); ;
                            _entity["hil_salesoffice"] = new EntityReference("hil_salesoffice", salesOffice.Id);
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException("Duplicate Entry");
                        }
                        //                        throw new InvalidPluginExecutionException("Debugger");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error: " + ex.Message);
            }
        }
    }
}
