using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.HealthIndicator
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
                    string fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                              <entity name='account'>
                                                <attribute name='name' />
                                                <attribute name='primarycontactid' />
                                                <attribute name='accountnumber' />
                                                <attribute name='hil_salesoffice' />
                                                <attribute name='hil_branch' />
                                                <attribute name='accountid' />
                                                <order attribute='name' descending='false' />
                                                <filter type='and'>
                                                  <condition attribute='ownerid' operator='eq' uiname='' uitype='systemuser' value='" + owner.Id + @"' />
                                                </filter>
                                              </entity>
                                            </fetch>";
                    EntityCollection acc = service.RetrieveMultiple(new FetchExpression(fetch));
                    tracingService.Trace("Acc Count " + acc.Entities.Count);
                    if (acc.Entities.Count != 1)
                    {
                        throw new InvalidPluginExecutionException("Access Denied");
                    }
                    else
                    {
                        String Date = iSysTime.Day.ToString().PadLeft(2, '0') +
                            iSysTime.Month.ToString().PadLeft(2, '0') +
                            (iSysTime.Year % 100).ToString();
                        string number = acc[0].GetAttributeValue<string>("accountnumber");
                        if (acc[0].Contains("hil_salesoffice") && acc[0]["hil_salesoffice"] != null)
                        {
                            salesOffice = (EntityReference)acc[0]["hil_salesoffice"];
                        }
                        if (acc[0].Contains("hil_branch") && acc[0]["hil_branch"] != null)
                        {
                            Branch = (EntityReference)acc[0]["hil_branch"];
                        }
                        tracingService.Trace("Acc Number  " + number);
                        String trackerNumber = Date + number;
                        fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
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
