using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace HavellsNewPlugin.HealthIndicator
{
    public class PostUpdateLine : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region Execute Main Region
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService traceservice = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            Entity entity = (Entity)context.InputParameters["Target"];
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                            && context.PrimaryEntityName.ToLower() == "hil_technicianhealthindicator" && context.MessageName.ToUpper() == "UPDATE")
                {
                    if (entity.Contains("hil_vaccinationstatus"))
                    {
                        traceservice.Trace("hil_vaccinationstatus");
                        EntityReference technication = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_technicianname")).GetAttributeValue<EntityReference>("hil_technicianname");
                        String fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='hil_attachment'>
                                            <attribute name='activityid' />
                                            <attribute name='subject' />
                                            <attribute name='createdon' />
                                            <order attribute='subject' descending='false' />
                                            <filter type='and'>
                                              <condition attribute='hil_isdeleted' operator='eq' value='0' />
                                              <condition attribute='hil_regardinguser' operator='eq' uiname='' uitype='systemuser' value='" + technication.Id + @"' />
                                            </filter>
                                          </entity>
                                        </fetch>";
                        EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(fetch));
                        traceservice.Trace("entCol " + entCol.Entities.Count);
                        OptionSetValue _healthVaccination = entity.GetAttributeValue<OptionSetValue>("hil_vaccinationstatus");
                        traceservice.Trace("_healthVaccination " + _healthVaccination.Value);
                        if (_healthVaccination.Value == 910590000)//Dose 1
                        {
                            traceservice.Trace("Dose 1 ");
                            if (entCol.Entities.Count < 1)
                            {
                                traceservice.Trace("Dose 1 ");
                                throw new InvalidPluginExecutionException("Kindly attach your Dose 1 Vaccination Certificate.");
                            }
                        }
                        if (_healthVaccination.Value == 910590001)//Dose 2
                        {
                            traceservice.Trace("Dose 2 ");
                            if (entCol.Entities.Count < 2)
                            {
                                traceservice.Trace("Dose 2 ");
                                throw new InvalidPluginExecutionException("Kindly attach your Final Vaccination Certificate.");
                            }
                        }
                        Entity user = service.Retrieve(technication.LogicalName, technication.Id, new ColumnSet("hil_vaccinationstatus"));
                        if (user.Contains("hil_vaccinationstatus"))
                        {

                            OptionSetValue _usserVaccination = user.GetAttributeValue<OptionSetValue>("hil_vaccinationstatus");
                            if (_usserVaccination.Value == 910590001)
                            {
                                throw new InvalidPluginExecutionException("You are not allowed to change Vaccination Status.");
                            }
                            if (_usserVaccination.Value == 910590000)
                            {
                                if (_healthVaccination.Value != 910590001)
                                    throw new InvalidPluginExecutionException("You are not allowed to change Vaccination Status.");
                            }
                        }
                        Entity newUser = new Entity(user.LogicalName);
                        newUser.Id = user.Id;
                        newUser["hil_vaccinationstatus"] = entity["hil_vaccinationstatus"];
                        service.Update(newUser);
                        // throw new InvalidPluginExecutionException("ssss");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error : " + ex.Message);
            }

        }
    }
}
