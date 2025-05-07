using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
//using Microsoft.Xrm.Sdk;
//using Microsoft.Xrm.Sdk.Query;


namespace HavellsNewPlugin.HealthIndicator
{
    public class PostCreate : IPlugin
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
                            && context.PrimaryEntityName.ToLower() == "hil_healthindicatorheader" && context.MessageName.ToUpper() == "CREATE")
                {
                    addTeam(service, entity);

                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error : " + ex.Message);
            }

        }
        public static void addTeam(IOrganizationService service, Entity header)
        {
            try
            {
                //Entity header =service.Retrieve("hil_healthindicatorheader", new Guid("1ae4d19f-8e9f-eb11-b1ac-0022486ecf2e"), new ColumnSet("hil_name", "ownerid"));
                EntityReference owner = header.GetAttributeValue<EntityReference>("ownerid");
                Entity OwnerEntity = service.Retrieve("systemuser", owner.Id, new ColumnSet("businessunitid", "positionid"));
                EntityReference BusinessUnit = OwnerEntity.GetAttributeValue<EntityReference>("businessunitid");
                EntityReference Position = OwnerEntity.GetAttributeValue<EntityReference>("positionid");
                String trckerNumber = header.GetAttributeValue<string>("hil_name");
                if (Position.Name == "Franchise")
                {
                    String fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='systemuser'>
                                    <attribute name='fullname' />
                                    <attribute name='businessunitid' />
                                    <attribute name='title' />
                                    <attribute name='mobilephone' />
                                    <attribute name='positionid' />
                                    <attribute name='hil_vaccinationstatus' />
                                    <attribute name='systemuserid' />
                                    <order attribute='fullname' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='isdisabled' operator='eq' value='0' />
                                      <condition attribute='businessunitid' operator='eq' uiname='' uitype='businessunit' value='" + BusinessUnit.Id.ToString() + @"' />
                                      <condition attribute='positionid' operator='eq' uiname='' uitype='position' value='{0197EA9B-1208-E911-A94D-000D3AF0694E}' />
                                    </filter>
                                  </entity>
                                </fetch>";
                    EntityCollection entcol = service.RetrieveMultiple(new FetchExpression(fetch));
                    int counter = createLine(1, entcol, BusinessUnit, owner, new EntityReference(header.LogicalName, header.Id), trckerNumber, service);
                }
                else if (Position.Name == "BSH")
                {
                    String fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='systemuser'>
                                    <attribute name='fullname' />
                                    <attribute name='businessunitid' />
                                    <attribute name='title' />
                                    <attribute name='mobilephone' />
                                    <attribute name='positionid' />
                                    <attribute name='systemuserid' />
                                    <order attribute='fullname' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='positionid' operator='in'>
                                        <value uiname='DSE' uitype='position'>{7D1ECBAB-1208-E911-A94D-000D3AF0694E}</value>
                                        <value uiname='Franchise Technician' uitype='position'>{0197EA9B-1208-E911-A94D-000D3AF0694E}</value>
                                      </condition>
                                      <condition attribute='isdisabled' operator='eq' value='0' />
                                      <condition attribute='businessunitid' operator='eq' uiname='' uitype='businessunit' value='" + BusinessUnit.Id.ToString() + @"' />
                                    </filter>
                                  </entity>
                                </fetch>";
                    EntityCollection entcol = service.RetrieveMultiple(new FetchExpression(fetch));
                    int counter = createLine(1, entcol, BusinessUnit, owner, new EntityReference(header.LogicalName, header.Id), trckerNumber, service);
                    fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='businessunit'>
                                        <attribute name='name' />
                                        <attribute name='websiteurl' />
                                        <attribute name='parentbusinessunitid' />
                                        <attribute name='businessunitid' />
                                        <order attribute='name' descending='false' />
                                        <filter type='and'>
                                          <condition attribute='parentbusinessunitid' operator='eq' uiname='' uitype='businessunit' value='" + BusinessUnit.Id + @"' />
                                        </filter>
                                      </entity>
                                    </fetch>";
                    entcol = service.RetrieveMultiple(new FetchExpression(fetch));
                    foreach (Entity bu in entcol.Entities)
                    {
                        fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='systemuser'>
                                    <attribute name='fullname' />
                                    <attribute name='businessunitid' />
                                    <attribute name='title' />
                                    <attribute name='mobilephone' />
                                    <attribute name='positionid' />
                                    <attribute name='systemuserid' />
                                    <order attribute='fullname' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='positionid' operator='in'>
                                        <value uiname='Franchise Technician' uitype='position'>{0197EA9B-1208-E911-A94D-000D3AF0694E}</value>
                                      </condition>
                                      <condition attribute='isdisabled' operator='eq' value='0' />
                                      <condition attribute='businessunitid' operator='eq' uiname='' uitype='businessunit' value='" + bu.Id.ToString() + @"' />
                                    </filter>
                                  </entity>
                                </fetch>";
                        EntityCollection entcol1 = service.RetrieveMultiple(new FetchExpression(fetch));
                        counter = createLine(counter, entcol1, new EntityReference(bu.LogicalName, bu.Id), owner, new EntityReference(header.LogicalName, header.Id), trckerNumber, service);
                    }
                }
            }

            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error : " + ex.Message);
            }
        }

        private static int createLine(int i, EntityCollection entcol, EntityReference BusinessUnit, EntityReference owner, EntityReference header, String trckerNumber, IOrganizationService service)
        {
            foreach (Entity user in entcol.Entities)
            {
                try
                {
                    Entity healthLine = new Entity("hil_technicianhealthindicator");
                    healthLine["hil_businessunitname"] = BusinessUnit;
                    healthLine["hil_aspname"] = owner;
                    healthLine["hil_technicianname"] = new EntityReference(user.LogicalName, user.Id);
                    healthLine["hil_technicianmobilenumber"] = user.GetAttributeValue<string>("mobilephone");
                    healthLine["hil_vaccinationstatus"] = user.Contains("hil_vaccinationstatus") ? user["hil_vaccinationstatus"] : null;
                    healthLine["hil_healthindicatorheader"] = new EntityReference(header.LogicalName, header.Id);
                    healthLine["hil_name"] = trckerNumber + "-" + i;
                    service.Create(healthLine);
                    i++;
                }
                catch (Exception ex)
                {
                    // throw new InvalidPluginExecutionException("Error " + ex.Message);
                    throw new InvalidPluginExecutionException("Error " + ex.Message);
                }
            }
            return i;
        }

    }
}
