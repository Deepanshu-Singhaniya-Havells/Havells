using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.IPGeo
{
    public class BigGeoPostCreate : IPlugin
    {
        private static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == hil_businessmapping.EntityLogicalName && context.MessageName.ToUpper() == "UPDATE")
                {
                    tracingService.Trace("1");
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if(entity.Attributes.Contains("hil_description"))
                    {
                        string Desc = (string)entity["hil_description"];
                        if(Desc == "Resolved")
                        {
                            hil_businessmapping iobj = (hil_businessmapping)service.Retrieve(hil_businessmapping.EntityLogicalName, entity.Id, new ColumnSet("hil_pincode", "hil_area", "hil_city", "hil_district", "hil_state", "hil_stagingcity", "hil_stagingdistrict", "hil_stagingstate", "hil_stagingpin", "hil_stagingsalesoffice", "hil_stagingterritory", "hil_stagingarea", "hil_stagingbranch"));
                            //ResolveLookups(service, iobj);
                            //SetName2(entity, service);
                            SetTextFields(iobj, service);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        public static void SetName2(Entity entity, IOrganizationService service)
        {
            try
            {

                tracingService.Trace("2");
                hil_businessmapping IPGeo = entity.ToEntity<hil_businessmapping>();
                {
                    if (IPGeo.hil_description != "Resolved")
                        return;


                    IPGeo = (hil_businessmapping)service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("hil_pincode", "hil_area", "hil_city", "hil_district", "hil_state", "hil_stagingcity", "hil_stagingdistrict", "hil_stagingstate", "hil_stagingpin", "hil_stagingsalesoffice", "hil_stagingterritory"));

                    tracingService.Trace("3");
                    String sName = String.Empty;
                    String sPinCode = String.Empty;
                    String sArea = String.Empty;
                    String sCity = String.Empty;
                    String sState = String.Empty;
                    String sDistrict = String.Empty;


                    if (IPGeo.hil_pincode != null)
                    { sName += IPGeo.hil_pincode.Name; }
                    if (IPGeo.hil_area != null)
                    { sName += " " + IPGeo.hil_area.Name; }
                    if (IPGeo.hil_city != null)
                    { sName += " " + IPGeo.hil_city.Name; }
                    if (IPGeo.hil_district != null)
                    { sName += " " + IPGeo.hil_district.Name; }
                    if (IPGeo.hil_state != null)
                    { sName += " " + IPGeo.hil_state.Name; }


                    tracingService.Trace("4" + sName);
                    hil_businessmapping upIPGeo = new hil_businessmapping();
                    upIPGeo.hil_businessmappingId = IPGeo.hil_businessmappingId.Value;
                    upIPGeo.hil_name = sName;
                    service.Update(upIPGeo);

                    tracingService.Trace("5");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }

        }
        public static void ResolveLookups(IOrganizationService service, hil_businessmapping iobj)
        {
            //hil_businessmapping _busMap = new hil_businessmapping();
            //_busMap = (hil_businessmapping)service.Retrieve(hil_businessmapping.EntityLogicalName, iobj.hil_businessmappingId.Value, new ColumnSet(false));
            if (iobj.hil_StagingArea != null)
            {
                QueryByAttribute QryA = new QueryByAttribute(hil_area.EntityLogicalName);
                QryA.ColumnSet = new ColumnSet(false);
                QryA.AddAttributeValue("hil_areacode", iobj.hil_StagingArea);
                //QryA.AddAttributeValue("createdon", DateTime.Now);
                EntityCollection FoundA = service.RetrieveMultiple(QryA);
                if (FoundA.Entities.Count > 0)
                {
                    hil_area Area = (hil_area)FoundA.Entities[0];
                    iobj.hil_area = new EntityReference(hil_area.EntityLogicalName, Area.Id);
                }
            }
                    if (iobj.hil_StagingBranch != null)
                    {
                        QueryByAttribute QryA = new QueryByAttribute(hil_branch.EntityLogicalName);
                        QryA.ColumnSet = new ColumnSet(false);
                        QryA.AddAttributeValue("hil_branchcode", iobj.hil_StagingBranch);
                        EntityCollection FoundA = service.RetrieveMultiple(QryA);
                        if (FoundA.Entities.Count > 0)
                        {
                            hil_branch Branch = (hil_branch)FoundA.Entities[0];
                    iobj.hil_branch = new EntityReference(hil_branch.EntityLogicalName, Branch.Id);
                        }
                    }
                    if (iobj.hil_StagingCity != null)
                    {
                //"hil_stagingcity", "hil_stagingdistrict", "hil_stagingstate", "hil_stagingpin", "hil_stagingsalesoffice", "hil_stagingterritory"
                QueryByAttribute QryA = new QueryByAttribute(hil_city.EntityLogicalName);
                        QryA.ColumnSet = new ColumnSet(false);
                        QryA.AddAttributeValue("hil_citycode", iobj.hil_StagingCity);
                        EntityCollection FoundA = service.RetrieveMultiple(QryA);
                        if (FoundA.Entities.Count > 0)
                        {
                            hil_city City = (hil_city)FoundA.Entities[0];
                    iobj.hil_city = new EntityReference(hil_city.EntityLogicalName, City.Id);
                        }
                    }
                    if (iobj.hil_StagingDistrict != null)
                    {
                        QueryByAttribute QryA = new QueryByAttribute(hil_district.EntityLogicalName);
                        QryA.ColumnSet = new ColumnSet(false);
                        QryA.AddAttributeValue("hil_districtcode", iobj.hil_StagingDistrict);
                        EntityCollection FoundA = service.RetrieveMultiple(QryA);
                        if (FoundA.Entities.Count > 0)
                        {
                            hil_district District = (hil_district)FoundA.Entities[0];
                    iobj.hil_district = new EntityReference(hil_district.EntityLogicalName, District.Id);
                        }
                    }
                    if (iobj.hil_StagingState != null)
                    {
                        QueryByAttribute QryA = new QueryByAttribute(hil_state.EntityLogicalName);
                        QryA.ColumnSet = new ColumnSet(false);
                        QryA.AddAttributeValue("hil_statecode", iobj.hil_StagingState);
                        EntityCollection FoundA = service.RetrieveMultiple(QryA);
                        if (FoundA.Entities.Count > 0)
                        {
                            hil_state State = (hil_state)FoundA.Entities[0];
                    iobj.hil_state = new EntityReference(hil_state.EntityLogicalName, State.Id);
                        }
                    }
                    if (iobj.hil_StagingPIN != null)
                    {
                        QueryByAttribute QryA = new QueryByAttribute(hil_pincode.EntityLogicalName);
                        QryA.ColumnSet = new ColumnSet(false);
                        QryA.AddAttributeValue("hil_name", iobj.hil_StagingPIN);
                        EntityCollection FoundA = service.RetrieveMultiple(QryA);
                        if (FoundA.Entities.Count > 0)
                        {
                            hil_pincode PinCode = (hil_pincode)FoundA.Entities[0];
                    iobj.hil_pincode = new EntityReference(hil_pincode.EntityLogicalName, PinCode.Id);
                        }
                    }
                    if (iobj.hil_StagingSalesOffice != null)
                    {
                        QueryByAttribute QryA = new QueryByAttribute(hil_salesoffice.EntityLogicalName);
                        QryA.ColumnSet = new ColumnSet(false);
                        QryA.AddAttributeValue("hil_salesofficecode", iobj.hil_StagingSalesOffice);
                        EntityCollection FoundA = service.RetrieveMultiple(QryA);
                        if (FoundA.Entities.Count > 0)
                        {
                            hil_salesoffice SalesOff = (hil_salesoffice)FoundA.Entities[0];
                    iobj.hil_salesoffice = new EntityReference(hil_salesoffice.EntityLogicalName, SalesOff.Id);
                        }
                    }
                    if (iobj.hil_StagingTerritory != null)
                    {
                        QueryByAttribute QryA = new QueryByAttribute(hil_subterritory.EntityLogicalName);
                        QryA.ColumnSet = new ColumnSet(false);
                        QryA.AddAttributeValue("hil_subterritorycode", iobj.hil_StagingTerritory);
                        EntityCollection FoundA = service.RetrieveMultiple(QryA);
                        if (FoundA.Entities.Count > 0)
                        {
                            hil_subterritory Territory = (hil_subterritory)FoundA.Entities[0];
                    iobj.hil_subterritory = new EntityReference(hil_subterritory.EntityLogicalName, Territory.Id);
                        }
                    }
            //iobj.hil_description = "resolve";
                    service.Update(iobj);
        }
        public static void SetTextFields(hil_businessmapping _enBus1, IOrganizationService service)
        {
            hil_businessmapping _enBus = (hil_businessmapping)service.Retrieve(hil_businessmapping.EntityLogicalName, _enBus1.Id, new ColumnSet("hil_area", "hil_state", "hil_city", "hil_district", "hil_region"));
            if(_enBus.hil_area != null)
            {
                _enBus["hil_areatext"] = (string)_enBus.hil_area.Name;
            }
            if(_enBus.hil_state != null)
            {
                _enBus["hil_statetext"] = (string)_enBus.hil_state.Name;
            }
            if(_enBus.hil_city != null)
            {
                _enBus["hil_citytext"] = (string)_enBus.hil_city.Name;
            }
            if(_enBus.hil_district != null)
            {
                _enBus["hil_districttext"] = (string)_enBus.hil_district.Name;
            }
            if(_enBus.hil_region != null)
            {
                _enBus["hil_regiontext"] = (string)_enBus.hil_region.Name;
            }
            service.Update(_enBus);
        }
    }
}
