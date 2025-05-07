using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace AutoAssignPendingwithCCJobs
{
    class Program
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion
        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                EntityCollection entcoll;
                //11042213998163

                string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='msdyn_workorder'>
                    <attribute name='msdyn_name' />
                    <attribute name='createdon' />
                    <attribute name='msdyn_substatus' />
                    <attribute name='hil_salesoffice' />
                    <attribute name='hil_branch' />
                    <attribute name='hil_address' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                      <condition attribute='msdyn_substatus' operator='in'>
                        <value uiname='Registered' uitype='msdyn_workordersubstatus'>{2327FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                        <value uiname='Pending for Allocation' uitype='msdyn_workordersubstatus'>{1F27FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                        <value uiname='Work Allocated' uitype='msdyn_workordersubstatus'>{2727FA6C-FA0F-E911-A94E-000D3AF060A1}</value>
                      </condition>
                      <condition attribute='hil_productcategory' operator='not-null' />
                      <condition attribute='hil_callsubtype' operator='not-null' />
                      <condition attribute='hil_productcategory' operator='ne' value='{B96F747D-16FA-E811-A94C-000D3AF0677F}' />
                    </filter>
                    <link-entity name='systemuser' from='systemuserid' to='owninguser' link-type='inner' alias='an'>
                      <filter type='and'>
                        <condition attribute='positionid' operator='eq' value='{1F965242-E80F-E911-A94E-000D3AF06A98}' />
                      </filter>
                    </link-entity>
                    <link-entity name='hil_address' from='hil_addressid' to='hil_address' visible='false' link-type='outer' alias='add'>
                        <attribute name='hil_businessgeo' />
                    </link-entity>
                  </entity> 
                </fetch>";

                entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                int i = 1,j;
                Entity entTemp;
                EntityReference _entAddress = null;
                j = entcoll.Entities.Count;
                string _fetchXMLbizgeo;
                foreach (Entity ent in entcoll.Entities)
                {
                    try
                    {
                        if (ent.Contains("hil_address") && !ent.Contains("hil_salesoffice"))
                        {
                            if (ent.Contains("add.hil_businessgeo")) 
                            {
                                EntityReference erBizgeo = (EntityReference)(ent.GetAttributeValue<AliasedValue>("add.hil_businessgeo").Value);
                                _fetchXMLbizgeo = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                <entity name='hil_businessmapping'>
                                <attribute name='hil_businessmappingid' />
                                <attribute name='hil_pincode' />
                                <filter type='and'>
                                    <condition attribute='hil_businessmappingid' operator='eq' value='{" + erBizgeo.Id + @"}' />
                                    <filter type='or'>
                                    <condition attribute='statuscode' operator='in'>
                                        <value>910590000</value>
                                        <value>2</value>
                                    </condition>
                                    <condition attribute='statecode' operator='eq' value='1' />
                                    </filter>
                                </filter>
                                </entity>
                                </fetch>";
                                EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(_fetchXMLbizgeo));
                                if (EntityList.Entities.Count > 0)
                                {
                                    _fetchXMLbizgeo = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                    <entity name='hil_businessmapping'>
                                    <attribute name='hil_businessmappingid' />
                                    <order attribute='hil_name' descending='false' />
                                    <filter type='and'>
                                        <condition attribute='statecode' operator='eq' value='0' />
                                        <condition attribute='hil_pincode' operator='eq' value='{" + EntityList.Entities[0].GetAttributeValue<EntityReference>("hil_pincode").Id + @"}' />
                                        <condition attribute='statuscode' operator='eq' value='1' />
                                    </filter>
                                    </entity>
                                    </fetch>";
                                    EntityList = _service.RetrieveMultiple(new FetchExpression(_fetchXMLbizgeo));
                                    if (EntityList.Entities.Count > 0)
                                    {
                                        _entAddress = ent.GetAttributeValue<EntityReference>("hil_address");
                                        entTemp = new Entity("hil_address", _entAddress.Id);
                                        entTemp["hil_businessgeo"] = EntityList.Entities[0].ToEntityReference();
                                        _service.Update(entTemp);
                                    }
                                }
                                else {
                                    _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='hil_address'>
                                        <attribute name='hil_state' />
                                        <attribute name='hil_salesoffice' />
                                        <attribute name='hil_region' />
                                        <attribute name='hil_pincode' />
                                        <attribute name='hil_district' />
                                        <attribute name='hil_city' />
                                        <attribute name='hil_branch' />
                                        <attribute name='hil_area' />
                                        <filter type='and'>
                                          <condition attribute='hil_addressid' operator='eq' value='{"+ ent.GetAttributeValue<EntityReference>("hil_address").Id + @"}' />
                                        </filter>
                                      </entity>
                                    </fetch>";
                                    EntityCollection EntityList1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                    if (EntityList1.Entities.Count > 0)
                                    {
                                        entTemp = new Entity("msdyn_workorder", ent.Id);
                                        if (EntityList1.Entities[0].Contains("hil_state"))
                                            entTemp["hil_state"] = EntityList1.Entities[0].GetAttributeValue<EntityReference>("hil_state");
                                        if (EntityList1.Entities[0].Contains("hil_salesoffice"))
                                            entTemp["hil_salesoffice"] = EntityList1.Entities[0].GetAttributeValue<EntityReference>("hil_salesoffice");
                                        if (EntityList1.Entities[0].Contains("hil_region"))
                                            entTemp["hil_region"] = EntityList1.Entities[0].GetAttributeValue<EntityReference>("hil_region");
                                        if (EntityList1.Entities[0].Contains("hil_pincode"))
                                            entTemp["hil_pincode"] = EntityList1.Entities[0].GetAttributeValue<EntityReference>("hil_pincode");
                                        if (EntityList1.Entities[0].Contains("hil_city"))
                                            entTemp["hil_city"] = EntityList1.Entities[0].GetAttributeValue<EntityReference>("hil_city");
                                        if (EntityList1.Entities[0].Contains("hil_branch"))
                                            entTemp["hil_branch"] = EntityList1.Entities[0].GetAttributeValue<EntityReference>("hil_branch");
                                        if (EntityList1.Entities[0].Contains("hil_area"))
                                            entTemp["hil_area"] = EntityList1.Entities[0].GetAttributeValue<EntityReference>("hil_area");
                                        if (EntityList1.Entities[0].Contains("hil_district"))
                                            entTemp["hil_district"] = EntityList1.Entities[0].GetAttributeValue<EntityReference>("hil_district");
                                        _service.Update(entTemp);
                                    }
                                }
                            }
                            _entAddress = ent.GetAttributeValue<EntityReference>("hil_address");
                            entTemp = new Entity("msdyn_workorder", ent.Id);
                            entTemp["hil_address"] = null;
                            _service.Update(entTemp);

                            entTemp = new Entity("msdyn_workorder", ent.Id);
                            entTemp["hil_address"] = _entAddress;
                            _service.Update(entTemp);
                        }
                        if (ent.GetAttributeValue<EntityReference>("msdyn_substatus").Id == new Guid("2727FA6C-FA0F-E911-A94E-000D3AF060A1"))
                        {
                            entTemp = new Entity("msdyn_workorder", ent.Id);
                            entTemp["msdyn_substatus"] = new EntityReference("msdyn_workordersubstatus", new Guid("1F27FA6C-FA0F-E911-A94E-000D3AF060A1"));
                            _service.Update(entTemp);
                        }
                        entTemp = new Entity("msdyn_workorder", ent.Id);
                        entTemp["hil_automaticassign"] = new OptionSetValue(2);
                        _service.Update(entTemp);
                        entTemp = new Entity("msdyn_workorder", ent.Id);
                        entTemp["hil_automaticassign"] = new OptionSetValue(1);
                        _service.Update(entTemp);
                        Console.WriteLine(ent.GetAttributeValue<string>("msdyn_name") + " : " + i.ToString() + "/" + j.ToString());
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ent.GetAttributeValue<string>("msdyn_name") + "/" + ex.Message + " : " + i.ToString() + "/" + j.ToString());
                    }
                    i += 1;
                }
            }
        }

        #region App Setting Load/CRM Connection
        static IOrganizationService ConnectToCRM(string connectionString)
        {
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;

        }
        #endregion
    }
}
