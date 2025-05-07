using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;


namespace BlankDistrictUpdateInJob
{
    public class Program
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
                UpdateBlankDistrictInJob();
            }
        }

        static void UpdateBlankDistrictInJob()
        {
            try
            {
                Guid _currentRecId = Guid.Empty;
                int i = 0;
                while (true)
                {
                    //string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    //<entity name='msdyn_workorder'>
                    //<attribute name='msdyn_name' />
                    //<attribute name='msdyn_workorderid' />
                    //<attribute name='hil_district' />
                    //<attribute name='hil_city' />
                    //<attribute name='hil_pincode' />
                    //<order attribute='createdon' descending='false' />
                    //<filter type='and'>
                    //<filter type='or'>
                    //    <condition attribute='hil_district' operator='null' />
                    //    <condition attribute='hil_city' operator='null' />
                    //    <condition attribute='hil_pincode' operator='null' />
                    //</filter>
                    //<condition attribute='hil_callsubtype' operator='eq' value='{E3129D79-3C0B-E911-A94E-000D3AF06CD4}' />
                    //<condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                    //</filter>
                    //<link-entity name='hil_address' from='hil_addressid' to='hil_address' link-type='inner' alias='aa'>
                    //    <attribute name='hil_district' />
                    //    <attribute name='hil_city' />
                    //    <attribute name='hil_businessgeo' />
                    //    <attribute name='hil_pincode' />
                    //    <link-entity name='hil_businessmapping' from='hil_businessmappingid' to='hil_businessgeo' link-type='inner' alias='ac'>
                    //        <attribute name='hil_district' />
                    //        <attribute name='hil_city' />
                    //        <attribute name='hil_pincode' />
                    //    </link-entity>
                    //</link-entity>
                    //</entity>
                    //</fetch>";

                    string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='msdyn_workorder'>
                    <attribute name='msdyn_name' />
                    <attribute name='msdyn_workorderid' />
                    <attribute name='hil_district' />
                    <attribute name='hil_city' />
                    <attribute name='hil_pincode' />
                    <order attribute='createdon' descending='false' />
                    <filter type='and'>
                    <filter type='or'>
                        <condition attribute='hil_district' operator='null' />
                        <condition attribute='hil_city' operator='null' />
                        <condition attribute='hil_pincode' operator='null' />
                    </filter>
                    <condition attribute='hil_callsubtype' operator='eq' value='{E3129D79-3C0B-E911-A94E-000D3AF06CD4}' />
                    <condition attribute='msdyn_substatus' operator='eq' value='{1727FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                    </filter>
                    <link-entity name='hil_address' from='hil_addressid' to='hil_address' link-type='inner' alias='aa'>
                        <attribute name='hil_district' />
                        <attribute name='hil_city' />
                        <attribute name='hil_businessgeo' />
                        <attribute name='hil_pincode' />
                        <link-entity name='hil_businessmapping' from='hil_businessmappingid' to='hil_businessgeo' link-type='inner' alias='ac'>
                            <attribute name='hil_district' />
                            <attribute name='hil_city' />
                            <attribute name='hil_pincode' />
                        </link-entity>
                    </link-entity>
                    </entity>
                    </fetch>";

                    EntityCollection EntityList = _service.RetrieveMultiple(new FetchExpression(fetchXML));
                    if (EntityList.Entities.Count == 0) { break; }
                    Entity entUpdate = null;
                    EntityReference erDistrictVal = null;
                    EntityReference erCityVal = null;
                    EntityReference erPinCodeVal = null;
                    int I = 1;
                    foreach (Entity entity in EntityList.Entities)
                    {
                        entUpdate = new Entity("msdyn_workorder");
                        entUpdate.Id = entity.Id;
                        if (!entity.Attributes.Contains("hil_district"))
                        {
                            if (entity.Attributes.Contains("aa.hil_district"))
                            {
                                erDistrictVal = ((EntityReference)entity.GetAttributeValue<AliasedValue>("aa.hil_district").Value);
                            }
                            else if (entity.Attributes.Contains("ac.hil_district"))
                            {
                                erDistrictVal = ((EntityReference)entity.GetAttributeValue<AliasedValue>("ac.hil_district").Value);
                            }
                        }
                        if (!entity.Attributes.Contains("hil_city"))
                        {
                            if (entity.Attributes.Contains("aa.hil_city"))
                            {
                                erCityVal = ((EntityReference)entity.GetAttributeValue<AliasedValue>("aa.hil_city").Value);
                            }
                            else if (entity.Attributes.Contains("ac.hil_city"))
                            {
                                erCityVal = ((EntityReference)entity.GetAttributeValue<AliasedValue>("ac.hil_city").Value);
                            }
                        }
                        if (!entity.Attributes.Contains("hil_pincode"))
                        {
                            if (entity.Attributes.Contains("aa.hil_pincode"))
                            {
                                erPinCodeVal = ((EntityReference)entity.GetAttributeValue<AliasedValue>("aa.hil_pincode").Value);
                            }
                            else if (entity.Attributes.Contains("ac.hil_pincode"))
                            {
                                erPinCodeVal = ((EntityReference)entity.GetAttributeValue<AliasedValue>("ac.hil_pincode").Value);
                            }
                        }
                        bool flg = false;

                        if (erDistrictVal != null)
                        {
                            entUpdate["hil_district"] = erDistrictVal;
                            entUpdate["hil_districttext"] = erDistrictVal.Name;
                            flg = true;
                        }
                        if (erCityVal != null)
                        {
                            entUpdate["hil_city"] = erCityVal;
                            entUpdate["hil_citytext"] = erCityVal.Name;
                            flg = true;
                        }
                        if (erPinCodeVal != null)
                        {
                            entUpdate["hil_pincode"] = erPinCodeVal;
                            flg = true;
                        }
                        if (flg)
                        {
                            _service.Update(entUpdate);
                        }
                        i += 1;
                        Console.WriteLine("Job # " + entity.GetAttributeValue<string>("msdyn_name") + " : " +  i.ToString() + "/" + EntityList.Entities.Count.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
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
