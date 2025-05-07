using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;

namespace D365WebJobs
{
    public class WarrantyTemplateDuplicateCheck
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
                string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='hil_warrantytemplate'>
                        <attribute name='hil_name' />
                        <attribute name='hil_warrantyperiod' />
                        <attribute name='hil_type' />
                        <attribute name='hil_product' />
                        <attribute name='hil_warrantytemplateid' />
                        <attribute name='hil_validto' />
                        <attribute name='hil_validfrom' />
                        <attribute name='hil_amcplan' />
                        <order attribute='hil_product' descending='false' />
                        <order attribute='hil_type' descending='false' />
                        <order attribute='hil_validfrom' descending='false' />
                        <filter type='and'>
                          <condition attribute='hil_templatestatus' operator='eq' value='2' />
                          <condition attribute='statecode' operator='eq' value='0' />
                          <condition attribute='hil_name' operator='eq' value='WT-20181230-01218' />
                        </filter>
                      </entity>
                    </fetch>";

                EntityReference _prodsubCatg = null;
                EntityReference _amcPlan= null;
                OptionSetValue _type = null;
                DateTime _x,_y, _x1, _y1;
                string _warrantyTemplate = string.Empty;

                EntityCollection entCol = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                int i = 1;
                foreach (Entity ent in entCol.Entities) {
                    Console.WriteLine($"Processing...{i++.ToString()}/{entCol.Entities.Count.ToString()} Warranty Template Id: {ent.GetAttributeValue<string>("hil_name")}");
                    _prodsubCatg = ent.GetAttributeValue<EntityReference>("hil_product");
                    _amcPlan = ent.Contains("hil_amcplan") ? ent.GetAttributeValue<EntityReference>("hil_amcplan") : null;
                    _type = ent.GetAttributeValue<OptionSetValue>("hil_type");
                    _x = ent.GetAttributeValue<DateTime>("hil_validfrom").AddMinutes(330).Date;
                    _y = ent.GetAttributeValue<DateTime>("hil_validto").AddMinutes(330).Date;
                    _warrantyTemplate = ent.GetAttributeValue<string>("hil_name");

                    string _amcPlanCond = string.Empty;
                    if(_amcPlan != null)
                    {
                        _amcPlanCond = $"<condition attribute='hil_amcplan' operator='eq' value='{_amcPlan.Id}' />";
                    }
                    string _fetchXML1 = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_warrantytemplate'>
                        <attribute name='hil_name' />
                        <attribute name='hil_validto' />
                        <attribute name='hil_validfrom' />
                        <filter type='and'>
                            <condition attribute='hil_product' operator='eq' value='{_prodsubCatg.Id}' />
                            <condition attribute='hil_type' operator='eq' value='{_type.Value}' />
                            <condition attribute='hil_warrantytemplateid' operator='ne' value='{ent.Id}' />
                            <condition attribute='statecode' operator='eq' value='0' />{_amcPlanCond}
                        </filter>
                        </entity>
                        </fetch>";
                    EntityCollection entCol1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML1));
                    bool _foundDuplicate;
                    foreach (Entity ent1 in entCol1.Entities) {
                        _foundDuplicate = false;
                        _x1 = ent1.GetAttributeValue<DateTime>("hil_validfrom").AddMinutes(330).Date;
                        _y1 = ent1.GetAttributeValue<DateTime>("hil_validto").AddMinutes(330).Date;

                        if(_x1 >= _x && _x1 <= _y)
                        {
                            _foundDuplicate = true;
                        }
                        else if(_y1 >=_x && _y1 <= _y)
                        {
                            _foundDuplicate = true;
                        }
                        else if(_x1 <= _x && _y1 >= _y)
                        {
                            _foundDuplicate = true;
                        }

                        if (_foundDuplicate)
                        {
                            //Console.WriteLine($"Found Duplicate for Warranty Template Id: {ent.GetAttributeValue<string>("hil_name")}/{ent1.GetAttributeValue<string>("hil_name")}");
                            //Entity _entUpdate = new Entity(ent1.LogicalName, ent1.Id);
                            //_entUpdate["hil_temp1"] = ent.GetAttributeValue<string>("hil_name");
                            //_service.Update(_entUpdate);
                        }
                    }
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
