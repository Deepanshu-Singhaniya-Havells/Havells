using System;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Configuration;

namespace D365WebJobs
{
    public class DataMigration
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
        #region Global Varialble declaration
        static IOrganizationService _service, _serviceDev;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        static int RowCount = 1;
        #endregion

        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            _serviceDev = ConnectToCRM(string.Format(connStr, "https://havellscrmqa.crm8.dynamics.com/"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {

                //CountRows();
                //List<MasterList> _masterList = new List<MasterList>();
                //_masterList.Add(new MasterList() { name = "hil_state", columnexclusion = new string[] { "hil_salesoffice" } });
                //_masterList.Add(new MasterList() { name = "hil_district", columnexclusion = new string[] { "hil_subterritory" } });
                //_masterList.Add(new MasterList() { name = "hil_region", columnexclusion = null });
                //_masterList.Add(new MasterList() { name = "hil_subterritory", columnexclusion = new string[] { "hil_state" } });
                //_masterList.Add(new MasterList() { name = "hil_salesoffice", columnexclusion = null });
                //_masterList.Add(new MasterList() { name = "hil_branch", columnexclusion = new string[] { "hil_region" } });
                //_masterList.Add(new MasterList() { name = "hil_city", columnexclusion = null });
                //Parallel.ForEach(_masterList, master =>
                //{
                //    Console.WriteLine("Migrating... " + master.name);
                //    SyncMasters(master.name, master.columnexclusion);
                //});

                //List<MasterList> _masterList = new List<MasterList>();
                //_masterList.Add(new MasterList() { name = "hil_materialgroup", columnexclusion = null });
                //_masterList.Add(new MasterList() { name = "hil_materialgroup2", columnexclusion = null });
                //_masterList.Add(new MasterList() { name = "hil_materialgroup3", columnexclusion = null });
                //_masterList.Add(new MasterList() { name = "hil_materialgroup4", columnexclusion = null });
                //_masterList.Add(new MasterList() { name = "hil_materialgroup5", columnexclusion = null });

                //_masterList.Add(new MasterList() { name = "account", columnexclusion = null });

                //Parallel.ForEach(_masterList, master =>
                //{
                //    Console.WriteLine("Migrating... " + master.name);
                //    SyncMasters(master.name, master.columnexclusion);
                //});

                var EntityName = ConfigurationManager.AppSettings["EntityName"].ToString();
                string ColumnNames = ConfigurationManager.AppSettings["ColumnNames"].ToString();

                SyncMasters(EntityName, ColumnNames.Split(','));
                ////SyncMasters("hil_area", new string[] { "hil_city" });

                Console.WriteLine("DONE.");
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

        private static void CountRows() {
            EntityCollection entCol = null;
            QueryExpression queryExp = new QueryExpression("account");
            queryExp.ColumnSet = new ColumnSet(true);
            queryExp.PageInfo = new PagingInfo
            {
                PageNumber = 1,
                Count = 5000
            };

            int totalRowCount = 0;
            int i = 1;
            while (true)
            {
                entCol = _serviceDev.RetrieveMultiple(queryExp);
                totalRowCount += entCol.Entities.Count;
                if (!entCol.MoreRecords) break;
                queryExp.PageInfo.PageNumber++;
                queryExp.PageInfo.PagingCookie = entCol.PagingCookie;
                Console.WriteLine("TOTAL LINES: " + totalRowCount.ToString());
            }
            Console.WriteLine("TOTAL LINES: " + totalRowCount.ToString());
        }
        #region Business Geo Mapping
        private static void SyncMasters(string _masterName, string[] _exclusions) {
            EntityCollection entCol = null;
            var DateAfter = ConfigurationManager.AppSettings["DateAfter"].ToString();
            var DateBefore = ConfigurationManager.AppSettings["DateBefore"].ToString();

            //FilterExpression _cond = new FilterExpression(LogicalOperator.And);
            //_cond.AddCondition("accountid", ConditionOperator.Equal, new Guid("b96ca8cb-310b-e911-a94d-000d3af0694e"));

            FilterExpression _cond = new FilterExpression(LogicalOperator.And);
            _cond.AddCondition("createdon", ConditionOperator.OnOrAfter, DateAfter);
            _cond.AddCondition("createdon", ConditionOperator.OnOrBefore, DateBefore);

            QueryExpression queryExp = new QueryExpression(_masterName);
            queryExp.ColumnSet = new ColumnSet(true);
            queryExp.PageInfo = new PagingInfo
            {
                PageNumber = 1,
                Count = 5000
            };
            //queryExp.Criteria = _cond;
            queryExp.AddOrder("createdon", OrderType.Ascending);

            int totalRowCount = 0;
            int i = 1;
            while (true)
            {
                entCol = _service.RetrieveMultiple(queryExp);
                totalRowCount += entCol.Entities.Count;
                Parallel.ForEach(entCol.Entities, entity =>
                {
                    SyncAccountRecord(entity, _exclusions);
                });
                //foreach (Entity ent in entCol.Entities)
                //{
                //    SyncRecord(ent, null);
                //}
                if (!entCol.MoreRecords) break;
                queryExp.PageInfo.PageNumber++;
                queryExp.PageInfo.PagingCookie = entCol.PagingCookie;
                Console.WriteLine("TOTAL LINES: " + totalRowCount.ToString());
            }
        }

        private static void SyncRecord(Entity ent, string[] _exclusions) {
            Entity entAccount = ent;
            if (ent.Attributes.Contains("ownerid")) {
                ent["ownerid"] = new EntityReference("systemuser", new Guid("5190416c-0782-e911-a959-000d3af06a98"));
            }
            Console.WriteLine("Processing..." + ent.LogicalName.ToString() + " CreatedOn: " + ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString() + " RowCount: " + RowCount++.ToString() + " Record Id: " + ent.Id.ToString());
            if (_exclusions != null)
            {
                foreach (string _colName in _exclusions)
                {
                    entAccount[_colName] = null;
                }
            }
            try
            {
                FilterExpression _cond = new FilterExpression(LogicalOperator.And);
                _cond.AddCondition(ent.LogicalName.ToString() + "id", ConditionOperator.Equal, ent.Id);

                QueryExpression qryExp = new QueryExpression()
                {
                    ColumnSet = new ColumnSet(false),
                    EntityName = ent.LogicalName.ToString(),
                    NoLock = true,
                    Criteria = _cond
                };

                EntityCollection entCheck = _serviceDev.RetrieveMultiple(qryExp);
                bool _intOV = false;
                OptionSetValue _stateCodeOV, _statusCodeOV;
                Entity entTemp = null;
                Guid recordId = Guid.Empty;

                if (ent["statuscode"].GetType().ToString() == "System.Int32")
                {
                    _intOV = true;
                    _stateCodeOV = new OptionSetValue(Convert.ToInt32(ent["statecode"].ToString()));
                    _statusCodeOV = new OptionSetValue(Convert.ToInt32(ent["statuscode"].ToString()));
                }
                else
                {
                    _stateCodeOV = (OptionSetValue)ent["statecode"];
                    _statusCodeOV = (OptionSetValue)ent["statuscode"];
                }
                if (entCheck.Entities.Count == 0)
                {
                    entAccount["statecode"] = new OptionSetValue(0);
                    entAccount["statuscode"] = new OptionSetValue(1);
                    recordId = _serviceDev.Create(entAccount);
                }
                else
                {
                    recordId = entAccount.Id;
                }
                if (_intOV || entCheck.Entities.Count > 0)
                {
                    entTemp = new Entity(ent.LogicalName.ToString(), recordId);
                    entTemp["statecode"] = _stateCodeOV;
                    entTemp["statuscode"] = _statusCodeOV;
                    if (entCheck.Entities.Count == 0)
                        _serviceDev.Update(entTemp);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR..." + ent.LogicalName.ToString() + " Id: " + ent.Id.ToString() + " \n " + ex.Message);
            }
        }
        private static void SyncAccountRecord(Entity ent, string[] _exclusions)
        {
            EntityCollection entCheck = null;
            Entity entAccount = new Entity(ent.LogicalName, ent.Id);
            Console.WriteLine("Processing..." + ent.LogicalName.ToString() + " CreatedOn: " + ent.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString() + " RowCount: " + RowCount++.ToString() + " Record Id: " + ent.Id.ToString());
            try
            {
                FilterExpression _cond = new FilterExpression(LogicalOperator.And);
                _cond.AddCondition(ent.LogicalName.ToString() + "id", ConditionOperator.Equal, ent.Id);

                QueryExpression qryExp = new QueryExpression()
                {
                    ColumnSet = new ColumnSet("ownerid"),
                    EntityName = ent.LogicalName.ToString(),
                    NoLock = true,
                    Criteria = _cond
                };

                entCheck = _serviceDev.RetrieveMultiple(qryExp);
                Guid recordId = Guid.Empty;

                entAccount["statecode"] = ent.Contains("statecode") ? ent["statecode"] : null;
                entAccount["statuscode"] = ent.Contains("statuscode") ? ent["statuscode"] : null;
                entAccount["accountnumber"] = ent.Contains("accountnumber") ? ent["accountnumber"] : null;
                entAccount["ownerid"] = ent.Contains("ownerid") ? ent["ownerid"] : null;
                entAccount["hil_claimprocessed"] = ent.Contains("hil_claimprocessed") ? ent["hil_claimprocessed"] : null;
                entAccount["name"] = ent.Contains("name") ? ent["name"] : null;
                entAccount["customertypecode"] = ent.Contains("customertypecode") ? ent["customertypecode"] : null;
                entAccount["telephone1"] = ent.Contains("telephone1") ? ent["telephone1"] : null;
                entAccount["emailaddress1"] = ent.Contains("emailaddress1") ? ent["emailaddress1"] : null;
                entAccount["hil_pan"] = ent.Contains("hil_pan") ? ent["hil_pan"] : null;
                entAccount["address1_postofficebox"] = ent.Contains("address1_postofficebox") ? ent["address1_postofficebox"] : null;
                entAccount["hil_state"] = ent.Contains("hil_state") ? ent["hil_state"] : null;
                entAccount["hil_district"] = ent.Contains("hil_district") ? ent["hil_district"] : null;
                entAccount["hil_branch"] = ent.Contains("hil_branch") ? ent["hil_branch"] : null;
                entAccount["hil_city"] = ent.Contains("hil_city") ? ent["hil_city"] : null;
                entAccount["hil_salesoffice"] = ent.Contains("hil_salesoffice") ? ent["hil_salesoffice"] : null;
                entAccount["hil_branch"] = ent.Contains("hil_branch") ? ent["hil_branch"] : null;
                entAccount["address1_line1"] = ent.Contains("address1_line1") ? ent["address1_line1"] : null;
                entAccount["hil_pincode"] = ent.Contains("hil_pincode") ? ent["hil_pincode"] : null;
                entAccount["hil_region"] = ent.Contains("hil_region") ? ent["hil_region"] : null;
                entAccount["address1_line2"] = ent.Contains("address1_line2") ? ent["address1_line2"] : null;
                entAccount["address1_line3"] = ent.Contains("address1_line3") ? ent["address1_line3"] : null;
                entAccount["hil_area"] = ent.Contains("hil_area") ? ent["hil_area"] : null;
                entAccount["hil_subterritory"] = ent.Contains("hil_subterritory") ? ent["hil_subterritory"] : null;
                entAccount["hil_fulladdress"] = ent.Contains("hil_fulladdress") ? ent["hil_fulladdress"] : null;
                entAccount["hil_inwarrantycustomersapcode"] = ent.Contains("hil_inwarrantycustomersapcode") ? ent["hil_inwarrantycustomersapcode"] : null;
                entAccount["hil_category"] = ent.Contains("hil_category") ? ent["hil_category"] : null;
                entAccount["hil_outwarrantycustomersapcode"] = ent.Contains("hil_outwarrantycustomersapcode") ? ent["hil_outwarrantycustomersapcode"] : null;
                entAccount["hil_vendorcode"] = ent.Contains("hil_vendorcode") ? ent["hil_vendorcode"] : null;
                entAccount["accountclassificationcode"] = ent.Contains("accountclassificationcode") ? ent["accountclassificationcode"] : null;

                if (entCheck.Entities.Count == 0)
                {
                    recordId = _serviceDev.Create(entAccount);
                }
                else {
                    if (entCheck.Entities[0].GetAttributeValue<EntityReference>("ownerid").Id != ((EntityReference)ent["ownerid"]).Id)
                        _serviceDev.Update(entAccount);
                }
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("Privilege")|| ex.Message.Contains("privilege"))
                {
                    entAccount["ownerid"] = new EntityReference("systemuser", new Guid("5190416c-0782-e911-a959-000d3af06a98"));
                    if (entCheck.Entities.Count == 0)
                    {
                        _serviceDev.Create(entAccount);
                    }
                    else
                    {
                        _serviceDev.Update(entAccount);
                    }
                }
                else {
                    Console.WriteLine("ERROR..." + ent.LogicalName.ToString() + " Id: " + ent.Id.ToString() + " \n " + ex.Message);
                }
            }
        }
        #endregion
    }
    public class MasterList {
        public string name { get; set; }
        public string[] columnexclusion { get; set; }
    }
}
