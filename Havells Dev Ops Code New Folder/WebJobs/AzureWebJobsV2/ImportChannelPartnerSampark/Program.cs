using Microsoft.Office.Interop.Excel;
using Microsoft.WindowsAzure.Storage.Auth;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportChannelPartnerSampark
{
    class ImportChannelPartnerSampark
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
                ChannelPartnerDataUpdatesFromSampark();
                //ClosePendingWorkDoneJobs();
            }
        }

        static void ChannelPartnerDataUpdatesFromSamparkExcel()
        {

            string strFileName = "C:\\Logs\\ChannelPartnerDataUpdatesFromSamparkExcel\\logs.txt";
            string filePath = @"C:\Kuldeep khare\RetailerMaster1.xlsx";
            string conn = string.Empty;

            Application excelApp = new Application();
            if (excelApp != null)
            {
                Workbook excelWorkbook = excelApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet excelWorksheet = (Worksheet)excelWorkbook.Sheets[1];

                Range excelRange = excelWorksheet.UsedRange;
                Range range;
                string _mobileNumber, _userType, _userSubType, _customerCode, _fullName, _businessName, _businessAddress, _businessAddress2, _pincode, _emailid, _mdmCode;
                QueryExpression Query1;
                EntityCollection entcoll, entcoll1, entcoll2;
                Entity jobEntity;
                string _fetchXML;

                for (int i = 1001; i <= excelRange.Rows.Count; i++)
                {
                    try
                    {
                        Guid _accountId = Guid.Empty;

                        WriteLogFile.WriteLog("Processing Row# " + i.ToString(), strFileName);
                        range = (excelWorksheet.Cells[i, 1] as Range);
                        _mobileNumber = range.Value.ToString();

                        range = (excelWorksheet.Cells[i, 2] as Range);
                        _userType = range.Value.ToString();

                        range = (excelWorksheet.Cells[i, 3] as Range);
                        _userSubType = range.Value != null ? range.Value.ToString() : "";

                        range = (excelWorksheet.Cells[i, 4] as Range);
                        _customerCode = range.Value != null ? range.Value.ToString() : "";

                        range = (excelWorksheet.Cells[i, 5] as Range);
                        _fullName = range.Value != null ? range.Value.ToString() : "";

                        range = (excelWorksheet.Cells[i, 6] as Range);
                        _businessName = range.Value != null ? range.Value.ToString() : "";

                        range = (excelWorksheet.Cells[i, 7] as Range);
                        _businessAddress = range.Value != null ? range.Value.ToString() : "";

                        range = (excelWorksheet.Cells[i, 8] as Range);
                        _businessAddress2 = range.Value != null ? range.Value.ToString() : "";

                        range = (excelWorksheet.Cells[i, 9] as Range);
                        _pincode = range.Value != null ? range.Value.ToString() : "";

                        range = (excelWorksheet.Cells[i, 10] as Range);
                        _emailid = range.Value != null ? range.Value.ToString() : "";

                        range = (excelWorksheet.Cells[i, 11] as Range);
                        _mdmCode = range.Value != null ? range.Value.ToString() : "";

                        WriteLogFile.WriteLog("Row Data " + _mobileNumber + "/" + _userType + "/" + _userSubType + "/" + _customerCode + "/" + _fullName + "/" + _businessName + "/" + _businessAddress + "/" + _businessAddress2 + "/" + _pincode + "/" + _emailid + "/" + _mdmCode, strFileName);

                        Query1 = new QueryExpression("account");
                        Query1.ColumnSet = new ColumnSet("hil_vendorcode");
                        Query1.Criteria = new FilterExpression(LogicalOperator.And);
                        Query1.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, _customerCode);
                        Query1.AddOrder("createdon", OrderType.Descending);
                        Query1.TopCount = 1;
                        entcoll = _service.RetrieveMultiple(Query1);
                        if (entcoll.Entities.Count > 0)
                        {
                            if (entcoll.Entities[0].Contains("hil_vendorcode"))
                            {
                                WriteLogFile.WriteLog(i.ToString() + "/" + excelRange.Rows.Count.ToString() + " Channel Partner is already exist.", strFileName);
                                //_entUpdate = new Entity("hil_samparkchannelpartner", ent.Id);
                                //_entUpdate["hil_rowstatus"] = true;
                                //_service.Update(_entUpdate);
                                WriteLogFile.WriteLog("},", strFileName);
                                continue;
                            }
                            _accountId = entcoll.Entities[0].Id;
                        }
                        else
                        {
                            Query1 = new QueryExpression("account");
                            Query1.ColumnSet = new ColumnSet("hil_vendorcode");
                            Query1.Criteria = new FilterExpression(LogicalOperator.And);
                            Query1.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, _mdmCode);
                            Query1.AddOrder("createdon", OrderType.Descending);
                            Query1.TopCount = 1;
                            entcoll = _service.RetrieveMultiple(Query1);
                            if (entcoll.Entities.Count > 0)
                            {
                                if (entcoll.Entities[0].Contains("hil_vendorcode"))
                                {
                                    WriteLogFile.WriteLog(i.ToString() + "/" + excelRange.Rows.Count.ToString() + " Channel Partner is already exist.", strFileName);

                                    WriteLogFile.WriteLog("},", strFileName); 
                                    continue;
                                }
                                _accountId = entcoll.Entities[0].Id;
                            }
                        }

                        if (_accountId != Guid.Empty)
                        {
                            jobEntity = new Entity("account", _accountId);
                            jobEntity["name"] = string.IsNullOrEmpty(_businessName) || string.IsNullOrWhiteSpace(_businessName) ? _fullName : _businessName;
                            jobEntity["emailaddress1"] = _emailid;
                            jobEntity["hil_vendorcode"] = _mdmCode;
                            if (!string.IsNullOrEmpty(_userSubType) && !string.IsNullOrWhiteSpace(_userSubType))
                            {
                                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='hil_usertype'>
                                    <attribute name='hil_usertypeid' />
                                    <attribute name='hil_parentusertype' />
                                    <filter type='and'>
                                      <condition attribute='hil_parentusertypename' operator='like' value='%{_userType}%' />
                                      <condition attribute='hil_name' operator='like' value='%{_userSubType}%' />
                                    </filter>
                                  </entity>
                                </fetch>";
                                entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                if (entcoll2.Entities.Count > 0)
                                {
                                    jobEntity["hil_usertype"] = entcoll2.Entities[0].GetAttributeValue<EntityReference>("hil_parentusertype");
                                    jobEntity["hil_usersubtype"] = entcoll2.Entities[0].ToEntityReference();
                                }
                                else
                                {
                                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='hil_usertype'>
                                        <attribute name='hil_usertypeid' />
                                        <filter type='and'>
                                          <condition attribute='hil_name' operator='like' value='%{_userType}%' />
                                        </filter>
                                      </entity>
                                    </fetch>";
                                    entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                    if (entcoll2.Entities.Count > 0)
                                    {
                                        jobEntity["hil_usertype"] = entcoll2.Entities[0].ToEntityReference();
                                    }
                                    //
                                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='hil_usertype'>
                                        <attribute name='hil_usertypeid' />
                                        <filter type='and'>
                                          <condition attribute='hil_name' operator='like' value='%{_userSubType}%' />
                                        </filter>
                                      </entity>
                                    </fetch>";
                                    entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                    if (entcoll2.Entities.Count > 0)
                                    {
                                        jobEntity["hil_usersubtype"] = entcoll2.Entities[0].ToEntityReference();
                                    }
                                }
                            }
                            else if (!string.IsNullOrEmpty(_userType) && !string.IsNullOrWhiteSpace(_userType))
                            {
                                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='hil_usertype'>
                                        <attribute name='hil_usertypeid' />
                                        <filter type='and'>
                                          <condition attribute='hil_name' operator='like' value='%{_userType}%' />
                                        </filter>
                                      </entity>
                                    </fetch>";
                                entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                if (entcoll2.Entities.Count > 0)
                                {
                                    jobEntity["hil_usertype"] = entcoll2.Entities[0].ToEntityReference();
                                }
                            }
                            _service.Update(jobEntity);
                            WriteLogFile.WriteLog(i.ToString() + "/" + excelRange.Rows.Count.ToString() + " Updating Channel Partner " + _customerCode + "/" + _mdmCode, strFileName);
                            continue;
                        }
                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_businessmapping'>
                        <attribute name='hil_businessmappingid' />
                        <attribute name='hil_pincode' />
                        <attribute name='hil_salesoffice' />
                        <attribute name='hil_subterritory' />
                        <attribute name='hil_region' />
                        <attribute name='hil_state' />
                        <attribute name='hil_branch' />
                        <attribute name='hil_city' />
                        <attribute name='hil_area' />
                        <attribute name='hil_district' />
                        <attribute name='hil_name' />
                        <attribute name='createdon' />
                        <order attribute='hil_name' descending='false' />
                        <filter type='and'>
                            <condition attribute='statecode' operator='eq' value='0' />
                            <condition attribute='hil_pincodename' operator='like' value='%{_pincode}%' />
                        </filter>
                        </entity>
                        </fetch>";

                        entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entcoll1.Entities.Count > 0)
                        {
                            jobEntity = new Entity("account");

                            if (!string.IsNullOrEmpty(_userSubType) && !string.IsNullOrWhiteSpace(_userSubType))
                            {
                                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='hil_usertype'>
                                    <attribute name='hil_usertypeid' />
                                    <attribute name='hil_parentusertype' />
                                    <filter type='and'>
                                      <condition attribute='hil_parentusertypename' operator='like' value='%{_userType}%' />
                                      <condition attribute='hil_name' operator='like' value='%{_userSubType}%' />
                                    </filter>
                                  </entity>
                                </fetch>";
                                entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                if (entcoll2.Entities.Count > 0)
                                {
                                    jobEntity["hil_usertype"] = entcoll2.Entities[0].GetAttributeValue<EntityReference>("hil_parentusertype");
                                    jobEntity["hil_usersubtype"] = entcoll2.Entities[0].ToEntityReference();
                                }
                                else
                                {
                                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='hil_usertype'>
                                        <attribute name='hil_usertypeid' />
                                        <filter type='and'>
                                          <condition attribute='hil_name' operator='like' value='%{_userType}%' />
                                        </filter>
                                      </entity>
                                    </fetch>";
                                    entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                    if (entcoll2.Entities.Count > 0)
                                    {
                                        jobEntity["hil_usertype"] = entcoll2.Entities[0].ToEntityReference();
                                    }
                                    //
                                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='hil_usertype'>
                                        <attribute name='hil_usertypeid' />
                                        <filter type='and'>
                                          <condition attribute='hil_name' operator='like' value='%{_userSubType}%' />
                                        </filter>
                                      </entity>
                                    </fetch>";
                                    entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                    if (entcoll2.Entities.Count > 0)
                                    {
                                        jobEntity["hil_usersubtype"] = entcoll2.Entities[0].ToEntityReference();
                                    }
                                }
                            }
                            else if (!string.IsNullOrEmpty(_userType) && !string.IsNullOrWhiteSpace(_userType))
                            {
                                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                      <entity name='hil_usertype'>
                                        <attribute name='hil_usertypeid' />
                                        <filter type='and'>
                                          <condition attribute='hil_name' operator='like' value='%{_userType}%' />
                                        </filter>
                                      </entity>
                                    </fetch>";
                                entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                                if (entcoll2.Entities.Count > 0)
                                {
                                    jobEntity["hil_usertype"] = entcoll2.Entities[0].ToEntityReference();
                                }
                            }
                            jobEntity["accountnumber"] = string.IsNullOrEmpty(_customerCode) || string.IsNullOrWhiteSpace(_customerCode) ? _mdmCode : _customerCode; ;
                            jobEntity["hil_vendorcode"] = _mdmCode;
                            jobEntity["name"] = string.IsNullOrEmpty(_businessName) || string.IsNullOrWhiteSpace(_businessName) ? _fullName : _businessName;
                            jobEntity["telephone1"] = _mobileNumber;
                            jobEntity["address1_line1"] = _businessAddress;
                            jobEntity["address1_line2"] = _businessAddress2;

                            jobEntity["hil_pincode"] = entcoll1.Entities[0]["hil_pincode"];
                            jobEntity["hil_state"] = entcoll1.Entities[0]["hil_state"];
                            jobEntity["hil_district"] = entcoll1.Entities[0]["hil_district"];
                            jobEntity["hil_city"] = entcoll1.Entities[0]["hil_city"];
                            jobEntity["hil_branch"] = entcoll1.Entities[0]["hil_branch"];
                            jobEntity["hil_salesoffice"] = entcoll1.Entities[0]["hil_salesoffice"];
                            jobEntity["hil_region"] = entcoll1.Entities[0]["hil_region"];
                            jobEntity["hil_area"] = entcoll1.Entities[0]["hil_area"];
                            jobEntity["hil_subterritory"] = entcoll1.Entities[0]["hil_subterritory"];

                            jobEntity["emailaddress1"] = _emailid;
                            _service.Create(jobEntity);
                        }
                        else
                        {
                            WriteLogFile.WriteLog(i.ToString() + "/" + excelRange.Rows.Count.ToString() + " Pincode does not exist. " + _pincode, strFileName);
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteLogFile.WriteLog("ERROR!!! Row# " + i.ToString() + "/" + excelRange.Rows.Count.ToString() + ex.Message, strFileName);
                    }
                }
            }

        }
        static void ChannelPartnerDataUpdatesFromSampark()
        {

            var _startIndex = ConfigurationManager.AppSettings["StartIndex"].ToString();
            var _endIndex = ConfigurationManager.AppSettings["EndIndex"].ToString();

            string strFileName = "Sampark_logs_" + _startIndex + "_" + _endIndex + ".txt";

            string _mobileNumber, _userType, _userSubType, _customerCode, _fullName, _businessName, _businessAddress, _businessAddress2, _pincode, _emailid, _mdmCode;
            QueryExpression Query1;
            EntityCollection entcoll, entcoll1, entcoll2;
            Entity jobEntity;
            string _fetchXML;
            int pageNo = 1;
            int totalCount = 0;
        loop:
            _fetchXML = $@"<fetch page='" + pageNo + $@"' version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='hil_samparkchannelpartner'>
                    <attribute name='hil_name' />
                    <attribute name='hil_usertype' />
                    <attribute name='hil_usersubtype' />
                    <attribute name='hil_pincode' />
                    <attribute name='modifiedby' />
                    <attribute name='hil_mdmcode' />
                    <attribute name='hil_index' />
                    <attribute name='hil_fullname' />
                    <attribute name='hil_emailid' />
                    <attribute name='hil_customercode' />
                    <attribute name='hil_businessname' />
                    <attribute name='hil_businessaddress2' />
                    <attribute name='hil_businessaddress' />
                    <attribute name='hil_samparkchannelpartnerid' />
                    <order attribute='hil_index' descending='false' />
                    <filter type='and'>
                        <condition attribute='hil_rowstatus' operator='ne' value='1' />
                        <condition attribute='hil_index' operator='ge' value='{_startIndex}' />
                        <condition attribute='hil_index' operator='le' value='{_endIndex}' />
                    </filter>
                    </entity>
                </fetch>";
            int i = 1;
            Entity _entUpdate;
            entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
            #region Action
            //foreach (Entity ent in entcoll.Entities)
            //{

            //    WriteLogFile.WriteLog("{", strFileName);
            //    WriteLogFile.WriteLog("Processing Index# : " + i.ToString(), strFileName);
            //    try
            //    {
            //        Guid _accountId = Guid.Empty;

            //        i = ent.GetAttributeValue<int>("hil_index");

            //        _mobileNumber = ParseValue(ent, "hil_name");

            //        _userType = ParseValue(ent, "hil_usertype");

            //        _userSubType = ParseValue(ent, "hil_usersubtype");

            //        _customerCode = ParseValue(ent, "hil_customercode");

            //        _fullName = ParseValue(ent, "hil_fullname");

            //        _businessName = ParseValue(ent, "hil_businessname");

            //        _businessAddress = ParseValue(ent, "hil_businessaddress");

            //        _businessAddress2 = ParseValue(ent, "hil_businessaddress2");

            //        _pincode = ParseValue(ent, "hil_pincode");

            //        _emailid = ParseValue(ent, "hil_emailid");

            //        _mdmCode = ParseValue(ent, "hil_mdmcode");

            //        WriteLogFile.WriteLog("Row Data : " + _mobileNumber + "/" + _userType + "/" + _userSubType + "/" + _customerCode + "/" + _fullName + "/" + _businessName + "/" + _businessAddress + "/" + _businessAddress2 + "/" + _pincode + "/" + _emailid + "/" + _mdmCode, strFileName);

            //        Query1 = new QueryExpression("account");
            //        Query1.ColumnSet = new ColumnSet("hil_vendorcode");
            //        Query1.Criteria = new FilterExpression(LogicalOperator.And);
            //        Query1.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, _customerCode);
            //        Query1.AddOrder("createdon", OrderType.Descending);
            //        Query1.TopCount = 1;
            //        EntityCollection entcol2 = _service.RetrieveMultiple(Query1);
            //        if (entcol2.Entities.Count > 0)
            //        {
            //            if (entcol2.Entities[0].Contains("hil_vendorcode"))
            //            {
            //                WriteLogFile.WriteLog(i.ToString() + "/" + entcoll.Entities.Count.ToString() + " Channel Partner is already exist.", strFileName);
            //                _entUpdate = new Entity("hil_samparkchannelpartner", ent.Id);
            //                _entUpdate["hil_rowstatus"] = true;
            //                _service.Update(_entUpdate); 
            //                WriteLogFile.WriteLog("},", strFileName);
            //                continue;
            //            }
            //            _accountId = entcol2.Entities[0].Id;
            //        }
            //        else
            //        {
            //            Query1 = new QueryExpression("account");
            //            Query1.ColumnSet = new ColumnSet("hil_vendorcode");
            //            Query1.Criteria = new FilterExpression(LogicalOperator.And);
            //            Query1.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, _mdmCode);
            //            Query1.AddOrder("createdon", OrderType.Descending);
            //            Query1.TopCount = 1;
            //            entcol2 = _service.RetrieveMultiple(Query1);
            //            if (entcol2.Entities.Count > 0)
            //            {
            //                if (entcol2.Entities[0].Contains("hil_vendorcode"))
            //                {
            //                    WriteLogFile.WriteLog(i.ToString() + "/" + entcoll.Entities.Count.ToString() + " Channel Partner is already exist.", strFileName);
            //                    _entUpdate = new Entity("hil_samparkchannelpartner", ent.Id);
            //                    _entUpdate["hil_rowstatus"] = true;
            //                    _service.Update(_entUpdate);
            //                    WriteLogFile.WriteLog("},", strFileName);
            //                    continue;
            //                }
            //                _accountId = entcol2.Entities[0].Id;
            //            }
            //        }

            //        if (_accountId != Guid.Empty)
            //        {
            //            jobEntity = new Entity("account", _accountId);
            //            jobEntity["name"] = string.IsNullOrEmpty(_businessName) || string.IsNullOrWhiteSpace(_businessName) ? _fullName : _businessName;
            //            jobEntity["emailaddress1"] = _emailid;
            //            jobEntity["hil_vendorcode"] = _mdmCode;
            //            if (!string.IsNullOrEmpty(_userSubType) && !string.IsNullOrWhiteSpace(_userSubType))
            //            {
            //                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //                      <entity name='hil_usertype'>
            //                        <attribute name='hil_usertypeid' />
            //                        <attribute name='hil_parentusertype' />
            //                        <filter type='and'>
            //                          <condition attribute='hil_parentusertypename' operator='like' value='%{_userType}%' />
            //                          <condition attribute='hil_name' operator='like' value='%{_userSubType}%' />
            //                        </filter>
            //                      </entity>
            //                    </fetch>";
            //                entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
            //                if (entcoll2.Entities.Count > 0)
            //                {
            //                    jobEntity["hil_usertype"] = entcoll2.Entities[0].GetAttributeValue<EntityReference>("hil_parentusertype");
            //                    jobEntity["hil_usersubtype"] = entcoll2.Entities[0].ToEntityReference();
            //                }
            //                else
            //                {
            //                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //                          <entity name='hil_usertype'>
            //                            <attribute name='hil_usertypeid' />
            //                            <filter type='and'>
            //                              <condition attribute='hil_name' operator='like' value='%{_userType}%' />
            //                            </filter>
            //                          </entity>
            //                        </fetch>";
            //                    entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
            //                    if (entcoll2.Entities.Count > 0)
            //                    {
            //                        jobEntity["hil_usertype"] = entcoll2.Entities[0].ToEntityReference();
            //                    }
            //                    //
            //                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //                          <entity name='hil_usertype'>
            //                            <attribute name='hil_usertypeid' />
            //                            <filter type='and'>
            //                              <condition attribute='hil_name' operator='like' value='%{_userSubType}%' />
            //                            </filter>
            //                          </entity>
            //                        </fetch>";
            //                    entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
            //                    if (entcoll2.Entities.Count > 0)
            //                    {
            //                        jobEntity["hil_usersubtype"] = entcoll2.Entities[0].ToEntityReference();
            //                    }
            //                }
            //            }
            //            else if (!string.IsNullOrEmpty(_userType) && !string.IsNullOrWhiteSpace(_userType))
            //            {
            //                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //                          <entity name='hil_usertype'>
            //                            <attribute name='hil_usertypeid' />
            //                            <filter type='and'>
            //                              <condition attribute='hil_name' operator='like' value='%{_userType}%' />
            //                            </filter>
            //                          </entity>
            //                        </fetch>";
            //                entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
            //                if (entcoll2.Entities.Count > 0)
            //                {
            //                    jobEntity["hil_usertype"] = entcoll2.Entities[0].ToEntityReference();
            //                }
            //            }
            //            _service.Update(jobEntity);

            //            _entUpdate = new Entity("hil_samparkchannelpartner", ent.Id);
            //            _entUpdate["hil_rowstatus"] = true;
            //            _service.Update(_entUpdate);

            //            WriteLogFile.WriteLog(i.ToString() + "/" + entcoll.Entities.Count.ToString() + " Updating Channel Partner " + _customerCode + "/" + _mdmCode, strFileName);

            //            WriteLogFile.WriteLog("},", strFileName);
            //            continue;
            //        }
            //        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //            <entity name='hil_businessmapping'>
            //            <attribute name='hil_businessmappingid' />
            //            <attribute name='hil_pincode' />
            //            <attribute name='hil_salesoffice' />
            //            <attribute name='hil_subterritory' />
            //            <attribute name='hil_region' />
            //            <attribute name='hil_state' />
            //            <attribute name='hil_branch' />
            //            <attribute name='hil_city' />
            //            <attribute name='hil_area' />
            //            <attribute name='hil_district' />
            //            <attribute name='hil_name' />
            //            <attribute name='createdon' />
            //            <order attribute='hil_name' descending='false' />
            //            <filter type='and'>
            //                <condition attribute='statecode' operator='eq' value='0' />
            //                <condition attribute='hil_pincodename' operator='like' value='%{_pincode}%' />
            //            </filter>
            //            </entity>
            //            </fetch>";

            //        entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
            //        if (entcoll1.Entities.Count > 0)
            //        {
            //            jobEntity = new Entity("account");

            //            if (!string.IsNullOrEmpty(_userSubType) && !string.IsNullOrWhiteSpace(_userSubType))
            //            {
            //                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //                      <entity name='hil_usertype'>
            //                        <attribute name='hil_usertypeid' />
            //                        <attribute name='hil_parentusertype' />
            //                        <filter type='and'>
            //                          <condition attribute='hil_parentusertypename' operator='like' value='%{_userType}%' />
            //                          <condition attribute='hil_name' operator='like' value='%{_userSubType}%' />
            //                        </filter>
            //                      </entity>
            //                    </fetch>";
            //                entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
            //                if (entcoll2.Entities.Count > 0)
            //                {
            //                    jobEntity["hil_usertype"] = entcoll2.Entities[0].GetAttributeValue<EntityReference>("hil_parentusertype");
            //                    jobEntity["hil_usersubtype"] = entcoll2.Entities[0].ToEntityReference();
            //                }
            //                else
            //                {
            //                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //                          <entity name='hil_usertype'>
            //                            <attribute name='hil_usertypeid' />
            //                            <filter type='and'>
            //                              <condition attribute='hil_name' operator='like' value='%{_userType}%' />
            //                            </filter>
            //                          </entity>
            //                        </fetch>";
            //                    entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
            //                    if (entcoll2.Entities.Count > 0)
            //                    {
            //                        jobEntity["hil_usertype"] = entcoll2.Entities[0].ToEntityReference();
            //                    }
            //                    //
            //                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //                          <entity name='hil_usertype'>
            //                            <attribute name='hil_usertypeid' />
            //                            <filter type='and'>
            //                              <condition attribute='hil_name' operator='like' value='%{_userSubType}%' />
            //                            </filter>
            //                          </entity>
            //                        </fetch>";
            //                    entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
            //                    if (entcoll2.Entities.Count > 0)
            //                    {
            //                        jobEntity["hil_usersubtype"] = entcoll2.Entities[0].ToEntityReference();
            //                    }
            //                }
            //            }
            //            else if (!string.IsNullOrEmpty(_userType) && !string.IsNullOrWhiteSpace(_userType))
            //            {
            //                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            //                          <entity name='hil_usertype'>
            //                            <attribute name='hil_usertypeid' />
            //                            <filter type='and'>
            //                              <condition attribute='hil_name' operator='like' value='%{_userType}%' />
            //                            </filter>
            //                          </entity>
            //                        </fetch>";
            //                entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
            //                if (entcoll2.Entities.Count > 0)
            //                {
            //                    jobEntity["hil_usertype"] = entcoll2.Entities[0].ToEntityReference();
            //                }
            //            }

            //            jobEntity["accountnumber"] = string.IsNullOrEmpty(_customerCode) || string.IsNullOrWhiteSpace(_customerCode) ? _mdmCode : _customerCode; ;
            //            jobEntity["hil_vendorcode"] = _mdmCode;
            //            jobEntity["name"] = string.IsNullOrEmpty(_businessName) || string.IsNullOrWhiteSpace(_businessName) ? _fullName : _businessName;
            //            jobEntity["telephone1"] = _mobileNumber;
            //            jobEntity["address1_line1"] = _businessAddress;
            //            jobEntity["address1_line2"] = _businessAddress2;

            //            if (entcoll1.Entities[0].Contains("hil_pincode"))
            //                jobEntity["hil_pincode"] = entcoll1.Entities[0]["hil_pincode"];

            //            if (entcoll1.Entities[0].Contains("hil_state"))
            //                jobEntity["hil_state"] = entcoll1.Entities[0]["hil_state"];

            //            if (entcoll1.Entities[0].Contains("hil_district"))
            //                jobEntity["hil_district"] = entcoll1.Entities[0]["hil_district"];

            //            if (entcoll1.Entities[0].Contains("hil_city"))
            //                jobEntity["hil_city"] = entcoll1.Entities[0]["hil_city"];

            //            if (entcoll1.Entities[0].Contains("hil_branch"))
            //                jobEntity["hil_branch"] = entcoll1.Entities[0]["hil_branch"];

            //            if (entcoll1.Entities[0].Contains("hil_salesoffice"))
            //                jobEntity["hil_salesoffice"] = entcoll1.Entities[0]["hil_salesoffice"];

            //            if (entcoll1.Entities[0].Contains("hil_region"))
            //                jobEntity["hil_region"] = entcoll1.Entities[0]["hil_region"];

            //            if (entcoll1.Entities[0].Contains("hil_area"))
            //                jobEntity["hil_area"] = entcoll1.Entities[0]["hil_area"];

            //            if (entcoll1.Entities[0].Contains("hil_subterritory"))
            //                jobEntity["hil_subterritory"] = entcoll1.Entities[0]["hil_subterritory"];

            //            jobEntity["emailaddress1"] = _emailid;
            //            _service.Create(jobEntity);
            //            Entity _ent = new Entity("hil_samparkchannelpartner", ent.Id);
            //            _ent["hil_rowstatus"] = true;
            //            _service.Update(_ent);
            //        }
            //        else
            //        {
            //            throw new Exception( " Pincode does not exist. " + _pincode);
            //        }
            //        //Entity _ent = new Entity("hil_samparkchannelpartner", ent.Id);
            //        //_ent["hil_rowstatus"] = true;
            //        //_service.Update(_ent);
            //    }
            //    catch (Exception ex)
            //    {
            //        WriteLogFile.WriteLog("ERROR!!! Row# " + i.ToString() + "/" + entcoll.Entities.Count.ToString() + ex.Message, strFileName);
            //    }

            //    WriteLogFile.WriteLog("},", strFileName);
            //}
            #endregion
            if (entcoll.Entities.Count > 0)
            {
                totalCount = totalCount + entcoll.Entities.Count;
                WriteLogFile.WriteLog(totalCount.ToString(), strFileName);
                pageNo++;
                goto loop;
            }
            WriteLogFile.WriteLog(totalCount.ToString(), strFileName);

        }

        static string ParseValue(Entity _obj, string _fieldName)
        {
            if (_obj.Contains(_fieldName))
                return _obj.GetAttributeValue<string>(_fieldName);
            else
                return string.Empty;
        }

        static void ClosePendingWorkDoneJobs()
        {
            string strFileName = "C:\\Logs\\ClosePendingWorkDoneJobs\\logs.txt";

            try
            {

                //var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
                //var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
                //string finalString = string.Format(connStr, CrmURL);
                EntityCollection jobsColl = new EntityCollection();

                string fetchquery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1000'>
                    <entity name='msdyn_workorder'>
                    <attribute name='msdyn_name'/>
                    <order attribute='hil_jobclosuredon' descending='true'/>
                    <filter type='and'>
                    <condition attribute='msdyn_substatus' operator='eq' value='2927FA6C-FA0F-E911-A94E-000D3AF060A1'/>
                    <condition attribute='hil_jobclosuredon' operator='on-or-before' value='2023-04-30'/>
                    <condition attribute='hil_callsubtype' operator='ne' value='55A71A52-3C0B-E911-A94E-000D3AF06CD4'/>
                    </filter>
                    </entity>
                    </fetch>";
                EntityCollection jobs;
                Entity entWorkorder;
                int i = 1;
                while (true)
                {
                    jobs = _service.RetrieveMultiple(new FetchExpression(fetchquery));
                    if (jobs.Entities.Count == 0) { break; }
                    foreach (Entity item in jobs.Entities)
                    {
                        try
                        {
                            entWorkorder = new Entity("msdyn_workorder", item.Id);
                            entWorkorder["hil_closeticket"] = true;
                            entWorkorder["hil_kkgcode_sms"] = new OptionSetValue(910590006);
                            entWorkorder["hil_webclosureremarks"] = "Closed from backend as per approval";
                            entWorkorder["hil_closureremarks"] = "Closed from backend as per approval";
                            _service.Update(entWorkorder);
                            WriteLogFile.WriteLog("Processing... " + i++.ToString() + " Jobs : " + item.Attributes["msdyn_name"].ToString(), strFileName);
                        }
                        catch (Exception ex)
                        {
                            WriteLogFile.WriteLog("Processing... " + i++.ToString() + " Error : " + item.Attributes["msdyn_name"].ToString() + "/" + ex.Message, strFileName);
                            //StreamWriter log;
                            //if (!File.Exists(@"C:\Kuldeep khare\Errors\logfile.xls"))
                            //{
                            //    log = new StreamWriter(@"C:\Kuldeep khare\Errors\logfile.xls");
                            //}
                            //else
                            //{
                            //    log = File.AppendText(@"C:\Kuldeep khare\Errors\logfile.xls");
                            //}
                            //log.Write(item.Attributes["msdyn_name"].ToString() + "||" + ex.Message);
                            //log.WriteLine();
                            //log.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLogFile.WriteLog("Error:- " + ex.Message, strFileName);
            }
        }

        #region App Setting Load/CRM Connection
        static IOrganizationService ConnectToCRM(string connectionString)
        {
            string strFileName = "C:\\Logs\\ConnectToCRM\\logs.txt";
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    WriteLogFile.WriteLog(((CrmServiceClient)service).LastCrmException.Message, strFileName);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                WriteLogFile.WriteLog("Error while Creating Conn: " + ex.Message, strFileName);
            }
            return service;

        }
        #endregion
    }
}
