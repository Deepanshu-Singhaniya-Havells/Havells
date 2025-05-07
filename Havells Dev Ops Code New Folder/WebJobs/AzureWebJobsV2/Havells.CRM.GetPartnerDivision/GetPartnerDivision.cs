using System;
using Microsoft.Xrm.Sdk;
using System.Net;
using Microsoft.Xrm.Sdk.Query;
using System.Runtime.Serialization.Json;
using System.IO;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Configuration;
using Microsoft.Crm.Sdk.Messages;

namespace Havells.CRM.GetPartnerDivision
{
    public class GetPartnerDivision
    {
        #region PARTNER DIVISION MAPPING
        public static void GetPartnerDivisionMappingFromAPI(IOrganizationService service)
        {

            WebClient webClient = new WebClient();
            string url = GetIntegrationUrl(service, "Partner Division");
            //  string RunDate = DateTime.Now.Date.ToString("yyyyMMdd") + "000000";
            //url = url + "?enquiryDate=" + RunDate;
            url = url + "?enquiryDate=" + getTimeStamp(service);
            var jsonData = webClient.DownloadData(url);
            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(PartnerDivisionRootObject));
            PartnerDivisionRootObject rootObject = (PartnerDivisionRootObject)ser.ReadObject(new MemoryStream(jsonData));
            createUpdatepartnerDivisoin(rootObject, service);
        }
        public static void GetPartnerDivisionFromExcel(IOrganizationService service)
        {
            string excelPath = ConfigurationManager.AppSettings["FilePath"].ToString();
            Application excelApp = new Application();
            List<PartnerDivisionResult> Results = new List<PartnerDivisionResult>();
            if (excelApp != null)
            {
                Workbook excelWorkbook = excelApp.Workbooks.Open(excelPath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet excelWorksheet = (Worksheet)excelWorkbook.Sheets[1];
                Range excelRange = excelWorksheet.UsedRange;
                int rowCount = excelRange.Rows.Count;
                int iLeft = excelRange.Rows.Count;
                int colCount = excelRange.Columns.Count;
                PartnerDivisionResult division = null;
                for (int i = 2; i <= rowCount; i++)
                {
                    division = new PartnerDivisionResult();
                    division.KUNNR = (excelWorksheet.Cells[i, 1] as Range).Value == null ? "" : (excelWorksheet.Cells[i, 1] as Range).Value.ToString();
                    division.dm_ch = (excelWorksheet.Cells[i, 2] as Range).Value == null ? "" : (excelWorksheet.Cells[i, 2] as Range).Value.ToString();
                    division.dm_div = (excelWorksheet.Cells[i, 3] as Range).Value == null ? "" : (excelWorksheet.Cells[i, 3] as Range).Value.ToString();
                    division.ERDAT = (excelWorksheet.Cells[i, 4] as Range).Value == null ? "" : (excelWorksheet.Cells[i, 4] as Range).Value.ToString();
                    division.STATUS = (excelWorksheet.Cells[i, 5] as Range).Value == null ? "" : (excelWorksheet.Cells[i, 5] as Range).Value.ToString();
                    division.DELETE_FLAG = (excelWorksheet.Cells[i, 6] as Range).Value == null ? "" : (excelWorksheet.Cells[i, 6] as Range).Value.ToString();
                    division.CREATEDBY = (excelWorksheet.Cells[i, 7] as Range).Value == null ? "" : (excelWorksheet.Cells[i, 7] as Range).Value.ToString();
                    division.CTIMESTAMP = (excelWorksheet.Cells[i, 8] as Range).Value == null ? "" : (excelWorksheet.Cells[i, 8] as Range).Value.ToString();
                    division.MODIFYBY = (excelWorksheet.Cells[i, 9] as Range).Value == null ? "" : (excelWorksheet.Cells[i, 9] as Range).Value.ToString();
                    division.MTIMESTAMP = (excelWorksheet.Cells[i, 10] as Range).Value == null ? "" : (excelWorksheet.Cells[i, 10] as Range).Value.ToString();
                    division.DM_SALES_OFFICE = (excelWorksheet.Cells[i, 11] as Range).Value == null ? "" : (excelWorksheet.Cells[i, 11] as Range).Value.ToString();
                    division.SAP_DIV = (excelWorksheet.Cells[i, 12] as Range).Value == null ? "" : (excelWorksheet.Cells[i, 12] as Range).Value.ToString();
                    division.SAP_CH = (excelWorksheet.Cells[i, 13] as Range).Value == null ? "" : (excelWorksheet.Cells[i, 13] as Range).Value.ToString();
                    //division.KDG = (excelWorksheet.Cells[i, 14] as Range).Value == null ? "" : (excelWorksheet.Cells[i, 14] as Range).Value.ToString();
                    //division.SAP_SA = (excelWorksheet.Cells[i, 15] as Range).Value == null ? "" : (excelWorksheet.Cells[i, 15] as Range).Value.ToString();
                    //division.VWERK = (excelWorksheet.Cells[i, 16] as Range).Value == null ? "" : (excelWorksheet.Cells[i, 16] as Range).Value.ToString();
                    Results.Add(division);
                }
                PartnerDivisionRootObject rootObject = new PartnerDivisionRootObject();
                rootObject.Results = Results;
                createUpdatepartnerDivisoin(rootObject, service);
            }
        }
        public static void createUpdatepartnerDivisoin(PartnerDivisionRootObject rootObject, IOrganizationService service)
        {
            string iKUNNR = string.Empty;
            bool isfranchise = false;
            Guid PartnnerDivisionMappingID = Guid.Empty;
            Guid DivisionId = Guid.Empty;
            Guid distributionchannel = Guid.Empty;
            Guid channelPartnerId = Guid.Empty;
            int iLeft = rootObject.Results.Count;
            int iDone = 0;
            int iTotal = rootObject.Results.Count;
            int GoodCount = 0;
            foreach (PartnerDivisionResult obj in rootObject.Results)
            {
                PartnnerDivisionMappingID = Guid.Empty;
                distributionchannel = Guid.Empty;
                iDone = iDone + 1;
                Console.WriteLine(iDone);

                if (obj.KUNNR.StartsWith("F") && obj.STATUS != "I")
                {
                    iKUNNR = obj.KUNNR.Substring(1);
                    iLeft = iLeft - 1;
                    Guid FranchiseeId = GetChanelPartner(iKUNNR, service);
                    DivisionId = GetDivisision(obj.dm_div, service);
                    if (FranchiseeId != Guid.Empty && DivisionId != Guid.Empty)
                    {
                        GoodCount = GoodCount + 1;
                        Entity iPartDivMapp = new Entity("hil_partnerdivisionmapping");
                        if (channelPartnerId != Guid.Empty)
                            iPartDivMapp["hil_franchiseedirectengineer"] = new EntityReference("account", channelPartnerId);
                        if (DivisionId != Guid.Empty)
                        {
                            EntityReference iProdCat = new EntityReference("product", DivisionId);
                            iPartDivMapp["hil_productcategory"] = (EntityReference)iProdCat;
                        }
                        if (obj.CTIMESTAMP == null)
                            iPartDivMapp["hil_mdmtimestamp"] = ConvertToDateTime(obj.CTIMESTAMP);
                        else
                            iPartDivMapp["hil_mdmtimestamp"] = ConvertToDateTime(obj.MTIMESTAMP);

                        iPartDivMapp["hil_franchiseecode"] = iKUNNR;
                        iPartDivMapp["hil_salesoffice"] = obj.DM_SALES_OFFICE;
                        iPartDivMapp["hil_channel"] = obj.dm_ch;
                        iPartDivMapp["hil_channelcode"] = obj.SAP_CH;
                        iPartDivMapp["hil_division"] = obj.dm_div;
                        iPartDivMapp["hil_divisioncode"] = obj.SAP_DIV;
                        iPartDivMapp["hil_name"] = obj.dm_div;
                        if (PartnnerDivisionMappingID == Guid.Empty)
                        {
                            service.Create(iPartDivMapp);
                        }
                        else
                        {
                            SetStateRequest setStateRequest = new SetStateRequest()
                            {
                                EntityMoniker = new EntityReference
                                {
                                    Id = PartnnerDivisionMappingID,
                                    LogicalName = iPartDivMapp.LogicalName,
                                },
                                State = new OptionSetValue(0),
                                Status = new OptionSetValue(1)
                            };
                            service.Execute(setStateRequest);//activate
                            iPartDivMapp.Id = PartnnerDivisionMappingID;
                            service.Update(iPartDivMapp);
                        }
                    }
                }
                else if (!obj.KUNNR.StartsWith("F") && obj.STATUS != "I")
                {
                    iKUNNR = obj.KUNNR;
                    iLeft = iLeft - 1;
                    isfranchise = GetFranchisee("F" + iKUNNR, service);
                    if (!isfranchise)
                    {
                        PartnnerDivisionMappingID = getPartnnerDivisionMapping(service, obj.KUNNR, obj.SAP_DIV, obj.SAP_CH);
                        channelPartnerId = GetChanelPartner(iKUNNR, service);
                        DivisionId = GetDivisision(obj.SAP_DIV, service);
                        if (channelPartnerId != Guid.Empty && DivisionId != Guid.Empty)
                        {
                            GoodCount = GoodCount + 1;
                            Entity iPartDivMapp = new Entity("hil_partnerdivisionmapping");
                            if (channelPartnerId != Guid.Empty)
                                iPartDivMapp["hil_franchiseedirectengineer"] = new EntityReference("account", channelPartnerId);
                            if (DivisionId != Guid.Empty)
                            {
                                EntityReference iProdCat = new EntityReference("product", DivisionId);
                                iPartDivMapp["hil_productcategory"] = (EntityReference)iProdCat;
                            }
                            if (obj.CTIMESTAMP == null)
                                iPartDivMapp["hil_mdmtimestamp"] = ConvertToDateTime(obj.CTIMESTAMP);
                            else
                                iPartDivMapp["hil_mdmtimestamp"] = ConvertToDateTime(obj.MTIMESTAMP);
                            distributionchannel = GetDistributionChannel(obj.SAP_CH, service);
                            if (distributionchannel != Guid.Empty)
                                iPartDivMapp["hil_distributionchannel"] = new EntityReference("hil_distributionchannel", distributionchannel);
                            iPartDivMapp["hil_franchiseecode"] = iKUNNR;
                            iPartDivMapp["hil_salesoffice"] = obj.DM_SALES_OFFICE;
                            iPartDivMapp["hil_channel"] = obj.dm_ch;
                            iPartDivMapp["hil_channelcode"] = obj.SAP_CH;
                            iPartDivMapp["hil_division"] = obj.dm_div;
                            iPartDivMapp["hil_divisioncode"] = obj.SAP_DIV;
                            iPartDivMapp["hil_name"] = obj.dm_div;
                            if (PartnnerDivisionMappingID == Guid.Empty)
                            {
                                try
                                {
                                   Guid mapping = service.Create(iPartDivMapp);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex.Message);
                                }
                            }
                            else
                            {
                                SetStateRequest setStateRequest = new SetStateRequest()
                                {
                                    EntityMoniker = new EntityReference
                                    {
                                        Id = PartnnerDivisionMappingID,
                                        LogicalName = iPartDivMapp.LogicalName,
                                    },
                                    State = new OptionSetValue(0),
                                    Status = new OptionSetValue(1)
                                };
                                service.Execute(setStateRequest);//activate
                                iPartDivMapp.Id = PartnnerDivisionMappingID;
                                service.Update(iPartDivMapp);
                            }
                        }
                    }
                }
                else if (obj.STATUS == "I")
                {
                    PartnnerDivisionMappingID = getPartnnerDivisionMapping(service, obj.KUNNR, obj.SAP_DIV, obj.SAP_CH);
                    //StateCode = 1 and StatusCode = 2 for deactivating Account or Contact
                    if (PartnnerDivisionMappingID != Guid.Empty)
                    {
                        SetStateRequest setStateRequest = new SetStateRequest()
                        {
                            EntityMoniker = new EntityReference
                            {
                                Id = PartnnerDivisionMappingID,
                                LogicalName = "hil_partnerdivisionmapping",
                            },
                            State = new OptionSetValue(1),
                            Status = new OptionSetValue(2)
                        };
                        service.Execute(setStateRequest);
                    }
                }
            }
        }
        public static DateTime? ConvertToDateTime(string _mdmTimeStamp)
        {
            DateTime? _dtMDMTimeStamp = null;
            try
            {
                _dtMDMTimeStamp = new DateTime(Convert.ToInt32(_mdmTimeStamp.Substring(0, 4)), Convert.ToInt32(_mdmTimeStamp.Substring(4, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(6, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(8, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(10, 2)), Convert.ToInt32(_mdmTimeStamp.Substring(12, 2)));
            }
            catch { }
            return _dtMDMTimeStamp;
        }

        public static Guid GetDistributionChannel(string sap_CH, IOrganizationService service)
        {
            Guid distributionchannel = new Guid();
            QueryExpression Query = new QueryExpression("hil_distributionchannel");
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_code", ConditionOperator.Equal, sap_CH);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                distributionchannel = Found.Entities[0].Id;
            }
            return distributionchannel;
        }
        public static Guid GetChanelPartner(string KUNNR, IOrganizationService service)
        {
            Guid iFranchise = new Guid();
            iFranchise = Guid.Empty;
            QueryExpression Query = new QueryExpression("account");
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            //Query.Criteria.AddCondition("hil_outwarrantycustomersapcode", ConditionOperator.Equal, KUNNR);
            Query.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, KUNNR);
            //Query.Criteria.AddCondition("hil_outwarrantycustomersapcode", ConditionOperator.Null);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                iFranchise = Found.Entities[0].Id;
            }
            return iFranchise;
        }
        public static bool GetFranchisee(string KUNNR, IOrganizationService service)
        {

            QueryExpression Query = new QueryExpression("account");
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_inwarrantycustomersapcode", ConditionOperator.Equal, KUNNR);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public static Guid getPartnnerDivisionMapping(IOrganizationService service, string kunnr, string division, string SAP_CH)
        {
            Guid partnerDivision = Guid.Empty;
            QueryExpression Query = new QueryExpression("hil_partnerdivisionmapping");
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_franchiseecode", ConditionOperator.Equal, kunnr);
            Query.Criteria.AddCondition("hil_divisioncode", ConditionOperator.Equal, division);
            Query.Criteria.AddCondition("hil_channelcode", ConditionOperator.Equal, SAP_CH);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count == 1)
            {
                partnerDivision = Found[0].Id;
            }
            return partnerDivision;
        }
        public static Guid GetDivisision(string SAP_CODE, IOrganizationService service)
        {
            Guid iDivision = new Guid();
            iDivision = Guid.Empty;
            QueryExpression Query = new QueryExpression("product");
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_sapcode", ConditionOperator.Equal, SAP_CODE);
            Query.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 2);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                iDivision = Found.Entities[0].Id;
            }
            return iDivision;
        }
        #endregion
        public static string getTimeStamp(IOrganizationService service)
        {
            string _enquiryDatetime = "20210804000000";
            QueryExpression qsCType = new QueryExpression("hil_partnerdivisionmapping");
            qsCType.ColumnSet = new ColumnSet("hil_mdmtimestamp");
            qsCType.NoLock = true;
            qsCType.TopCount = 1;
            qsCType.AddOrder("hil_mdmtimestamp", OrderType.Descending);
            EntityCollection entCol = service.RetrieveMultiple(qsCType);
            if (entCol.Entities.Count > 0)
            {
                DateTime _cTimeStamp = entCol.Entities[0].GetAttributeValue<DateTime>("hil_mdmtimestamp").AddMinutes(330).AddSeconds(-30);
                if (_cTimeStamp.Year.ToString().PadLeft(4, '0') != "0001")
                    _enquiryDatetime = _cTimeStamp.Year.ToString() + _cTimeStamp.Month.ToString().PadLeft(2, '0') + _cTimeStamp.Day.ToString().PadLeft(2, '0') + _cTimeStamp.Hour.ToString().PadLeft(2, '0') + _cTimeStamp.Minute.ToString().PadLeft(2, '0') + (_cTimeStamp.Second + 1).ToString().PadLeft(2, '0');
            }
            return _enquiryDatetime;
        }
        #region GET INTEGRATION URL
        public static string GetIntegrationUrl(IOrganizationService service, string RecName)
        {
            string sUrl = string.Empty;
            QueryExpression Query = new QueryExpression("hil_integrationconfiguration");
            Query.ColumnSet = new ColumnSet("hil_url");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, RecName);
            Query.Criteria.AddCondition("hil_url", ConditionOperator.NotNull);
            EntityCollection Found = service.RetrieveMultiple(Query);
            Entity enConfig = Found.Entities[0];
            sUrl = enConfig.GetAttributeValue<string>("hil_url");
            return sUrl;
        }
        #endregion
    }
    #region GET PARTNER DIVISION OBJECT
    public class PartnerDivisionResult
    {
        public string KUNNR { get; set; }
        public string dm_ch { get; set; }
        public string dm_div { get; set; }
        public string DM_SALES_OFFICE { get; set; }
        public string ERDAT { get; set; }
        public string STATUS { get; set; }
        public string DELETE_FLAG { get; set; }
        public string CREATEDBY { get; set; }
        public string MODIFYBY { get; set; }
        public string CTIMESTAMP { get; set; }
        public string MTIMESTAMP { get; set; }
        public string SAP_CH { get; set; }
        public string SAP_DIV { get; set; }
    }
    public class PartnerDivisionRootObject
    {
        public object Result { get; set; }
        public List<PartnerDivisionResult> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }
    #endregion
}
