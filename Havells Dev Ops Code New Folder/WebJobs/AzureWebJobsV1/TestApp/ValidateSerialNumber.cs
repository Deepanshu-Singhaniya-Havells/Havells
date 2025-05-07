using System;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Net;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Runtime.Serialization.Json;
using System.IO;
using Havells_Plugin.HelperIntegration;

namespace TestApp
{
    public class ValidateSerialNumber
    {
        public string SERIAL_NO { get; set; }
        public string PROD_CAT { get; set; }
        public string PROD_CATID { get; set; }
        public string PROD_SCAT { get; set; }
        public string PROD_SCATID { get; set; }
        public string PROD_NAME { get; set; }
        public string PROD_ID { get; set; }
        public string MESSAGE { get; set; }
        public bool IF_VALID { get; set; }
        public bool ModeOfEntry { get; set; }

        public ValidateSerialNumber GetResponseFromSAP(ValidateSerialNumber iValidate, IOrganizationService service)
        {
            //IOrganizationService service = ConnectToCRM.GetOrgService();
            try
            {
                bool IfExisting = CheckIfExistingSerialNumber(service, iValidate.SERIAL_NO);
                if (IfExisting == false || iValidate.ModeOfEntry)
                {
                    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                    {
                        //tracingService.Trace("4");
                        #region Credentials
                        String sUserName = String.Empty;
                        String sPassword = String.Empty;
                        var obj2 = from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
                                   where _IConfig.hil_name == "Credentials"
                                   select new { _IConfig };
                        foreach (var iobj2 in obj2)
                        {
                            if (iobj2._IConfig.hil_Username != String.Empty)
                                sUserName = iobj2._IConfig.hil_Username;
                            if (iobj2._IConfig.hil_Password != String.Empty)
                                sPassword = iobj2._IConfig.hil_Password;
                        }
                        #endregion
                        var obj = from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
                                  where _IConfig.hil_name == "SerialNumberValidation"
                                  select new { _IConfig.hil_Url };
                        foreach (var iobj in obj)
                        {
                            if (iobj.hil_Url != null)
                            {
                                String sUrl = iobj.hil_Url + iValidate.SERIAL_NO;
                                WebClient webClient = new WebClient();
                                //Authentication
                                string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
                                webClient.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;

                                webClient.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                                var jsonData = webClient.DownloadData(sUrl);
                                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(SerialNumberValidation));
                                SerialNumberValidation rootObject = (SerialNumberValidation)ser.ReadObject(new MemoryStream(jsonData));

                                if (rootObject.EX_PRD_DET != null)//valid
                                {
                                    //tracingService.Trace("9");
                                    #region UpdateMaterialCodeonAsset
                                    var obj1 = (from _Product in orgContext.CreateQuery<Product>()
                                                where _Product.ProductNumber == rootObject.EX_PRD_DET.MATNR
                                                select new
                                                {
                                                    _Product.ProductId,
                                                    _Product.hil_MaterialGroup,
                                                    _Product.hil_Division,
                                                    _Product.ProductNumber
                                                }).Take(1);
                                    foreach (var iobj1 in obj1)
                                    {
                                        iValidate.PROD_ID = iobj1.ProductId.Value.ToString();
                                        iValidate.PROD_NAME = iobj1.ProductNumber;
                                        QueryExpression Query = new QueryExpression(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName);
                                        Query.ColumnSet = new ColumnSet("hil_productcategorydivision", "hil_name", "hil_productsubcategorymg", "statecode");
                                        Query.Criteria = new FilterExpression(LogicalOperator.And);
                                        Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, iobj1.hil_MaterialGroup.Name);
                                        EntityCollection Found = service.RetrieveMultiple(Query);
                                        if (Found.Entities.Count > 0)
                                        {
                                            hil_stagingdivisonmaterialgroupmapping iMapping = Found.Entities[0].ToEntity<hil_stagingdivisonmaterialgroupmapping>();
                                            iValidate.PROD_CAT = iMapping.hil_ProductCategoryDivision.Name;
                                            iValidate.PROD_CATID = iMapping.hil_ProductCategoryDivision.Id.ToString();
                                            iValidate.PROD_SCAT = iMapping.hil_name;
                                            iValidate.PROD_SCATID = iMapping.Id.ToString();
                                            iValidate.MESSAGE = "VALID SERIAL NUMBER";
                                            iValidate.IF_VALID = true;
                                        }
                                        else
                                        {
                                            iValidate.PROD_CAT = "";
                                            iValidate.PROD_CATID = "";
                                            iValidate.PROD_SCAT = "";
                                            iValidate.PROD_SCATID = "";
                                            iValidate.MESSAGE = "SERIAL NUMBER VALID / SUB CATEGORY NOT FOUND";
                                            iValidate.IF_VALID = false;
                                        }
                                    }
                                    #endregion
                                }
                                else//Not valid
                                {
                                    iValidate.PROD_ID = "";
                                    iValidate.PROD_NAME = "";
                                    iValidate.PROD_CAT = "";
                                    iValidate.PROD_CATID = "";
                                    iValidate.PROD_SCAT = "";
                                    iValidate.PROD_SCATID = "";
                                    iValidate.MESSAGE = "SERIAL NUMBER NOT VALID";
                                    iValidate.IF_VALID = false;
                                }
                            }
                        }
                    }
                }
                else
                {
                    iValidate.PROD_ID = "";
                    iValidate.PROD_NAME = "";
                    iValidate.PROD_CAT = "";
                    iValidate.PROD_CATID = "";
                    iValidate.PROD_SCAT = "";
                    iValidate.PROD_SCATID = "";
                    iValidate.MESSAGE = "SERIAL NUMBER ALREADY EXISTS";
                    iValidate.IF_VALID = false;
                }
            }
            catch (Exception ex)
            {
                iValidate.PROD_ID = "";
                iValidate.PROD_NAME = "";
                iValidate.PROD_CAT = "";
                iValidate.PROD_CATID = "";
                iValidate.PROD_SCAT = "";
                iValidate.PROD_SCATID = "";
                iValidate.MESSAGE = ex.Message.ToUpper();
                iValidate.IF_VALID = false;
            }
            return iValidate;
        }

        public ValidateSerialNumber ValidateSerialNumberWithSAP(ValidateSerialNumber iValidate)
        {
            //IOrganizationService service = ConnectToCRM.GetOrgService();
            try
            {
                //using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                //{
                //tracingService.Trace("4");
                #region Credentials
                String sUserName = String.Empty;
                String sPassword = String.Empty;

                //var obj2 = from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
                //           where _IConfig.hil_name == "Credentials"
                //           select new { _IConfig };
                //foreach (var iobj2 in obj2)
                //{
                //    if (iobj2._IConfig.hil_Username != String.Empty)
                //        sUserName = iobj2._IConfig.hil_Username;
                //    if (iobj2._IConfig.hil_Password != String.Empty)
                //        sPassword = iobj2._IConfig.hil_Password;
                //}
                #endregion
                //var obj = from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
                //          where _IConfig.hil_name == "SerialNumberValidation"
                //          select new { _IConfig.hil_Url };
                //foreach (var iobj in obj)
                //{
                //    if (iobj.hil_Url != null)
                //    {
                String sUrl = "https://middleware.havells.com:50001/RESTAdapter/Service/ProductDetails?IM_SERIAL_NO=" + iValidate.SERIAL_NO;
                WebClient webClient = new WebClient();
                //Authentication
                string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes("D365_HAVELLS" + ":" + "PRDD365@1234"));
                webClient.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;

                webClient.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                var jsonData = webClient.DownloadData(sUrl);
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(SerialNumberValidation));
                SerialNumberValidation rootObject = (SerialNumberValidation)ser.ReadObject(new MemoryStream(jsonData));

                if (rootObject.EX_PRD_DET != null)//valid
                {
                    iValidate.SERIAL_NO = iValidate.SERIAL_NO;
                    iValidate.MESSAGE = "VALID SERIAL NUMBER";
                    iValidate.IF_VALID = true;
                }
                else//Not valid
                {
                    iValidate.SERIAL_NO = iValidate.SERIAL_NO;
                    iValidate.MESSAGE = "SERIAL NUMBER NOT VALID";
                    iValidate.IF_VALID = false;
                }
                //    }
                //}
                //}
            }
            catch (Exception ex)
            {
                iValidate.SERIAL_NO = iValidate.SERIAL_NO;
                iValidate.MESSAGE = ex.Message.ToUpper();
                iValidate.IF_VALID = false;
            }
            return iValidate;
        }
        public bool CheckIfExistingSerialNumber(IOrganizationService service, string Serial)
        {
            bool IfExisting = true;
            QueryExpression Query = new QueryExpression(msdyn_customerasset.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, Serial);
            EntityCollection Found = service.RetrieveMultiple(Query);
            {
                if(Found.Entities.Count > 0)
                {
                    IfExisting = true;
                }
                else
                {
                    IfExisting = false;
                }
            }
            return IfExisting;
        }

        public CancelJobResponse CancelServiceJob(CancelJobRequest reqParam, IOrganizationService service)
        {
            CancelJobResponse res = new CancelJobResponse();
            try
            {
                if (reqParam.JobGuid == string.Empty || reqParam.JobGuid == null)
                {
                    res = new CancelJobResponse { StatusCode = "204", StatusDescription = "Job Id is required" };
                    return res;
                }
                //IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    QueryExpression Query = new QueryExpression("msdyn_workorder");
                    Query.ColumnSet = new ColumnSet("msdyn_substatus");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("msdyn_workorderid", ConditionOperator.Equal, new Guid(reqParam.JobGuid));
                    EntityCollection entCols = service.RetrieveMultiple(Query);
                    if (entCols.Entities.Count > 0)
                    {
                        EntityReference _erSubstatus = entCols.Entities[0].GetAttributeValue<EntityReference>("msdyn_substatus");
                        if (_erSubstatus.Id.ToString().ToUpper() == "1527FA6C-FA0F-E911-A94E-000D3AF060A1")
                        {
                            res = new CancelJobResponse { StatusCode = "204", StatusDescription = "This Job is already Cancelled" };
                            return res;
                        }
                        string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_statustransitionmatrix'>
                        <attribute name='hil_statustransitionmatrixid' />
                        <filter type='and'>
                            <condition attribute='hil_basestatus' operator='eq' value='"+ _erSubstatus.Id + @"' />
                            <filter type='or'>
                            <condition attribute='hil_statustransition1' operator='eq' value='{1527FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                            <condition attribute='hil_statustransition10' operator='eq' value='{1527FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                            <condition attribute='hil_statustransition2' operator='eq' value='{1527FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                            <condition attribute='hil_statustransition3' operator='eq' value='{1527FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                            <condition attribute='hil_statustransition4' operator='eq' value='{1527FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                            <condition attribute='hil_statustransition5' operator='eq' value='{1527FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                            <condition attribute='hil_statustransition6' operator='eq' value='{1527FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                            <condition attribute='hil_statustransition7' operator='eq' value='{1527FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                            <condition attribute='hil_statustransition8' operator='eq' value='{1527FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                            <condition attribute='hil_statustransition9' operator='eq' value='{1527FA6C-FA0F-E911-A94E-000D3AF060A1}' />
                            </filter>
                        </filter>
                        </entity>
                        </fetch>";
                        entCols = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entCols.Entities.Count == 0) {
                            res = new CancelJobResponse { StatusCode = "204", StatusDescription = "This Job cannot be Cancelled, please contact with Customer Care." };
                            return res;
                        }
                    }
                    else {
                        res = new CancelJobResponse { StatusCode = "204", StatusDescription = "Invalid Job ID." };
                        return res;
                    }

                    Query = new QueryExpression("msdyn_workordersubstatus");
                    Query.ColumnSet = new ColumnSet(false);
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, "Canceled");
                    entCols = service.RetrieveMultiple(Query);
                    if (entCols.Entities.Count > 0)
                    {
                        Entity _jobUpdate = new Entity("msdyn_workorder");
                        _jobUpdate.Id = new Guid(reqParam.JobGuid);
                        _jobUpdate["hil_closureremarks"] = "Job cancelled by Customer via Sync App";
                        _jobUpdate["hil_jobcancelreason"] = new OptionSetValue(7);
                        _jobUpdate["msdyn_substatus"] = new EntityReference(entCols.EntityName, entCols[0].Id);

                        DateTime _today = DateTime.Now;

                        _today = _today.AddHours(5);
                        _today = _today.AddMinutes(30);

                        _jobUpdate["hil_jobclosureon"] = _today;
                        _jobUpdate["msdyn_timeclosed"] = _today;

                        try
                        {
                            service.Update(_jobUpdate);
                        }
                        catch (Exception ex)
                        {
                            res = new CancelJobResponse { StatusCode = "204", StatusDescription = "This Job cannot be Cancelled, please contact with Customer Care. \n" + ex.Message };
                            return res;
                        }
                        res.StatusCode = "200";
                        res.StatusDescription = "Job is Cancelled";
                    }
                    else
                    {
                        res = new CancelJobResponse { StatusCode = "204", StatusDescription = "Job Sub status(Cancel) is not defined in System." };
                        return res;
                    }
                }
                else
                {
                    res = new CancelJobResponse { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                res = new CancelJobResponse { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return res;
        }
    }

    public class CancelJobRequest
    {
        public string JobGuid { get; set; }
    }
    public class CancelJobResponse
    {
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
    }
}
