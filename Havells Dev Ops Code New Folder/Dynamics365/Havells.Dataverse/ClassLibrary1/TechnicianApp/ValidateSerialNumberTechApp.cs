using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;

namespace Havells.Dataverse.CustomConnector.TechnicianApp
{
    public class ValidateSerialNumberTechApp : IPlugin 
    {
        public static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (!context.InputParameters.Contains("SerialNumber"))
                {
                    context.OutputParameters["Status"] = false;
                    context.OutputParameters["Result"] = "Serial Number is required.";
                    return;
                }
                else
                {
                    try
                    {
                        string SerialNumber = (string)context.InputParameters["SerialNumber"];
                        StringBuilder _returnMessage = new StringBuilder();
                        if (string.IsNullOrEmpty(SerialNumber))
                        {
                            context.OutputParameters["Status"] = false;
                            context.OutputParameters["Result"] = "Serial Number is required. ";
                            return;
                        }

                        string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='msdyn_customerasset'>
                                <attribute name='createdon' />
                                <attribute name='msdyn_product' />
                                <attribute name='msdyn_name' />
                                <attribute name='hil_productsubcategorymapping' />
                                <attribute name='hil_productcategory' />
                                <attribute name='msdyn_customerassetid' />
                                <attribute name='hil_invoicedate' />
                                <attribute name='hil_customer' />
                                <order attribute='createdon' descending='true' />
                                <filter type='and'>
                                  <condition attribute='msdyn_name' operator='eq' value='{SerialNumber}' />
                                </filter>
                              </entity>
                            </fetch>";
                        EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (_entCol.Entities.Count > 0)
                        {
                            _returnMessage.AppendLine("Serial Number is already exist.");
                            if (_entCol.Entities[0].Contains("hil_productcategory"))
                                _returnMessage.AppendLine($"Product Category: {_entCol.Entities[0].GetAttributeValue<EntityReference>("hil_productcategory").Name}");
                            if (_entCol.Entities[0].Contains("hil_productsubcategorymapping"))
                                _returnMessage.AppendLine($"Product Subcategory: {_entCol.Entities[0].GetAttributeValue<EntityReference>("hil_productsubcategorymapping").Name}");
                            if (_entCol.Entities[0].Contains("msdyn_product"))
                                _returnMessage.AppendLine($"Model Code: {_entCol.Entities[0].GetAttributeValue<EntityReference>("msdyn_product").Name}");
                            if (_entCol.Entities[0].Contains("hil_invoicedate"))
                                _returnMessage.AppendLine($"Invoice Date: {_entCol.Entities[0].GetAttributeValue<DateTime>("hil_invoicedate").AddMinutes(330).Date.ToString("dd/MM/yyyy")}");
                            if (_entCol.Entities[0].Contains("hil_customer"))
                                _returnMessage.AppendLine($"Customer: {_entCol.Entities[0].GetAttributeValue<EntityReference>("hil_customer").Name}");
                            context.OutputParameters["Status"] = false;
                            context.OutputParameters["Result"] = _returnMessage.ToString();
                            return;
                        }
                        else
                        {
                            SerialNumberInfo _retObj = ValidateSerialNumberWithSAP(service, SerialNumber);
                            if (_retObj.SerialNumber == SerialNumber)
                            {
                                context.OutputParameters["Status"] = true;
                                context.OutputParameters["Result"] = JsonConvert.SerializeObject(_retObj);
                                return;
                            }
                            else
                            {
                                context.OutputParameters["Status"] = false;
                                context.OutputParameters["Result"] = _retObj.SerialNumber;
                                return;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        context.OutputParameters["Result"] = "ERROR: " + ex.Message;
                        context.OutputParameters["Status"] = false;
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace("Status...5");
                context.OutputParameters["Result"] = "D365 Internal Server Error : " + ex.Message;
                context.OutputParameters["Status"] = false;
            }
        }
        private SerialNumberInfo ValidateSerialNumberWithSAP(IOrganizationService service, string SerialNumber)
        {
            SerialNumberInfo _retValue = new SerialNumberInfo();
            string sUserName = string.Empty;
            string sPassword = string.Empty;
            string hil_Url = string.Empty;
            QueryExpression Query = new QueryExpression("hil_integrationconfiguration");
            Query.ColumnSet = new ColumnSet("hil_username", "hil_password");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "Credentials");
            EntityCollection EntColl = service.RetrieveMultiple(Query);

            sUserName = EntColl.Entities[0].Contains("hil_username") ? EntColl.Entities[0].GetAttributeValue<string>("hil_username") : null;
            sPassword = EntColl.Entities[0].Contains("hil_password") ? EntColl.Entities[0].GetAttributeValue<string>("hil_password") : null;
            Query = new QueryExpression("hil_integrationconfiguration");
            Query.ColumnSet = new ColumnSet("hil_url");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "SerialNumberValidation");
            EntColl = service.RetrieveMultiple(Query);

            hil_Url = EntColl.Entities[0].Contains("hil_url") ? EntColl.Entities[0].GetAttributeValue<string>("hil_url") : null;
            if (hil_Url != null)
            {
                String sUrl = hil_Url + SerialNumber;
                WebClient webClient = new WebClient();
                string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
                webClient.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
                webClient.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                var jsonData = webClient.DownloadData(sUrl);

                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(SerialNumberValidation));
                SerialNumberValidation rootObject = (SerialNumberValidation)ser.ReadObject(new MemoryStream(jsonData));

                if (rootObject.EX_PRD_DET != null)//valid
                {
                    string fetchQuery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                    <entity name='product'>
                    <attribute name='hil_materialgroup' />
                    <attribute name='hil_division' />
                    <filter type='and'>
                    <condition attribute='name' operator='eq' value='{rootObject.EX_PRD_DET.MATNR}' />
                    </filter>
                    </entity>
                    </fetch>";

                    EntityCollection _entCol = service.RetrieveMultiple(new FetchExpression(fetchQuery));
                    if (_entCol.Entities.Count > 0)
                    {
                        _retValue.ModelCode = rootObject.EX_PRD_DET.MATNR;
                        _retValue.ModelName = rootObject.EX_PRD_DET.MAKTX;

                        if (rootObject.EX_PRD_DET.VBELN != null)
                            _retValue.PrimaryInvoiceNumber = rootObject.EX_PRD_DET.VBELN;
                        if (rootObject.EX_PRD_DET.FKDAT != null)
                            _retValue.PrimaryInvoiceDate = rootObject.EX_PRD_DET.FKDAT;
                        if (rootObject.EX_PRD_DET.KUNAG != null)
                            _retValue.SoldToPartyCode = rootObject.EX_PRD_DET.KUNAG;
                        if (rootObject.EX_PRD_DET.NAME1 != null)
                            _retValue.SoldToParty = rootObject.EX_PRD_DET.NAME1;
                        if (rootObject.EX_PRD_DET.MFD != null)
                            _retValue.ManufacturingDate = $"{rootObject.EX_PRD_DET.MFD.Substring(0, 4)}-{rootObject.EX_PRD_DET.MFD.Substring(4, 2)}-{rootObject.EX_PRD_DET.MFD.Substring(6, 2)}";

                        if (_entCol.Entities[0].Contains("hil_division"))
                        {
                            _retValue.CategoryId = _entCol.Entities[0].GetAttributeValue<EntityReference>("hil_division").Id;
                            _retValue.CategoryName = _entCol.Entities[0].GetAttributeValue<EntityReference>("hil_division").Name;
                        }
                        if (_entCol.Entities[0].Contains("hil_materialgroup"))
                        {
                            _retValue.SubCategoryId = _entCol.Entities[0].GetAttributeValue<EntityReference>("hil_materialgroup").Id;
                            _retValue.SubCategoryName = _entCol.Entities[0].GetAttributeValue<EntityReference>("hil_materialgroup").Name;
                        }
                        _retValue.ModelId = _entCol.Entities[0].Id;
                        _retValue.SerialNumber = SerialNumber;
                    }
                    else
                    {
                        _retValue.SerialNumber = $"{_retValue.ModelCode} Model Code doesn't exist in CRM.";
                    }
                }
                else
                {
                    _retValue.SerialNumber = "Serial Number doesn't exist in SAP.";
                }
            }
            return _retValue;
        }
    }
    public class SerialNumberInfo
    {
        public string SerialNumber { get; set; }
        public Guid CategoryId { get; set; }
        public string CategoryName { get; set; }
        public Guid SubCategoryId { get; set; }
        public string SubCategoryName { get; set; }
        public Guid ModelId { get; set; }
        public string ModelName { get; set; }
        public string ModelCode { get; set; }
        public string SoldToParty { get; set; }
        public string SoldToPartyCode { get; set; }
        public string PrimaryInvoiceNumber { get; set; }
        public string PrimaryInvoiceDate { get; set; }
        public string ManufacturingDate { get; set; }

    }
    public class SerialNumberValidation
    {
        public EX_PRD_DET EX_PRD_DET;
    }
    public class EX_PRD_DET
    {

        public string SERIAL_NO;
        public string MATNR;
        public string MAKTX;
        public string SPART;
        public string REGIO;
        public string VBELN;
        public string FKDAT;
        public string KUNAG;
        public string NAME1;
        public string MFD;
        public string WTY_STATUS;
    }
}
