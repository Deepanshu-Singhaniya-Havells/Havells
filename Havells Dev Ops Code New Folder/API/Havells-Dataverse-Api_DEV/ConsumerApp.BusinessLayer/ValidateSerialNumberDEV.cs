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
using System.Collections.Generic;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class ValidateSerialNumberDEV
    {
        [DataMember]
        public string SERIAL_NO { get; set; }
        [DataMember(IsRequired = false)]
        public string PROD_CAT { get; set; }
        [DataMember(IsRequired = false)]
        public string PROD_CATID { get; set; }
        [DataMember(IsRequired = false)]
        public string PROD_SCAT { get; set; }
        [DataMember(IsRequired = false)]
        public string PROD_SCATID { get; set; }
        [DataMember(IsRequired = false)]
        public string PROD_NAME { get; set; }
        [DataMember(IsRequired = false)]
        public string PROD_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string DIVISION { get; set; }
        [DataMember(IsRequired = false)]
        public int BRAND { get; set; }
        [DataMember(IsRequired = false)]
        public string BRANDName { get; set; }
        [DataMember(IsRequired = false)]
        public string REGION { get; set; }
        [DataMember(IsRequired = false)]
        public string INV_NO { get; set; }
        [DataMember(IsRequired = false)]
        public string INV_DATE { get; set; }
        [DataMember(IsRequired = false)]
        public string CUST_CODE { get; set; }
        [DataMember(IsRequired = false)]
        public string CUST_NAME { get; set; }
        [DataMember(IsRequired = false)]
        public string WTY_STATUS { get; set; }
        [DataMember(IsRequired = false)]
        public string IS_TYPE { get; set; }
        [DataMember(IsRequired = false)]
        public string MESSAGE { get; set; }
        [DataMember(IsRequired = false)]
        public bool IF_VALID { get; set; }

        public ValidateSerialNumberDEV GetResponseFromSAPDEV(ValidateSerialNumberDEV iValidate)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            try
            {
                bool IfExisting = false;
                if (IfExisting == false)
                {
                    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                    {
                        QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                        qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                        qsCType.NoLock = true;
                        qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                        qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "ValidateServialNumberMaterialReturn");
                        Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];

                        string hil_Url = integrationConfiguration.GetAttributeValue<string>("hil_url");
                        String sUserName = integrationConfiguration.GetAttributeValue<string>("hil_username");
                        String sPassword = integrationConfiguration.GetAttributeValue<string>("hil_password");

                        String sUrl = hil_Url + iValidate.SERIAL_NO;

                        WebClient webClient = new WebClient();
                        string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
                        webClient.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;

                        var jsonData = webClient.DownloadData(sUrl);
                        DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(SerialNumberValidationDev));
                        SerialNumberValidationDev rootObject = (SerialNumberValidationDev)ser.ReadObject(new MemoryStream(jsonData));

                        if (rootObject.EX_PRD_DET != null)//valid
                        {
                            //tracingService.Trace("9");
                            #region UpdateMaterialCodeonAsset
                            var obj1 = (from _Product in orgContext.CreateQuery<Product>()
                                        where _Product.ProductNumber == rootObject.EX_PRD_DET[0].MATNR
                                        select new
                                        {
                                            _Product.ProductId,
                                            _Product.hil_MaterialGroup,
                                            _Product.hil_Division,
                                            _Product.ProductNumber,
                                            _Product.hil_BrandIdentifier
                                        }).Take(1);
                            foreach (var iobj1 in obj1)
                            {
                                iValidate.PROD_ID = iobj1.ProductId.Value.ToString();
                                iValidate.PROD_NAME = iobj1.ProductNumber;
                                iValidate.DIVISION = rootObject.EX_PRD_DET[0].SPART;
                                iValidate.REGION = rootObject.EX_PRD_DET[0].REGIO;
                                iValidate.INV_NO = rootObject.EX_PRD_DET[0].VBELN;
                                iValidate.INV_DATE = rootObject.EX_PRD_DET[0].FKDAT;
                                iValidate.CUST_CODE = rootObject.EX_PRD_DET[0].KUNAG;
                                iValidate.CUST_NAME = rootObject.EX_PRD_DET[0].NAME1;
                                iValidate.WTY_STATUS = rootObject.EX_PRD_DET[0].WTY_STATUS;
                                iValidate.IS_TYPE = rootObject.EX_PRD_DET[0].IS_TYPE;

                                if (iobj1.hil_BrandIdentifier != null)
                                {
                                    iValidate.BRAND = iobj1.hil_BrandIdentifier.Value;
                                    iValidate.BRANDName = iobj1.hil_BrandIdentifier.Value == 1 ? "Havells" : iobj1.hil_BrandIdentifier.Value == 2 ? "Lloyd" : "WaterPurifier";
                                }
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
                        //    }
                        //}
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

        public bool CheckIfExistingSerialNumber(IOrganizationService service, string Serial)
        {
            bool IfExisting = true;
            QueryExpression Query = new QueryExpression(msdyn_customerasset.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, Serial);
            EntityCollection Found = service.RetrieveMultiple(Query);
            {
                if (Found.Entities.Count > 0)
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
    }

    [DataContract]
    public class SerialNumberValidationDev
    {
        [DataMember]
        public string EX_RETURN { get; set; }
        [DataMember]
        public List<EX_PRD_DETDev> EX_PRD_DET { get; set; }
    }
    [DataContract]
    public class EX_PRD_DETDev
    {
        [DataMember]
        public string SERIAL_NO { get; set; }
        [DataMember]
        public string MATNR { get; set; }
        [DataMember]
        public string MAKTX { get; set; }
        [DataMember]
        public string SPART { get; set; }
        [DataMember]
        public string REGIO { get; set; }
        [DataMember]
        public string VBELN { get; set; }
        [DataMember]
        public string FKDAT { get; set; }
        [DataMember]
        public string KUNAG { get; set; }
        [DataMember]
        public string NAME1 { get; set; }
        [DataMember]
        public string WTY_STATUS { get; set; }
        [DataMember]
        public string IS_TYPE { get; set; }
    }
}
