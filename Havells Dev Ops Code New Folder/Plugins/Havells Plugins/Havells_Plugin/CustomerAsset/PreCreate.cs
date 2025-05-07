using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
// Microsoft Dynamics CRM namespace(s)
using System.Net;
using System.Runtime.Serialization.Json;
using System.IO;
using Havells_Plugin.HelperIntegration;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk.Query;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Security.Authentication;

namespace Havells_Plugin.CustomerAsset
{
    public class PreCreate : IPlugin
    {
        public static ITracingService tracingService = null;

        public void Execute(IServiceProvider serviceProvider)
        {
            #region PluginConfig
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "msdyn_customerasset"
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    //Function to create unit warranty
                    //HelperWarranty.Warranty_InitFun(entity,service);
                    tracingService.Trace("1");
                    msdyn_customerasset cAst = entity.ToEntity<msdyn_customerasset>();

                    ValidateSerialNumberRegExp(cAst.msdyn_name);
                    IfDuplicate(service, cAst);
                    AssetDefaultValues(service, cAst);


                    if (!entity.Contains("msdyn_parentasset"))
                    {
                        if (cAst.hil_ProductSubcategory != null)
                        {
                            Product iSCat = (Product)service.Retrieve(Product.EntityLogicalName, cAst.hil_ProductSubcategory.Id, new ColumnSet("hil_isserialized"));
                            if (iSCat.hil_IsSerialized != null && iSCat.hil_IsSerialized.Value == 1 && entity.GetAttributeValue<bool>("hil_invoiceavailable"))
                            {
                                SerialNumberValidation_SAP(entity, service);
                            }
                            else
                            {
                                tracingService.Trace("2");
                                PopulateName(entity, service);
                            }
                        }
                    }
                    if (entity.Attributes.Contains("hil_source"))
                    {
                        OptionSetValue optSource = entity.GetAttributeValue<OptionSetValue>("hil_source");
                        if (optSource.Value == 9)//SFA
                        {
                            entity["statuscode"] = new OptionSetValue(910590001);
                            entity["hil_branchheadapprovalstatus"] = new OptionSetValue(1);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message);
            }
            #endregion
        }
        public static void SetSourceOfCreation(Entity customerAsset, IOrganizationService service)
        {
            OptionSetValue optSource = null;
            EntityReference erCreatedby = null;
            EntityReference erCreatedonbehalfby = null;
            if (customerAsset.Attributes.Contains("hil_source"))
            {
                optSource = customerAsset.GetAttributeValue<OptionSetValue>("hil_source");
            }
            else {
                erCreatedby = customerAsset.GetAttributeValue<EntityReference>("createdby");
                erCreatedonbehalfby = customerAsset.GetAttributeValue<EntityReference>("createdonbehalfby");
                if (erCreatedby.Name.ToUpper() == "SYSTEM")
                {
                    optSource = new OptionSetValue(2); //Customer Portal
                }
                else if (erCreatedonbehalfby != null)
                {
                    if (erCreatedonbehalfby.Name.ToUpper().IndexOf("CRM ADMIN") >= 0)
                    {
                        optSource = new OptionSetValue(3); //Technician App
                    }
                    else
                    {
                        optSource = new OptionSetValue(4); //Web Client
                    }
                }
                else
                {
                    Entity ent = service.Retrieve(SystemUser.EntityLogicalName, erCreatedby.Id, new ColumnSet("positionid"));
                    if (ent != null)
                    {
                        if (ent.GetAttributeValue<EntityReference>("positionid").Id == new Guid("0197EA9B-1208-E911-A94D-000D3AF0694E") || ent.GetAttributeValue<EntityReference>("positionid").Id == new Guid("7D1ECBAB-1208-E911-A94D-000D3AF0694E"))
                        {
                            optSource = new OptionSetValue(3); //Technician App
                        }
                        else
                        {
                            optSource = new OptionSetValue(4); //Web Client
                        }
                    }
                    else
                    {
                        optSource = new OptionSetValue(4); //Web Client
                    }
                }
                customerAsset["hil_source"] = optSource;
            }
            if (optSource.Value == 3 || optSource.Value == 4)
            {
                bool invoiceAvailable = false;
                if (customerAsset.Attributes.Contains("hil_invoiceavailable"))
                {
                    invoiceAvailable = customerAsset.GetAttributeValue<bool>("hil_invoiceavailable");
                }
                if (invoiceAvailable)
                {

                }
                else {
                    customerAsset["statuscode"] = new OptionSetValue(910590001);
                    customerAsset["hil_branchheadapprovalstatus"] = new OptionSetValue(1);
                    customerAsset["hil_warrantystatus"] = new OptionSetValue(2);
                    customerAsset["hil_warrantytilldate"] = DateTime.Now.AddDays(-1);
                    customerAsset["hil_createwarranty"] = false;
                }
                //hil_invoiceavailable {true,false}
                //hil_createwarranty {true,false}
                //hil_warrantystatus {1-IN,2-OUT}
                //hil_warrantytilldate
                //hil_branchheadapprovalstatus {0 -Pending,1-Approved}
                //statuscode {910590000-Pending,910590001--Approved}
            }
            else {
                customerAsset["statuscode"] = new OptionSetValue(910590000);
                customerAsset["hil_branchheadapprovalstatus"] = new OptionSetValue(0);
                customerAsset["hil_createwarranty"] = false;
            }
        }
        #region Populate Name
        public static void PopulateName(Entity Ticket, IOrganizationService service)
        {
            string AutoName = string.Empty;
            string ModelName = string.Empty;

            if (!Ticket.Attributes.Contains("msdyn_name"))
            {
                if (Ticket.Attributes.Contains("hil_modelname"))
                {
                    ModelName = Ticket.GetAttributeValue<string>("hil_modelname");
                    ModelName = ModelName.Replace(" ", string.Empty);
                    ModelName = ModelName.Substring(0, 8);
                    AutoName = ModelName.ToUpper() + "-";
                }
                else
                {
                    AutoName = "UNDEF-";
                }
                AutoName = AutoName + DateTime.Now.AddMinutes(330).ToString("dd-mm");
                string RunningNum = GetRunningNumber(service);
                if (RunningNum != string.Empty)
                {
                    AutoName = AutoName + RunningNum;
                }
                Ticket["msdyn_name"] = AutoName;
            }
        }
        #endregion
        #region Get Running Number
        public static string GetRunningNumber(IOrganizationService service)
        {
            string RunningNum = string.Empty;
            Int32 Runn = 0;
            QueryExpression Query = new QueryExpression(hil_integrationconfiguration.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("hil_penaltyhours", "hil_name");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, "Non Serialized Asset Serial Number"));
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                hil_integrationconfiguration iConf = Found.Entities[0].ToEntity<hil_integrationconfiguration>();
                if (iConf.hil_penaltyhours != null)
                {
                    RunningNum = "000" + iConf.hil_penaltyhours.ToString();
                    Runn = Convert.ToInt32(iConf.hil_penaltyhours) + Convert.ToInt32(1);
                    iConf.hil_penaltyhours = Runn;
                    service.Update(iConf);
                }
                else
                {
                    RunningNum = "NAA" + DateTime.Now.Second.ToString();
                    iConf.hil_penaltyhours = Convert.ToInt32(0);
                    service.Update(iConf);
                }
            }
            else
            {
                RunningNum = "NAA" + DateTime.Now.Second.ToString();
            }
            return RunningNum;
        }
        #endregion
        //public static void SerialNumberValidation_SAP(Entity entity, IOrganizationService service)
        //{
        //    try
        //    {
        //        String sSerialNumber = entity.GetAttributeValue<String>("msdyn_name");
        //        using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
        //        {
        //            #region Credentials
        //            String sUserName = String.Empty;
        //            String sPassword = String.Empty;
        //            var obj2 = from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
        //                       where _IConfig.hil_name == "Credentials"
        //                       select new { _IConfig };
        //            foreach (var iobj2 in obj2)
        //            {
        //                if (iobj2._IConfig.hil_Username != String.Empty)
        //                    sUserName = iobj2._IConfig.hil_Username;
        //                if (iobj2._IConfig.hil_Password != String.Empty)
        //                    sPassword = iobj2._IConfig.hil_Password;
        //            }
        //            #endregion

        //            var obj = from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
        //                where _IConfig.hil_name == "SerialNumberValidation"
        //                select new { _IConfig.hil_Url };
        //            foreach (var iobj in obj)
        //            {
        //                if (iobj.hil_Url != null)
        //                {
        //                    String sUrl = iobj.hil_Url + sSerialNumber;

        //                    WebClient webClient = new WebClient();
        //                    string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
        //                    webClient.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
        //                    webClient.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
        //                    var jsonData = webClient.DownloadData(sUrl);

        //                    DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(SerialNumberValidation));
        //                    SerialNumberValidation rootObject = (SerialNumberValidation)ser.ReadObject(new MemoryStream(jsonData));

        //                    if (rootObject.EX_PRD_DET != null)//valid
        //                    {
        //                        tracingService.Trace("Serial Number Found i n SAP");
        //                        tracingService.Trace("SAP Response:" + rootObject.EX_PRD_DET.MATNR);
        //                        #region UpdateMaterialCodeonAsset
        //                        string _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                          <entity name='product'>
        //                            <attribute name='name' />
        //                            <attribute name='productid' />
        //                            <attribute name='hil_materialgroup' />
        //                            <attribute name='hil_division' />
        //                            <order attribute='name' descending='false' />
        //                            <filter type='and'>
        //                              <condition attribute='name' operator='eq' value='{rootObject.EX_PRD_DET.MATNR.Trim()}' />
        //                            </filter>
        //                          </entity>
        //                        </fetch>";
        //                        EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                        foreach (Entity ent  in entCol.Entities)
        //                        {
        //                            tracingService.Trace("Division: " + ent.GetAttributeValue<EntityReference>("hil_division").Name);
        //                            tracingService.Trace("Material Group: " + ent.GetAttributeValue<EntityReference>("hil_materialgroup").Name);

        //                            entity["msdyn_product"] = ent.ToEntityReference();
        //                            if (ent.Contains("hil_division"))
        //                                entity["hil_productcategory"] = ent.GetAttributeValue<EntityReference>("hil_division");
        //                            if (ent.Contains("hil_materialgroup"))
        //                                entity["hil_productsubcategory"] = ent.GetAttributeValue<EntityReference>("hil_materialgroup");

        //                        }
        //                        tracingService.Trace("Model: " + entity.GetAttributeValue<EntityReference>("msdyn_product").Id);
        //                        tracingService.Trace("Prod Category: " + entity.GetAttributeValue<EntityReference>("hil_productcategory").Id);
        //                        tracingService.Trace("Prod Subcategory: " + entity.GetAttributeValue<EntityReference>("hil_productsubcategory").Id);
        //                        #endregion
        //                        entity["hil_isserialnumberverified"] = true;//{Asset Status: "SAP Verified"}
        //                    }
        //                    else//Not valid
        //                    {
        //                        entity["hil_isserialnumberverified"] = false;//{Asset Status: "SAP Not Verified"}
        //                        msdyn_customerasset iAsset = entity.ToEntity<msdyn_customerasset>();
        //                        if (iAsset.hil_ProductSubcategory != null)
        //                        {
        //                            Product iDivision = (Product)service.Retrieve(Product.EntityLogicalName, iAsset.hil_ProductSubcategory.Id, new ColumnSet("hil_isverificationrequired", "hil_productregistrationwef"));
        //                            if (iDivision.Attributes.Contains("hil_isverificationrequired"))
        //                            {
        //                                bool isVerified = (bool)iDivision["hil_isverificationrequired"];
        //                                DateTime productRegistrationDateWef = new DateTime(1900, 1, 1);
        //                                DateTime invoicedate = DateTime.Now;
        //                                if (iAsset.hil_InvoiceDate != null)
        //                                {
        //                                    invoicedate = Convert.ToDateTime(iAsset.hil_InvoiceDate);
        //                                }
        //                                if (iDivision.Attributes.Contains("hil_productregistrationwef"))
        //                                {
        //                                    productRegistrationDateWef = (DateTime)iDivision["hil_productregistrationwef"];
        //                                }

        //                                if (isVerified == true)
        //                                {
        //                                    if (invoicedate >= productRegistrationDateWef)
        //                                    {
        //                                        throw new InvalidPluginExecutionException("Asset Serial Number is not valid.");
        //                                    }
        //                                }
        //                            }
        //                        }
        //                        else {
        //                            throw new InvalidPluginExecutionException("Product Subcategory/Model is required.");
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new InvalidPluginExecutionException(ex.Message);
        //    }
        //}
        public static void SerialNumberValidation_SAP(Entity entity, IOrganizationService service)
        {
            try
            {
                String sSerialNumber = entity.GetAttributeValue<String>("msdyn_name");
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
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
                            String sUrl = iobj.hil_Url + sSerialNumber;

                            WebClient webClient = new WebClient();
                            string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
                            webClient.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
                            webClient.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                            var jsonData = webClient.DownloadData(sUrl);

                            DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(SerialNumberValidation));
                            SerialNumberValidation rootObject = (SerialNumberValidation)ser.ReadObject(new MemoryStream(jsonData));

                            if (rootObject.EX_PRD_DET != null)//valid
                            {
                                #region UpdateMaterialCodeonAsset
                                var obj1 = (from _Product in orgContext.CreateQuery<Product>()
                                            where _Product.ProductNumber == rootObject.EX_PRD_DET.MATNR
                                            select new
                                            {
                                                _Product.ProductId,
                                                _Product.hil_MaterialGroup,
                                                _Product.hil_Division
                                            }).Take(1);
                                foreach (var iobj1 in obj1)
                                {
                                    if (iobj1.ProductId != null)
                                        entity["msdyn_product"] = new EntityReference(Product.EntityLogicalName, iobj1.ProductId.Value);
                                    if (iobj1.hil_Division != null)
                                        entity["hil_productcategory"] = new EntityReference(Product.EntityLogicalName, iobj1.hil_Division.Id);
                                    if (iobj1.hil_MaterialGroup != null)
                                        entity["hil_productsubcategory"] = new EntityReference(Product.EntityLogicalName, iobj1.hil_MaterialGroup.Id);
                                }
                                #endregion
                                entity["hil_isserialnumberverified"] = true;//{Asset Status: "SAP Verified"}
                            }
                            else//Not valid
                            {
                                entity["hil_isserialnumberverified"] = false;//{Asset Status: "SAP Not Verified"}
                                msdyn_customerasset iAsset = entity.ToEntity<msdyn_customerasset>();
                                if (iAsset.hil_ProductSubcategory != null)
                                {
                                    Product iDivision = (Product)service.Retrieve(Product.EntityLogicalName, iAsset.hil_ProductSubcategory.Id, new ColumnSet("hil_isverificationrequired", "hil_productregistrationwef"));
                                    if (iDivision.Attributes.Contains("hil_isverificationrequired"))
                                    {
                                        bool isVerified = (bool)iDivision["hil_isverificationrequired"];
                                        DateTime productRegistrationDateWef = new DateTime(1900, 1, 1);
                                        DateTime invoicedate = DateTime.Now;
                                        if (iAsset.hil_InvoiceDate != null)
                                        {
                                            invoicedate = Convert.ToDateTime(iAsset.hil_InvoiceDate);
                                        }
                                        if (iDivision.Attributes.Contains("hil_productregistrationwef"))
                                        {
                                            productRegistrationDateWef = (DateTime)iDivision["hil_productregistrationwef"];
                                        }

                                        if (isVerified == true)
                                        {
                                            if (invoicedate >= productRegistrationDateWef)
                                            {
                                                throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXX - Serial Number is not valid - XXXXXXXXXXXXXXXXXXXXXXXX");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.PreCreate.SerialNumberValidation_SAP" + ex.Message);
            }
        }
        #region Customer Asset Duplicate Check
        public static void IfDuplicate(IOrganizationService service, msdyn_customerasset cAst)
        {
            if (cAst.msdyn_name != null)
            {
                //if (cAst.hil_IsSerialized != null)
                //{
                //    if (cAst.hil_IsSerialized.Value == 1)
                //    {
                //DateTime cOn = Convert.ToDateTime(cAst.CreatedOn);
                QueryExpression Query = new QueryExpression(msdyn_customerasset.EntityLogicalName);
                Query.ColumnSet = new ColumnSet("msdyn_name", "hil_isserialized", "statecode");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition(new ConditionExpression("msdyn_name", ConditionOperator.Equal, cAst.msdyn_name));
                //Query.Criteria.AddCondition(new ConditionExpression("hil_isserialized", ConditionOperator.Equal, 1));
                //Query.Criteria.AddCondition(new ConditionExpression("createdon", ConditionOperator.E));
                EntityCollection Found = service.RetrieveMultiple(Query);
                if (Found.Entities.Count > 0)
                {
                    foreach (msdyn_customerasset iAsset in Found.Entities)
                    {
                        if (iAsset.statecode.Equals(msdyn_customerassetState.Active))
                        {
                            if (cAst.GetAttributeValue<bool>("hil_serialnumberscan") == true)
                            {
                                iAsset.msdyn_name = iAsset.msdyn_name + " -X";
                                iAsset["hil_reviewforduplicateserialnumber"] = true;
                                service.Update(iAsset);
                            }
                            else
                            {
                                throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXX - SERIAL NUMBER ALREADY EXISTS - XXXXXXXXXXXXXXXXXXXXXXXX");
                            }
                        }
                    }
                }
                //    }
                //}
            }
        }
        #endregion
        #region Customer Asset Default Values
        public static void AssetDefaultValues(IOrganizationService service, msdyn_customerasset Asset)
        {
            //if (Asset.Attributes.Contains("hil_source"))
            //{
            //    OptionSetValue source = Asset.GetAttributeValue<OptionSetValue>("hil_source");
            //    if (source.Value == 1 || source.Value == 2)
            //    {
            //        Asset["statuscode"] = new OptionSetValue(910590000);//{Asset Status: "Pending for Approval"}
            //    }
            //}
            //else {
            //    throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - ASSET SOURCE IS REQUIRED - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
            //}

            Asset["hil_createwarranty"] = false;
            if (Asset.hil_productsubcategorymapping != null)
            {
                if (Asset.hil_ProductCategory == null || Asset.hil_ProductSubcategory == null)
                {
                    hil_stagingdivisonmaterialgroupmapping Map = (hil_stagingdivisonmaterialgroupmapping)service.Retrieve(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, Asset.hil_productsubcategorymapping.Id, new ColumnSet(true));
                    if (Map.hil_ProductCategoryDivision != null)
                    {
                        Asset.hil_ProductCategory = Map.hil_ProductCategoryDivision;
                        if (Map.hil_ProductSubCategory != null)
                            Asset.hil_ProductSubcategory = Map.hil_ProductSubCategoryMG;
                        //service.Update(Asset);
                    }
                }
            }
            else
            {
                if (Asset.hil_ProductCategory != null || Asset.hil_ProductSubcategory != null)
                {
                    Guid MapId = GetCatSubCatMappingRec(service, Asset.hil_ProductCategory.Id, Asset.hil_ProductSubcategory.Id);
                    if (MapId != Guid.Empty)
                        Asset.hil_productsubcategorymapping = new EntityReference(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, MapId);
                }
            }
            if (Asset.hil_InvoiceAvailable != null && Asset.hil_InvoiceAvailable == true)
            {
                OptionSetValue _source = null;
                if (Asset.hil_InvoiceDate == null)
                {
                    throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - INVOICE DATE IS REQUIRED - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                }
                if (Asset.Contains("hil_source"))
                {
                    _source = (OptionSetValue)Asset["hil_source"];
                }
                if (_source == null) {
                    throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - SOURCE IS REQUIRED - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                }
                if (_source.Value != 5 && _source.Value != 6)
                {
                    if (Asset.hil_InvoiceNo == null) // VALIDATION LOGIC IS REQUIRED
                    {
                        throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - INVOICE NUMBER IS REQUIRED - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                    }
                    if (Asset.hil_InvoiceValue == null || Asset.hil_InvoiceValue == 0)
                    {
                        throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - INVOICE VALUE IS REQUIRED - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                    }
                    if (!Asset.Attributes.Contains("hil_purchasedfrom"))
                    {
                        throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - PURCHASED FROM IS REQUIRED - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                    }
                    else
                    {
                        if (Asset.GetAttributeValue<string>("hil_purchasedfrom") == null)
                            throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - PURCHASED FROM IS REQUIRED - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                    }
                    if (!Asset.Attributes.Contains("hil_retailerpincode"))
                    {
                        throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - PURCHASED FROM LOCATION IS REQUIRED - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                    }
                    else
                    {
                        if (Asset.GetAttributeValue<string>("hil_retailerpincode") == null)
                            throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - PURCHASED FROM LOCATION IS REQUIRED - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                    }
                }
            }
            if (Asset.hil_InvoiceDate != null)
            {
                DateTime Now = DateTime.Now;
                DateTime dtInvoice = Convert.ToDateTime(Asset.hil_InvoiceDate);
                if(Now.Date < dtInvoice.Date)
                {
                    throw new InvalidPluginExecutionException("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX - INVOICE DATE CAN'T BE GREATER THAN TODAY'S DATE - XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
                }
            }
        }
        public static Guid GetCatSubCatMappingRec(IOrganizationService service, Guid CategoryId, Guid SubCategoryId)
        {
            Guid MapId = new Guid();
            hil_stagingdivisonmaterialgroupmapping iMap = new hil_stagingdivisonmaterialgroupmapping();
            QueryExpression Query = new QueryExpression(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("hil_productcategorydivision", ConditionOperator.Equal, CategoryId));
            Query.Criteria.AddCondition(new ConditionExpression("hil_productsubcategorymg", ConditionOperator.Equal, SubCategoryId));
            EntityCollection Found = service.RetrieveMultiple(Query);
            if(Found.Entities.Count > 0)
            {
                MapId = Found.Entities[0].Id;
            }
            return MapId;
        }
        #endregion

        static void ValidateSerialNumberRegExp(string srno)
        {
            string AlphaRegex = @"^[a-zA-Z0-9]*$";
            if (!System.Text.RegularExpressions.Regex.IsMatch(srno, AlphaRegex))
                throw new InvalidPluginExecutionException("Serial Number is not valid.");
        }
    }
}