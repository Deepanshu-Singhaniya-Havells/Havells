using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
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
    public class PostUpdate : IPlugin
    {
        
        #region PluginConfig
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "msdyn_customerasset"
                    && context.MessageName.ToUpper() == "UPDATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    msdyn_customerasset Asset = entity.ToEntity<msdyn_customerasset>();
                    if (Asset.hil_InvoiceAvailable != null) // Invoice Available edited
                    {
                        if (Asset.hil_InvoiceAvailable == true)
                        {
                            Entity entityCustAsset = service.Retrieve("msdyn_customerasset", entity.Id, new ColumnSet("msdyn_name", "msdyn_product", "hil_productcategory", "hil_productsubcategory", "hil_isserialnumberverified"));
                            Product iSCat = (Product)service.Retrieve(Product.EntityLogicalName, entityCustAsset.GetAttributeValue<EntityReference>("hil_productsubcategory").Id, new ColumnSet("hil_isserialized"));
                            if (iSCat.hil_IsSerialized != null && iSCat.hil_IsSerialized.Value == 1)
                            {
                                SerialNumberValidation_SAP(entityCustAsset, service);
                            }
                        }
                    }
                    if (Asset.hil_CreateWarranty != null)
                    {
                        if (Asset.hil_CreateWarranty == true)
                        {
                            //HelperWarrantyModule.Init_Warranty(Asset.Id, service);
                            WarrantyEngine.ExecuteWarrantyEngine(service, entity.Id);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.CustomerAsset_PostCreate.CustomerAssestConfig" + ex.Message);
            }
            #endregion
        }
        public static void SerialNumberValidation_SAP(Entity entity, IOrganizationService service)
        {
            try
            {
                Entity ent = new Entity(entity.LogicalName, entity.Id);
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
                                        ent["msdyn_product"] = new EntityReference(Product.EntityLogicalName, iobj1.ProductId.Value);
                                    if (iobj1.hil_Division != null)
                                        ent["hil_productcategory"] = new EntityReference(Product.EntityLogicalName, iobj1.hil_Division.Id);
                                    if (iobj1.hil_MaterialGroup != null)
                                        ent["hil_productsubcategory"] = new EntityReference(Product.EntityLogicalName, iobj1.hil_MaterialGroup.Id);
                                }
                                #endregion
                                ent["hil_isserialnumberverified"] = true;//{Asset Status: "SAP Verified"}
                            }
                            else//Not valid
                            {
                                ent["hil_isserialnumberverified"] = false;//{Asset Status: "SAP Not Verified"}
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
                            service.Update(ent);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.PreCreate.SerialNumberValidation_SAP" + ex.Message);
            }
        }
        //public static void CreateWorkOrder(IOrganizationService service, Guid AsstId)
        //{
        //    msdyn_customerasset CustAsst = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, AsstId, new ColumnSet(true));
        //    msdyn_workorder Wo = new msdyn_workorder();
        //    Wo.hil_CustomerRef = CustAsst.hil_Customer;
        //    Wo.msdyn_CustomerAsset = new EntityReference(msdyn_customerasset.EntityLogicalName, CustAsst.Id);
        //    Wo.msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, new Guid("217d1e71-6e7b-e811-a95d-000d3af05df5"));
        //    Wo.hil_Productcategory = CustAsst.hil_ProductCategory;
        //    Wo.hil_CallSubType = new EntityReference(hil_callsubtype.EntityLogicalName, new Guid("41b1a66a-717f-e811-a95c-000d3af075c4"));
        //    Wo.hil_CallType = new EntityReference(hil_calltype.EntityLogicalName, new Guid("84909f31-b673-e811-a95a-000d3af068d4"));
        //    Wo.hil_natureofcomplaint = new EntityReference(hil_natureofcomplaint.EntityLogicalName, new Guid("9745e6a7-d084-e811-a95c-000d3af075c4"));
        //    service.Create(Wo);
        //}
        #endregion
    }
}