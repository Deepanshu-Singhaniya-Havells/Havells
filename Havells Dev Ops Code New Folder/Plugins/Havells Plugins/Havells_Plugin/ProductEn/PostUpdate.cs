using System;
using System.Linq;
using Microsoft.Crm.Sdk.Messages;
using System.Text;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk;
using System.Net;
using System.IO;
// Microsoft Dynamics CRM namespace(s)
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;

namespace Havells_Plugin.ProductEn
{
    public class PostUpdate : IPlugin
    {

        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
             //#region MainRegion
            try
            {
                if (context.InputParameters.Contains("Target")
                    && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == Product.EntityLogicalName
                    && context.MessageName.ToUpper() == "UPDATE")
                {
                    OrganizationServiceContext orgContext = new OrganizationServiceContext(service);
                    Entity entity = (Entity)context.InputParameters["Target"];
                    if (entity.Contains("hil_resolve"))
                    {
                        if (entity.GetAttributeValue<OptionSetValue>("hil_resolve").Value == 1)
                        {
                            ResolveDataProductMaterialFromPlugin(service, entity.Id);

                            //Product enProduct = (Product)service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("productnumber"));
                            //Decimal price= GetPriceOfMaterial(enProduct.ProductNumber);
                            //enProduct.hil_Amount = new Money(price);
                            //service.Update(enProduct);

                        }
                    }

                    
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("" + ex.Message);
            }
        }

        public static Decimal GetPriceOfMaterial(String MaterialCode)
        {
            Decimal result = 0;
            try
            {

                {
                    WebRequest request = WebRequest.Create("https://hsappoqa.havells.com:50001/RESTAdapter/MDMService/Core/Product/GetPriceDTLMaster");
                    request.Method = "POST";

                    MaterialPrice priceReq = new MaterialPrice();
                    priceReq.Condition = "ZWEB";//MRP
                    priceReq.FromDate = "19900103000000";
                    priceReq.IsInitialLoad = 0;
                    priceReq.MaterialCode = MaterialCode;
                    priceReq.ToDate = DateTime.Now.ToString("yyyyMMddHHmmss"); ;

                    String postData = JsonConvert.SerializeObject(priceReq);


                    byte[] byteArray = Encoding.UTF8.GetBytes(postData);
                    // Set the ContentType property of the WebRequest.  
                    request.ContentType = "application/x-www-form-urlencoded";
                    // Set the ContentLength property of the WebRequest.  
                    request.ContentLength = byteArray.Length;
                    // Get the request stream.  
                    Stream dataStream = request.GetRequestStream();
                    // Write the data to the request stream.  
                    dataStream.Write(byteArray, 0, byteArray.Length);
                    // Close the Stream object.  
                    dataStream.Close();
                    // Get the response.  
                    WebResponse response = request.GetResponse();
                    // Display the status. 

                    //  tracingService.Trace("8" + ((HttpWebResponse)response).StatusDescription);

                    Console.WriteLine(((HttpWebResponse)response).StatusDescription);
                    // Get the stream containing content returned by the server.  
                    dataStream = response.GetResponseStream();
                    // Open the stream using a StreamReader for easy access.  
                    StreamReader reader = new StreamReader(dataStream);
                    // Read the content.  
                    string responseFromServer = reader.ReadToEnd();


                    MaterialPriceResponseRoot mterialPriceResponseRootobj = JsonConvert.DeserializeObject<MaterialPriceResponseRoot>(responseFromServer);


                    foreach (MaterialPriceResponseObj obj in mterialPriceResponseRootobj.Results)
                    {
                        result = Convert.ToDecimal(obj.KBETR);
                    }


                }

            }
            catch (Exception ex)
            {
                // throw new InvalidPluginExecutionException("Havells_Plugin.CustomerAsset.PostCreate.GetChildAsset" + ex.Message);
            }
            return result;
        }


        public static void ResolveDataProductMaterialFromPlugin(IOrganizationService service, Guid fsProductId)
        {
            using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
            {
                var obj = from _Product in orgContext.CreateQuery<Product>()
                           where _Product.hil_hierarchylevel.Value == 5//material
                           && _Product.ProductId.Value == fsProductId
                           select new { _Product };
                Int32 iLeftCount = Enumerable.Count(obj);
                Int32 idoneCount = 0;
                foreach (var iobj in obj)
                {
                    try
                    {
                        String sDivision = String.Empty;
                        String sMG = String.Empty;
                        String sMG1 = String.Empty;
                        String sMG2 = String.Empty;
                        String sMG3 = String.Empty;
                        String sMG4 = String.Empty;
                        String sMG5 = String.Empty;
                        String sSBU = String.Empty;
                        String sSAPCode = String.Empty;
                        String sHSN = String.Empty;
                        String sMPG = String.Empty;

                        Guid fsDivision = Guid.Empty;
                        Guid fsMG = Guid.Empty;
                        Guid fsMG1 = Guid.Empty;
                        Guid fsMG2 = Guid.Empty;
                        Guid fsMG3 = Guid.Empty;
                        Guid fsMG4 = Guid.Empty;
                        Guid fsMG5 = Guid.Empty;
                        Guid fsSBU = Guid.Empty;
                        Guid fsSAPCode = Guid.Empty;
                        Guid fsHSN = Guid.Empty;
                        Guid fsMPG = Guid.Empty;
                        Decimal MRP = 0;

                       // if (iobj._Product.ProductNumber != null)
                           // MRP = GetPriceOfMaterial(iobj._Product.ProductNumber);

                        if (iobj._Product.hil_StagingDivision != null)
                            sDivision = iobj._Product.hil_StagingDivision;

                        if (iobj._Product.hil_StagingMaterialGroup != null)
                            sMG = iobj._Product.hil_StagingMaterialGroup;

                        if (iobj._Product.hil_StagingMaterialGroup1 != null)
                            sMG1 = iobj._Product.hil_StagingMaterialGroup1;

                        if (iobj._Product.hil_StagingMaterialGroup2 != null)
                            sMG2 = iobj._Product.hil_StagingMaterialGroup2;

                        if (iobj._Product.hil_StagingMaterialGroup3 != null)
                            sMG3 = iobj._Product.hil_StagingMaterialGroup3;

                        if (iobj._Product.hil_StagingMaterialGroup4 != null)
                            sMG4 = iobj._Product.hil_StagingMaterialGroup4;

                        if (iobj._Product.hil_StagingMaterialGroup5 != null)
                            sMG5 = iobj._Product.hil_StagingMaterialGroup5;

                        if (iobj._Product.hil_StagingSBU != null)
                            sSBU = iobj._Product.hil_StagingSBU;

                        if (iobj._Product.hil_StagingHSNCode != null)
                            sHSN = iobj._Product.hil_StagingHSNCode;

                        if (iobj._Product.hil_SAPCode != null)
                            sSAPCode = iobj._Product.hil_SAPCode;

                        if (iobj._Product.hil_StagingMPGCode != null)
                            sMPG = iobj._Product.hil_StagingMPGCode;


                        if (sSBU != String.Empty)
                        {
                            fsSBU = GetGuidbyNameProducts(Product.EntityLogicalName, "hil_uniquekey", sSBU, 1, service);
                        }

                        if (sDivision != String.Empty)
                        {
                            //division hierarchy level 2
                            //fsDivision = GetGuidbyNameProducts(Product.EntityLogicalName, "hil_stagingdivision", sDivision, 2, service);
                            fsDivision = GetGuidbyNameProducts(Product.EntityLogicalName, "hil_uniquekey", sDivision, 2, service, 0, true);
                        }
                        if (sMG != String.Empty)
                        {
                            //MG hierarchy level 3
                            fsMG = GetGuidbyNameProducts(Product.EntityLogicalName, "hil_uniquekey", sMG, 3, service);
                        }

                        if (sMG1 != String.Empty)
                            fsMG1 = Helper.GetGuidbyName(Product.EntityLogicalName, "hil_uniquekey", sMG1, service, -1);

                        if (sMG2 != String.Empty)
                            fsMG2 = Helper.GetGuidbyName(hil_materialgroup2.EntityLogicalName, "hil_code", sMG2, service, -1);
                        if (sMG3 != String.Empty)
                            fsMG3 = Helper.GetGuidbyName(hil_materialgroup3.EntityLogicalName, "hil_code", sMG3, service, -1);
                        if (sMG4 != String.Empty)
                            fsMG4 = Helper.GetGuidbyName(hil_materialgroup4.EntityLogicalName, "hil_code", sMG4, service, -1);
                        if (sMG5 != String.Empty)
                            fsMG5 = Helper.GetGuidbyName(hil_materialgroup5.EntityLogicalName, "hil_code", sMG5, service, -1);

                        if (sMPG != null && sMPG != String.Empty)
                        {
                            sMPG = iobj._Product.hil_StagingMPGCode;
                            sHSN = sHSN + sMPG;
                        }
                        if (sHSN != "" && sHSN != String.Empty)
                        {
                            fsHSN = Helper.GetGuidbyName(hil_hsncode.EntityLogicalName, "hil_name", sHSN, service, -1);
                          
                        }

                        Product upProduct = new Product();

                        upProduct.ProductId = iobj._Product.ProductId;

                        if (fsDivision != Guid.Empty)
                            upProduct.hil_Division = new EntityReference(Product.EntityLogicalName, fsDivision);
                        if (fsMG != Guid.Empty)
                            upProduct.hil_MaterialGroup = new EntityReference(Product.EntityLogicalName, fsMG);
                        if (fsMG1 != Guid.Empty)
                            upProduct.hil_MaterialGroup1 = new EntityReference(Product.EntityLogicalName, fsMG1);

                        if (fsMG2 != Guid.Empty)
                            upProduct.hil_MaterialGroup2 = new EntityReference(hil_materialgroup2.EntityLogicalName, fsMG2);

                        if (fsMG3 != Guid.Empty)
                            upProduct.hil_MaterialGroup3 = new EntityReference(hil_materialgroup3.EntityLogicalName, fsMG3);

                        if (fsMG4 != Guid.Empty)
                            upProduct.hil_MaterialGroup4 = new EntityReference(hil_materialgroup4.EntityLogicalName, fsMG4);

                        if (fsMG5 != Guid.Empty)
                            upProduct.hil_MaterialGroup5 = new EntityReference(hil_materialgroup5.EntityLogicalName, fsMG5);

                        if (fsSBU != Guid.Empty)
                        {
                            upProduct.hil_SBU = new EntityReference(Product.EntityLogicalName, fsSBU);
                        }
                        if (fsHSN != Guid.Empty)
                        {
                            upProduct.hil_HSNcode = new EntityReference(hil_hsncode.EntityLogicalName, fsHSN);
                        }
                        //   upProduct.hil_Amount = new Money(MRP);

                        upProduct["hil_resolve"] = new OptionSetValue(2);

                        service.Update(upProduct);

                        SetStateRequest state = new SetStateRequest();
                        state.State = new OptionSetValue(0);
                        state.Status = new OptionSetValue(1);
                        state.EntityMoniker = new EntityReference(Product.EntityLogicalName, iobj._Product.ProductId.Value);
                        SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                       // Console.WriteLine("Record Processed: " + ++idoneCount + " Left: " + --iLeftCount);
                    }
                    catch (Exception ex)
                    {
                        --idoneCount;
                        Console.WriteLine("Error at " + ++idoneCount);
                    }
                }
            }
        }
        public static Guid GetGuidbyNameProducts(String sEntityName, String sFieldName, String sFieldValue, Int32 iHierarchyLevel, IOrganizationService service, int iStatusCode = 0, Boolean checkCOntains = false)
        {
            Guid fsResult = Guid.Empty;
            try
            {
                QueryExpression qe = new QueryExpression(sEntityName);
                if (checkCOntains)
                    qe.Criteria.AddCondition(sFieldName, ConditionOperator.BeginsWith, sFieldValue);
                else
                    qe.Criteria.AddCondition(sFieldName, ConditionOperator.Equal, sFieldValue);
                qe.AddOrder("createdon", OrderType.Descending);
                qe.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, iHierarchyLevel);
                if (iStatusCode >= 0)
                {
                    // qe.Criteria.AddCondition("statecode", ConditionOperator.Equal, iStatusCode);
                }
                qe.Criteria.AddCondition("hil_deleteflag", ConditionOperator.NotEqual, 1);
                EntityCollection enColl = service.RetrieveMultiple(qe);
                if (enColl.Entities.Count > 0)
                {
                    fsResult = enColl.Entities[0].Id;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Helper.GetGuidbyName" + ex.Message);
            }
            return fsResult;
        }




    }
}