using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System.Linq;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;

namespace Havells_Plugin.ProductRequest
{
    public class PreCreate : IPlugin
    {
        static ITracingService tracingService = null;
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == hil_productrequest.EntityLogicalName.ToLower()
                    && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace("1");
                    PopulateData1(entity,service);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ProductRequest.PreCreate.Execute" + ex.Message);
            }
        }

        public static void PopulateData(Entity entity, IOrganizationService service)
        {
            try
            {
                //tracingService.Trace("2");
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    //tracingService.Trace("3");
                    hil_productrequest PO = entity.ToEntity<hil_productrequest>();
                    if (PO.hil_PartCode != null)
                    {
                        //tracingService.Trace("4 --> "+ PO.hil_productrequestId.Value);
                        //tracingService.Trace("4.1 --> " + PO.hil_PartCode.Id);
                        var obj = from _PartCode in orgContext.CreateQuery<Product>()
                                  join _Division in orgContext.CreateQuery<Product>() on _PartCode.hil_Division.Id equals _Division.Id
                                  where _PartCode.ProductId.Value == PO.hil_PartCode.Id
                                  select new
                                  {
                                      _PartCode.ProductNumber,
                                      _PartCode.hil_Division,
                                      _Division.hil_SAPCode
                                  };
                        //tracingService.Trace("5"+Enumerable.Count(obj));
                        foreach (var iobj in obj)
                        {
                            PO.hil_ProductCodeValue = iobj.ProductNumber;
                            PO.hil_Division = iobj.hil_Division;
                            PO.hil_DivisionSapCode = iobj.hil_SAPCode;
                            if(PO.hil_WarrantyStatus != null)
                            {
                                //tracingService.Trace("7-> " + PO.hil_WarrantyStatus.Value);
                                string POtypeCode = GetSAPPOTypeCode(service, PO.hil_WarrantyStatus.Value, PO.hil_CustomerSAPCode);
                                tracingService.Trace("8-> " + POtypeCode);
                                PO.hil_POTypeCode = POtypeCode;
                            }
                            //tracingService.Trace("6-> " + iobj.hil_SAPCode);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ProductRequest.PreCreate.PopulateData" + ex.Message);
            }
        }
        public static void PopulateData1(Entity entity, IOrganizationService service)
        {
            try
            {
                tracingService.Trace("2");
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    tracingService.Trace("3");
                    hil_productrequest PO = entity.ToEntity<hil_productrequest>();
                    if (PO.hil_PartCode != null)
                    {
                        if (PO.hil_Job != null)
                        {
                            String ProductNumber = PO.hil_PartCode.Name;
                            var obj2 = from _Product in orgContext.CreateQuery<Product>()
                                       where _Product.ProductId.Value == PO.hil_PartCode.Id
                                       select new { _Product.ProductNumber };
                            foreach (var obj in obj2)
                            {
                                if (obj.ProductNumber != null)
                                    ProductNumber = obj.ProductNumber;
                                break;
                            }

                            tracingService.Trace("4 --> " + PO.hil_productrequestId.Value);
                            tracingService.Trace("4.1 --> " + PO.hil_PartCode.Id);
                            tracingService.Trace("4.2 --> " + PO.hil_PartCode.Name);
                            var obj1 = from _DIvision in orgContext.CreateQuery<Product>()
                                       join _Job in orgContext.CreateQuery<msdyn_workorder>()
                                       on _DIvision.ProductId.Value equals _Job.hil_Productcategory.Id
                                       where _Job.msdyn_workorderId.Value == PO.hil_Job.Id
                                       select new
                                       {
                                           _DIvision.hil_SAPCode,
                                           _DIvision.ProductId
                                       };
                            foreach (var iobj in obj1)
                            {
                                PO.hil_ProductCodeValue = ProductNumber;
                                PO.hil_Division = new EntityReference(Product.EntityLogicalName, iobj.ProductId.Value);
                                PO.hil_DivisionSapCode = iobj.hil_SAPCode;
                                if (PO.hil_WarrantyStatus != null)
                                {
                                    tracingService.Trace("7-> " + PO.hil_WarrantyStatus.Value);
                                    string POtypeCode = GetSAPPOTypeCode(service, PO.hil_WarrantyStatus.Value, PO.hil_CustomerSAPCode);//
                                    tracingService.Trace("8-> " + POtypeCode);
                                    PO.hil_POTypeCode = POtypeCode;
                                }
                                tracingService.Trace("6-> " + iobj.hil_SAPCode);

                            }
                        }
                        else
                        {
                            var obj = from _PartCode in orgContext.CreateQuery<Product>()
                                      join _Division in orgContext.CreateQuery<Product>() on _PartCode.hil_Division.Id equals _Division.Id
                                      where _PartCode.ProductId.Value == PO.hil_PartCode.Id
                                      select new
                                      {
                                          _PartCode.ProductNumber,
                                          _PartCode.hil_Division,
                                          _Division.hil_SAPCode
                                      };
                            tracingService.Trace("5" + Enumerable.Count(obj));
                            foreach (var iobj in obj)
                            {
                                PO.hil_ProductCodeValue = iobj.ProductNumber;
                                PO.hil_Division = iobj.hil_Division;
                                PO.hil_DivisionSapCode = iobj.hil_SAPCode;
                                if (PO.hil_WarrantyStatus != null)
                                {
                                    tracingService.Trace("7-> " + PO.hil_WarrantyStatus.Value);
                                    string POtypeCode = GetSAPPOTypeCode(service, PO.hil_WarrantyStatus.Value, PO.hil_CustomerSAPCode);
                                    tracingService.Trace("8-> " + POtypeCode);
                                    PO.hil_POTypeCode = POtypeCode;
                                }
                                tracingService.Trace("6-> " + iobj.hil_SAPCode);
                            }
                        
                        }
                    }
                
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.ProductRequest.PreCreate.PopulateData" + ex.Message);
            }
        }


        public static string GetSAPPOTypeCode(IOrganizationService service, int WrtyCode, string CustType)
        {
            string _inPOCode = string.Empty;
            try
            {
                QueryByAttribute Query = new QueryByAttribute(hil_integrationconfiguration.EntityLogicalName);
                Query.ColumnSet = new ColumnSet("hil_potypecode");
                if (WrtyCode == 1)
                {
                    Query.AddAttributeValue("hil_name", "InWarrantySAPPOTypeCode");
                }
                else if(WrtyCode == 2 && CustType.StartsWith("E"))
                {
                    Query.AddAttributeValue("hil_name", "InWarrantySAPPOTypeCode");
                }
                else
                {
                    Query.AddAttributeValue("hil_name", "OutWarrantySAPPOTypeCode");
                }
                EntityCollection Found = service.RetrieveMultiple(Query);
                if(Found.Entities.Count > 0)
                {
                    hil_integrationconfiguration _inConf = (hil_integrationconfiguration)Found.Entities[0];
                    _inPOCode = _inConf.hil_POTypeCode;
                }
            }
            catch(Exception ex)
            {
                throw new InvalidPluginExecutionException("ProductRequest.PreCreate.GetSAPPOTypeCode : " + ex.Message);
            }
            return _inPOCode;
        }
    }
}
