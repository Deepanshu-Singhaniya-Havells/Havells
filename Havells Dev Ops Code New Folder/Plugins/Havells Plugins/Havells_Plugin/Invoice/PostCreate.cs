using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System.Linq;
using Microsoft.Xrm.Sdk.Query;

namespace Havells_Plugin.Invoice
{
    public class PostCreate : IPlugin
    {
        private static ITracingService tracingService = null;
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
                    && context.PrimaryEntityName.ToLower() == hil_invoice.EntityLogicalName && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    CreateGrnLine(entity, service);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Invoice.PostCreate.Execute" + ex.Message);
            }
        }

        public static void CreateGrnLine(Entity entity, IOrganizationService service)
        {
            try
            {
                Guid fsProductCodeId = Guid.Empty;
                Guid fsAccount = Guid.Empty;
                Guid fsOwnerId = Guid.Empty;
                Guid fsPOLineId = Guid.Empty;
                Guid fsPOHeaderId = Guid.Empty;
                OptionSetValue opWrtyStatus = null;
                hil_invoice Invoice = entity.ToEntity<hil_invoice>();
                OrganizationServiceContext orgContext = new OrganizationServiceContext(service);

                if (Invoice.hil_ProductCode != null)
                {
                    fsProductCodeId = Helper.GetGuidbyNameInvoiceOnly(Product.EntityLogicalName, "name", Invoice.hil_ProductCode, service);
                }

                #region GetAccountandOwner
                var obj = (from _PO in orgContext.CreateQuery<hil_productrequestheader>()
                           join _Account in orgContext.CreateQuery<Account>()
                           on _PO.hil_Account.Id equals _Account.Id
                           where _PO.hil_SalesOrderNo == Invoice.hil_Salesordernumber
                           select new
                           {
                               _PO.hil_Account//account
                               ,
                               _Account.OwnerId
                           }).Take(1);

                foreach (var iobj in obj)
                {
                    if (iobj.hil_Account != null)
                    {
                        fsAccount = iobj.hil_Account.Id;
                    }
                    if (iobj.OwnerId != null)
                    {
                        fsOwnerId = iobj.OwnerId.Id;
                    }
                }
                #endregion

                var obj1 = (from _PO in orgContext.CreateQuery<hil_productrequestheader>()
                            join _Account in orgContext.CreateQuery<Account>()
                            on _PO.hil_Account.Id equals _Account.Id
                            join _POLine in orgContext.CreateQuery<hil_productrequest>()
                            on _PO.hil_productrequestheaderId.Value equals _POLine.hil_PRHeader.Id
                            where _PO.hil_SalesOrderNo == Invoice.hil_Salesordernumber
                            && _POLine.hil_PartCode.Id == fsProductCodeId
                            select new
                            {
                                _PO.hil_Account//account
                               ,
                                _Account.OwnerId
                               ,
                                _POLine.hil_productrequestId,
                                _POLine.hil_WarrantyStatus
                               ,
                                _PO.hil_productrequestheaderId
                            }).Take(1);
                foreach (var iobj in obj1)
                {
                    if (iobj.hil_productrequestheaderId != Guid.Empty)
                        fsPOHeaderId = iobj.hil_productrequestheaderId.Value;
                    if (iobj.hil_productrequestId != Guid.Empty)
                        fsPOLineId = iobj.hil_productrequestId.Value;
                    if (iobj.hil_WarrantyStatus != null)
                        opWrtyStatus = iobj.hil_WarrantyStatus;

                    hil_grnline crGrnLine = new hil_grnline();
                    if (Invoice.hil_Quantity != null)
                        crGrnLine.hil_Quantity = Invoice.hil_Quantity.Value;
                    if (fsProductCodeId != Guid.Empty)
                        crGrnLine.hil_ProductCode = new EntityReference(Product.EntityLogicalName, fsProductCodeId);
                    if (Invoice.hil_Salesordernumber != null)
                        crGrnLine.hil_SoNumber = Invoice.hil_Salesordernumber;
                    if (opWrtyStatus != null)
                    {
                        crGrnLine.hil_WarrantyStatus = opWrtyStatus;
                    }
                    if (fsAccount != Guid.Empty)
                    {
                        crGrnLine.hil_Account = new EntityReference(Account.EntityLogicalName, fsAccount);
                    }
                    if (fsOwnerId != Guid.Empty)
                    {
                        crGrnLine.OwnerId = new EntityReference(SystemUser.EntityLogicalName, fsOwnerId);
                    }

                    if (fsPOHeaderId != Guid.Empty)
                    {
                        crGrnLine["hil_prheader"] = new EntityReference(hil_productrequestheader.EntityLogicalName, fsPOHeaderId);
                    }
                    if (fsPOLineId != Guid.Empty)
                    {
                        crGrnLine.hil_ProductRequest = new EntityReference(hil_productrequest.EntityLogicalName, fsPOLineId);
                    }
                    if (Invoice.hil_name != null)
                    {
                        crGrnLine["hil_invoicenumber"] = Invoice.hil_name;
                    }
                    service.Create(crGrnLine);
                    //tag grn line to invoice
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Invoice.PostCreate.CreateGrnLine" + ex.Message);
            }

        }
    }
}
