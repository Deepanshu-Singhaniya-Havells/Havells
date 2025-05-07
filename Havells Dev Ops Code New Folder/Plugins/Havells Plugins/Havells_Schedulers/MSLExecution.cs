using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Schedulers
{
    public class MSLExecution
    {
        #region MSL 
        public static void InitiateExecution(IOrganizationService service)
        {
            using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
            {
                var obj = from _IConfig in orgContext.CreateQuery<Account>()
                          join _Invoice in orgContext.CreateQuery<hil_inventory>()
                          on _IConfig.AccountId equals _Invoice.hil_OwnerAccount.Id
                          join _mslConfig in orgContext.CreateQuery<hil_minimumstocklevel>()
                          on _IConfig.AccountId equals _mslConfig.hil_Account.Id
                          join _pdt in orgContext.CreateQuery<Product>()
                          on _mslConfig.hil_SparePart.Id equals _pdt.ProductId
                          where (_Invoice.hil_Part.Id == _mslConfig.hil_SparePart.Id) &&
                                (_IConfig.hil_Schedule1 == DateTime.Now.Day || _IConfig.hil_Schedule2 == DateTime.Now.Day) &&
                                (_IConfig.CustomerTypeCode.Value == 5 || _IConfig.CustomerTypeCode.Value == 6 || _IConfig.CustomerTypeCode.Value == 9) &&
                                (_Invoice.hil_AvailableQty < _mslConfig.hil_MSLQuantity) && (_Invoice.hil_InventoryType.Value == 1)
                          select new
                          {
                              _IConfig.hil_InWarrantyCustomerSAPCode,
                              _IConfig.AccountId,
                              _IConfig.OwnerId,
                              _IConfig.CustomerTypeCode,
                              _mslConfig.hil_MSLQuantity,
                              _Invoice.hil_Part,
                              _Invoice.hil_AvailableQty,
                              _pdt.hil_StagingDivision,
                              _pdt.ProductId
                          };
                foreach (var iobj in obj)
                {
                    Product Division = new Product();
                    int _poType = 910590000;
                    EntityReference AccOwner = new EntityReference(SystemUser.EntityLogicalName, iobj.OwnerId.Id);
                    EntityReference Acc = new EntityReference(Account.EntityLogicalName, iobj.AccountId.Value);
                    if (iobj.hil_StagingDivision != null)
                    {
                        Division = GetThisDivision(service, iobj.hil_StagingDivision);
                    }
                    Guid Header = Havells_Plugin.HelperPO.CreatePOHeader(service, new EntityReference(), AccOwner, Acc, new EntityReference(Product.EntityLogicalName, Division.Id), Division.hil_SAPCode, _poType, iobj.hil_InWarrantyCustomerSAPCode, new OptionSetValue(1));
                    if (Header != Guid.Empty)
                    {
                        string DistributionChannel = Havells_Plugin.HelperPO.getDistributionChannel(new OptionSetValue(1), iobj.CustomerTypeCode, service);
                        Guid Line = Havells_Plugin.HelperPO.CreatePO(service, AccOwner.Id, Guid.Empty, Header, (iobj.hil_MSLQuantity.Value - iobj.hil_AvailableQty.Value), iobj.AccountId.Value, iobj.hil_Part.Id, _poType, iobj.CustomerTypeCode, iobj.hil_InWarrantyCustomerSAPCode, DistributionChannel, new OptionSetValue(1));
                        if(Line != Guid.Empty)
                        {
                            hil_productrequest _prReq = (hil_productrequest)service.Retrieve(hil_productrequest.EntityLogicalName, Line, new ColumnSet(false));
                            _prReq["hil_mslquantity"] = (iobj.hil_MSLQuantity.Value - iobj.hil_AvailableQty.Value);
                            _prReq["hil_approvedquantity"] = (iobj.hil_MSLQuantity.Value - iobj.hil_AvailableQty.Value);
                            _prReq.statuscode = new OptionSetValue(910590004);
                            service.Update(_prReq);
                        }
                    }
                }
            }
        }
        #region Get Division
        //public static Product GetThisDivision(IOrganizationService service, string StageDiv)
        //{
        //    Product Division = new Product();
        //    QueryByAttribute Query = new QueryByAttribute(Product.EntityLogicalName);
        //    Query.AddAttributeValue("hil_stagingdivision", StageDiv);
        //    Query.ColumnSet = new ColumnSet(false);
        //    EntityCollection Found = service.RetrieveMultiple(Query);
        //    if (Found.Entities.Count > 0)
        //    {
        //        Division = (Product)Found.Entities[0];
        //    }
        //    return Division;
        //}
        #endregion
        #endregion
    }
}