using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;

namespace Havells_Plugin
{
    public class HelperCallAllocationRouting
    {
        public static void CallAllocationRoute(msdyn_workorderincident WrkInc, IOrganizationService service)
        {
            if(WrkInc.Attributes.Contains("msdyn_customerasset"))
            {
                try
                {
                    using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                    {
                        var obj = from _WoIn in orgContext.CreateQuery<msdyn_workorderincident>()
                                  join _Wo in orgContext.CreateQuery<msdyn_workorder>() on _WoIn.msdyn_WorkOrder.Id equals _Wo.Id
                                  join _CustAst in orgContext.CreateQuery<msdyn_customerasset>() on _Wo.msdyn_CustomerAsset.Id equals _CustAst.Id
                                  join _Pdt in orgContext.CreateQuery<Product>() on _CustAst.msdyn_Product.Id equals _Pdt.Id
                                  where _WoIn.msdyn_WorkOrder != null
                                  && _Wo.msdyn_CustomerAsset != null
                                  && _CustAst.msdyn_Product != null
                                  && _Pdt.hil_Division != null
                                  select new
                                  {
                                      _WoIn.msdyn_WorkOrder,
                                      _Wo.msdyn_CustomerAsset,
                                      _CustAst.msdyn_Product,
                                      _Pdt.hil_Division
                                  };
                        foreach (var iobj in obj)
                        {
                            Guid WorkOrder = new Guid();
                            Guid CustomerAsset = new Guid();
                            Guid Product = new Guid();
                            Guid Division = new Guid();
                            Guid PinCode = new Guid();
                            if (iobj.msdyn_WorkOrder != null)
                            {
                                WorkOrder = iobj.msdyn_WorkOrder.Id;
                            }
                            if(iobj.msdyn_CustomerAsset != null)
                            {
                                CustomerAsset = iobj.msdyn_CustomerAsset.Id;
                            }
                            if(iobj.msdyn_Product != null)
                            {
                                Product = iobj.msdyn_Product.Id;
                            }
                            if(iobj.hil_Division != null)
                            {
                                Division = iobj.hil_Division.Id;
                            }
                            PinCode = FindContactGuId(WorkOrder, service);
                            if((PinCode.ToString() != "00000000-0000-0000-0000-000000000000") && 
                                (Division.ToString() != "00000000-0000-0000-0000-000000000000"))
                            {
                                AssignWorkOrderToBranchHead(PinCode, Division, WorkOrder, service);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException("HelperCallAllocationRouting_CallAllocationRoute_" +ex.Message);
                }
            }
        }
        public static Guid FindContactGuId(Guid WorkOrder, IOrganizationService service)
        {
            Guid PinCode = new Guid();
            EntityReference Cont = new EntityReference();
            Contact ContEt = new Contact();
            EntityReference PnCode = new EntityReference(hil_pincode.EntityLogicalName);
            msdyn_workorder WrkOrd = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, WorkOrder, new ColumnSet("hil_customerref"));
            if(WrkOrd.Attributes.Contains("hil_customerref"))
            {
                Cont = (EntityReference)WrkOrd.hil_CustomerRef;
                if(Cont.LogicalName == Contact.EntityLogicalName)
                {
                    ContEt = (Contact)service.Retrieve(Contact.EntityLogicalName, Cont.Id, new ColumnSet("hil_pincode"));
                    if (ContEt.Attributes.Contains("hil_pincode"))
                    {
                        PnCode = (EntityReference)ContEt.hil_pincode;
                        PinCode = PnCode.Id;
                    }
                }
            }
            return PinCode;
        }
        public static void AssignWorkOrderToBranchHead(Guid PinCode, Guid Division, Guid WO, IOrganizationService service)
        {
            QueryByAttribute Query = new QueryByAttribute(hil_assignmentmatrix.EntityLogicalName);
            ColumnSet Col = new ColumnSet(true);
            Query.ColumnSet = Col;
            Query.AddAttributeValue("hil_pincode", PinCode);
            Query.AddAttributeValue("hil_division", Division);
            //review active check
            EntityCollection Found = service.RetrieveMultiple(Query);
            if(Found.Entities.Count > 0)  //review condition should never be == 1
            {
                foreach(hil_assignmentmatrix Mtrx in Found.Entities)
                {
                    EntityReference BrnHd = new EntityReference(Account.EntityLogicalName);
                    EntityReference BrnHdOwner = new EntityReference(SystemUser.EntityLogicalName);
                    if(Mtrx.Attributes.Contains("hil_branchhead"))
                    {
                        BrnHd = (EntityReference)Mtrx.hil_BranchHead;
                        Account BranchHead = (Account)service.Retrieve(Account.EntityLogicalName, BrnHd.Id, new ColumnSet("ownerid"));
                        BrnHdOwner = (EntityReference)BranchHead.OwnerId;
                        Helper.Assign(SystemUser.EntityLogicalName, msdyn_workorder.EntityLogicalName, BrnHdOwner.Id, WO, service);                      //Assign
                        break;
                    }
                }
            }
        }
    }
}