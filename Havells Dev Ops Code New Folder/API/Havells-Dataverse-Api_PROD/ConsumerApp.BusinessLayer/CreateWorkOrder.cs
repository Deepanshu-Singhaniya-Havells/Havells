using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using ConsumerApp.BusinessLayer;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class CreateWorkOrder
    {
        [DataMember]
        public string Asset { get; set; }
        [DataMember]
        public string Method { get; set; }
        [DataMember]
        public string ContGuid { get; set; }
        [DataMember]
        public string ProductSubCategory { get; set; }
        [DataMember]
        public string IncidentType { get; set; }
        [DataMember(IsRequired = false)]
        public string NatureOfComplaint { get; set; }
        [DataMember(IsRequired = false)]
        public bool status { get; set; }
        [DataMember]
        public int FileType { get; set; }
        [DataMember(IsRequired = false)]
        public string ErrCode { get; set; }
        [DataMember(IsRequired = false)]
        public string ErrDesc { get; set; }
        [DataMember(IsRequired = false)]
        public string WorkOrderID { get; set; }
        [DataMember(IsRequired = false)]
        public string CustomerId { get; set; }
        public string CustomerAddress { get; set; }
        [DataMember]
        public string ServiceAttachment { get; set; }


        public CreateWorkOrder SubmitServiceRequest(CreateWorkOrder bridge)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();//get org service obj for connection
            try
            {
                if (bridge.Method == "CreateWOService")
                {
                    Contact Cont = (Contact)service.Retrieve(Contact.EntityLogicalName, new Guid(bridge.ContGuid), new ColumnSet("mobilephone", "emailaddress1"));
                    msdyn_workorder iWrkOd = new msdyn_workorder();
                    Guid PriceList = GetGuidbyName(PriceLevel.EntityLogicalName, "name", "Default Price List", service);
                    Guid Nature = GetGuidbyName(hil_natureofcomplaint.EntityLogicalName, "hil_name", "Service", service,0);
                    Guid IncidentType1 = GetGuidbyName(msdyn_incidenttype.EntityLogicalName, "msdyn_name", "Service", service);
                    Guid ServiceAccount = GetGuidbyName(Account.EntityLogicalName, "name", "Dummy Customer", service);
                    iWrkOd.hil_CustomerRef = new EntityReference(Contact.EntityLogicalName, new Guid(bridge.ContGuid));
                    iWrkOd.hil_Email = Cont.EMailAddress1;
                    iWrkOd.hil_mobilenumber = Cont.MobilePhone;
                    iWrkOd.msdyn_CustomerAsset = new EntityReference(msdyn_customerasset.EntityLogicalName, new Guid(bridge.Asset));
                    msdyn_customerasset iAsst = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, new Guid(bridge.Asset), new ColumnSet("msdyn_product", "hil_productsubcategory", "hil_productsubcategorymapping", "hil_productcategory"));
                    if (iAsst.hil_ProductCategory != null)
                    {
                        iWrkOd.hil_Productcategory = iAsst.hil_ProductCategory;
                    }
                    if (iAsst.hil_ProductSubcategory != null)
                    {
                        iWrkOd.hil_ProductSubcategory = iAsst.hil_ProductSubcategory;
                    }
                    if (iAsst.hil_productsubcategorymapping != null)
                    {
                        iWrkOd.hil_ProductCatSubCatMapping = iAsst.hil_productsubcategorymapping;
                    }
                    iWrkOd.hil_CustomerComplaintDescription = bridge.NatureOfComplaint;
                    if (ServiceAccount != Guid.Empty)
                    {
                        iWrkOd.msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, ServiceAccount);
                        iWrkOd.msdyn_BillingAccount = new EntityReference(Account.EntityLogicalName, ServiceAccount);
                    }
                    iWrkOd.hil_CallSubType = new EntityReference(hil_callsubtype.EntityLogicalName, new Guid(bridge.IncidentType));
                    if (PriceList != Guid.Empty)
                    {
                        iWrkOd.msdyn_PriceList = new EntityReference(PriceLevel.EntityLogicalName, PriceList);
                    }
                    if (Nature != Guid.Empty)
                    {
                        iWrkOd.hil_natureofcomplaint = new EntityReference(hil_natureofcomplaint.EntityLogicalName, Nature);
                    }
                    if (IncidentType1 != Guid.Empty)
                    {
                        iWrkOd.msdyn_PrimaryIncidentType = new EntityReference(msdyn_incidenttype.EntityLogicalName, IncidentType1);
                    }
                    EntityReference CallType = new EntityReference(hil_calltype.EntityLogicalName);
                    hil_callsubtype Call = (hil_callsubtype)service.Retrieve(hil_callsubtype.EntityLogicalName, new Guid(bridge.IncidentType), new ColumnSet(true));
                    if (Call.hil_CallType != null)
                    {
                        CallType = Call.hil_CallType;
                        iWrkOd.hil_CallType = new EntityReference(hil_calltype.EntityLogicalName, CallType.Id);
                    }
                    iWrkOd.hil_quantity = 1;
                    iWrkOd.hil_SourceofJob = new OptionSetValue(4);
                    Guid enJobId = service.Create(iWrkOd);
                    if (enJobId != Guid.Empty)
                    {
                        msdyn_workorder enJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, enJobId, new ColumnSet("msdyn_name", "hil_fulladdress"));
                        if (enJob.msdyn_name != null)
                        {
                            bridge.WorkOrderID = enJob.msdyn_name;
                            bridge.CustomerAddress = enJob.hil_FullAddress != null ? enJob.hil_FullAddress : string.Empty;
                        }
                    }
                    if (enJobId != Guid.Empty)
                    {
                        new ProductRegistration().AttachNotes(service, bridge.ServiceAttachment, enJobId, bridge.FileType, "msdyn_workorder");
                        bridge.status = true;
                    }
                    bridge.status = true;
                    bridge.ErrCode = "XX";
                    bridge.ErrDesc = "XX";
                }
                else
                {
                    bridge.status = false;
                    bridge.ErrCode = "INCORRECT METHOD";
                    bridge.ErrDesc = "INCORRECT METHOD";
                }
                //OrganizationRequest req = new OrganizationRequest("hil_ConsumerApp_WorkOrder");
                //{
                //    req["Method"] = bridge.Method;
                //    req["ContactGuId"] = bridge.ContGuid;
                //    req["ProductSubCategory"] = bridge.ProductSubCategory;
                //    req["NatureOfComplaint"] = bridge.NatureOfComplaint;
                //    req["IncidentType"] = bridge.IncidentType;
                //    req["Asset"] = bridge.Asset;
                //};
                //OrganizationResponse response = service.Execute(req);
                ////bridge.ContGuid = (string)response["ContGuid"];
                //bridge.status = (bool)response["StatusCode"];
                //bridge.ErrCode = (string)response["ErrorCode"];
                //bridge.ErrDesc = (string)response["ErrorDescription"];
                //bridge.CustomerId = (string)response["CustomerGuId"];
                //bridge.WorkOrderID = (string)response["WorkOrderID"];
            }
            catch (Exception ex)
            {
                bridge.status = false;
                bridge.ErrCode = ex.Message.ToUpper();
                bridge.ErrDesc = ex.Message.ToUpper();
            }
            return (bridge);
        }
        public static Guid GetGuidbyName(String sEntityName, String sFieldName, String sFieldValue, IOrganizationService service, int iStatusCode = 0)
        {
            Guid fsResult = Guid.Empty;
            try
            {
                QueryExpression qe = new QueryExpression(sEntityName);
                qe.Criteria.AddCondition(sFieldName, ConditionOperator.Equal, sFieldValue);
                qe.AddOrder("createdon", OrderType.Descending);
                if (iStatusCode >= 0)
                {
                    qe.Criteria.AddCondition("statecode", ConditionOperator.Equal, iStatusCode);
                }
                EntityCollection enColl = service.RetrieveMultiple(qe);
                if (enColl.Entities.Count > 0)
                {
                    fsResult = enColl.Entities[0].Id;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Havells_Plugin.Helper.GetGuidbyName" + ex.Message);
            }
            return fsResult;
        }
    }

}