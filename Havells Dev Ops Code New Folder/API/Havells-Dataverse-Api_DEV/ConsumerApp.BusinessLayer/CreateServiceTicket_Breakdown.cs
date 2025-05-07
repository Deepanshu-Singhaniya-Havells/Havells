using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class CreateServiceTicket_Breakdown
    {
        [DataMember]
        public string ASSET { get; set; }
        [DataMember]
        public string METHOD { get; set; }
        [DataMember]
        public string CONTACT_GUID { get; set; }
        [DataMember]
        public string PRODUCT_SUBCATEGORY { get; set; }
        [DataMember]
        public string INCIDENT_TYPE { get; set; }
        [DataMember(IsRequired = false)]
        public string NATURE_OF_COMPLAINT { get; set; }
        [DataMember]
        public string SERV_ADDRESS { get; set; }
        [DataMember]
        public string ServiceAttachment { get; set; }
        [DataMember]
        public int FileType { get; set; }
        public OutputCreateServiceTicket_Breakdown SubmitJob_New(CreateServiceTicket_Breakdown bridge)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();//get org service obj for connection
            OutputCreateServiceTicket_Breakdown iOut = new OutputCreateServiceTicket_Breakdown();
            try
            {
                if (bridge.METHOD == "CreateWOService")
                {
                    msdyn_workorder iWrkOd = new msdyn_workorder();
                    Contact Cont = (Contact)service.Retrieve(Contact.EntityLogicalName, new Guid(bridge.CONTACT_GUID), new ColumnSet("mobilephone", "emailaddress1"));
                    Guid PriceList = ConsumerApp.BusinessLayer.CreateWorkOrder.GetGuidbyName(PriceLevel.EntityLogicalName, "name", "Default Price List", service);
                    Guid Nature = ConsumerApp.BusinessLayer.CreateWorkOrder.GetGuidbyName(hil_natureofcomplaint.EntityLogicalName, "hil_name", "Service", service);
                    Guid IncidentType1 = ConsumerApp.BusinessLayer.CreateWorkOrder.GetGuidbyName(msdyn_incidenttype.EntityLogicalName, "msdyn_name", "Service", service);
                    Guid ServiceAccount = ConsumerApp.BusinessLayer.CreateWorkOrder.GetGuidbyName(Account.EntityLogicalName, "name", "Dummy Customer", service);
                    msdyn_customerasset iAsst = (msdyn_customerasset)service.Retrieve(msdyn_customerasset.EntityLogicalName, new Guid(bridge.ASSET), new ColumnSet("msdyn_product", "hil_productsubcategory", "hil_productsubcategorymapping", "hil_productcategory"));
                    hil_callsubtype Call = (hil_callsubtype)service.Retrieve(hil_callsubtype.EntityLogicalName, new Guid(bridge.INCIDENT_TYPE), new ColumnSet(true));
                    hil_address ThisAddress = (hil_address)service.Retrieve(hil_address.EntityLogicalName, new Guid(bridge.SERV_ADDRESS), new ColumnSet(true));
                    iWrkOd.hil_CustomerRef = new EntityReference(Contact.EntityLogicalName, new Guid(bridge.CONTACT_GUID));
                    iWrkOd.hil_Email = Cont.EMailAddress1;
                    iWrkOd.hil_mobilenumber = Cont.MobilePhone;
                    iWrkOd.hil_Address = new EntityReference(hil_address.EntityLogicalName, new Guid(bridge.SERV_ADDRESS));
                    if (ThisAddress.hil_PinCode != null)
                    {
                        iWrkOd.hil_pincode = ThisAddress.hil_PinCode;
                        iWrkOd.hil_PinCodeText = ThisAddress.hil_PinCode.Name;
                    }
                    if (ThisAddress.hil_CIty != null)
                    {
                        iWrkOd.hil_City = ThisAddress.hil_CIty;
                        iWrkOd.hil_CityText = ThisAddress.hil_CIty.Name;
                    }
                    if (ThisAddress.hil_State != null)
                    {
                        iWrkOd.hil_state = ThisAddress.hil_State;
                        iWrkOd.hil_StateText = ThisAddress.hil_State.Name;
                    }
                    if (ThisAddress.hil_Region != null)
                    {
                        iWrkOd.hil_Region = ThisAddress.hil_Region;
                        iWrkOd.hil_RegionText = ThisAddress.hil_Region.Name;
                    }
                    if (ThisAddress.hil_SalesOffice != null)
                    {
                        iWrkOd["hil_salesoffice"] = new EntityReference(hil_salesoffice.EntityLogicalName, ThisAddress.hil_SalesOffice.Id);
                    }
                    if (ThisAddress.hil_Branch != null)
                    {
                        iWrkOd.hil_Branch = ThisAddress.hil_Branch;
                        iWrkOd.hil_BranchText = ThisAddress.hil_Branch.Name;
                    }
                    iWrkOd.msdyn_CustomerAsset = new EntityReference(msdyn_customerasset.EntityLogicalName, new Guid(bridge.ASSET));
                    if (iAsst.hil_ProductCategory != null)
                    {
                        iWrkOd.hil_Productcategory = iAsst.hil_ProductCategory;
                    }
                    if (iAsst.hil_ProductSubcategory != null)
                    {
                        iWrkOd.hil_ProductSubcategory = iAsst.hil_ProductSubcategory;
                        Product enSubCategory = (Product)service.Retrieve(Product.EntityLogicalName, iAsst.hil_ProductSubcategory.Id, new ColumnSet("hil_minimumthreshold", "hil_maximumthreshold"));
                        if (enSubCategory.hil_MinimumThreshold != null && enSubCategory.hil_MaximumThreshold != null)
                        {
                            iWrkOd.hil_MinQuantity = enSubCategory.hil_MinimumThreshold;
                            iWrkOd.hil_MaxQuantity = enSubCategory.hil_MaximumThreshold;
                        }
                    }
                    if (iAsst.hil_productsubcategorymapping != null)
                    {
                        iWrkOd.hil_ProductCatSubCatMapping = iAsst.hil_productsubcategorymapping;
                    }
                    iWrkOd.hil_CustomerComplaintDescription = bridge.NATURE_OF_COMPLAINT;
                    if (ServiceAccount != Guid.Empty)
                    {
                        iWrkOd.msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, ServiceAccount);
                        iWrkOd.msdyn_BillingAccount = new EntityReference(Account.EntityLogicalName, ServiceAccount);
                    }
                    iWrkOd.hil_CallSubType = new EntityReference(hil_callsubtype.EntityLogicalName, new Guid(bridge.INCIDENT_TYPE));
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
                    if (Call.hil_CallType != null)
                    {
                        CallType = Call.hil_CallType;
                        iWrkOd.hil_CallType = new EntityReference(hil_calltype.EntityLogicalName, CallType.Id);
                    }
                    iWrkOd.hil_quantity = 1;
                    iWrkOd.hil_SourceofJob = new OptionSetValue(4);
                    Guid JobId = service.Create(iWrkOd);
                    if (JobId != Guid.Empty)
                    {
                        msdyn_workorder enJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, JobId, new ColumnSet("msdyn_name"));

                        iOut.WorkOrderID = enJob.msdyn_name;
                        new ProductRegistration().AttachNotes(service, bridge.ServiceAttachment, JobId, bridge.FileType, "msdyn_workorder");

                    }
                    iOut.STATUS = true;
                    iOut.ERROR_CODE = "XX";
                    iOut.ERROR_DESC = "XX";
                }
            }
            catch (Exception ex)
            {
                iOut.STATUS = false;
                iOut.ERROR_CODE = "1";
                iOut.ERROR_DESC = ex.Message.ToUpper();
            }
            return iOut;
        }
    }
    public class OutputCreateServiceTicket_Breakdown
    {
        [DataMember(IsRequired = false)]
        public bool STATUS { get; set; }
        [DataMember(IsRequired = false)]
        public string ERROR_CODE { get; set; }
        [DataMember(IsRequired = false)]
        public string ERROR_DESC { get; set; }
        [DataMember(IsRequired = false)]
        public string WorkOrderID { get; set; }
    }
}