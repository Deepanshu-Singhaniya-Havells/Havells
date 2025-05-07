using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class ProductRegistration
    {
        [DataMember]
        public string Method { get; set; }
        [DataMember]
        public string CustomerGUID { get; set; }
        [DataMember(IsRequired = false)]
        public string ProductSerialNumber { get; set; }
        [DataMember]
        public string StagingProductSubCategory { get; set; }
        //[DataMember]
        //public string Model { get; set; }
        [DataMember]
        public string InvoiceAttachment { get; set; }
        [DataMember]
        public string InvoiceDate { get; set; }
        [DataMember]
        public string InvoiceNumber { get; set; }
        [DataMember]
        public string PinCode { get; set; }
        [DataMember]
        public int FileType { get; set; }
        [DataMember(IsRequired = false)]
        public int PROD_LOC { get; set; }
        [DataMember]
        public string Model { get; set; }
        [DataMember(IsRequired = false)]
        public string DealerPin { get; set; }
        [DataMember(IsRequired = false)]
        public bool status { get; set; }//Output
        [DataMember(IsRequired = false)]
        public string ErrCode { get; set; }
        [DataMember(IsRequired = false)]
        public string ErrDesc { get; set; }
        [DataMember(IsRequired = false)]
        public string PURCHASE_FROM { get; set; }
        [DataMember(IsRequired = false)]
        public string RetailerPINCode { get; set; }
        [DataMember(IsRequired = false)]
        public decimal InvoiceValue { get; set; }
        public ProductRegistration InitiateProductRegistration(ProductRegistration bridge)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();//get org service obj for connection
            try
            {
                if (bridge.Method == "ProductRegistration")
                {
                    hil_stagingdivisonmaterialgroupmapping Stage = (hil_stagingdivisonmaterialgroupmapping)service.Retrieve(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, new Guid(bridge.StagingProductSubCategory), new ColumnSet("hil_productcategorydivision", "hil_productsubcategorymg"));
                    Guid DummyAc = GetGuidbyName98(Account.EntityLogicalName, "name", "Dummy Customer", service);
                    msdyn_customerasset _CstA = new msdyn_customerasset();
                    if (bridge.ProductSerialNumber != null)
                    {
                        _CstA.msdyn_name = bridge.ProductSerialNumber;
                    }
                    if (bridge.CustomerGUID != null)
                    {
                        _CstA.hil_Customer = new EntityReference(Contact.EntityLogicalName, new Guid(bridge.CustomerGUID));
                    }
                    if (bridge.Model != null)
                    {
                        _CstA.msdyn_Product = new EntityReference(Product.EntityLogicalName, new Guid(bridge.Model));
                    }
                    if (bridge.DealerPin != null)
                    {
                        QueryExpression Qry = new QueryExpression(hil_pincode.EntityLogicalName);
                        //Qry.EntityName = hil_pincode.EntityLogicalName;
                        ColumnSet Col = new ColumnSet("hil_name", "hil_pincodeid", "hil_city", "hil_state");
                        Qry.ColumnSet = Col;
                        Qry.Criteria = new FilterExpression(LogicalOperator.And);
                        Qry.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                        Qry.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, bridge.DealerPin));
                        EntityCollection Colec = service.RetrieveMultiple(Qry);
                        if (Colec.Entities.Count >= 1)
                        {
                            hil_pincode iPin = Colec.Entities[0].ToEntity<hil_pincode>();
                            EntityReference State = new EntityReference();
                            EntityReference City = new EntityReference();
                            QueryExpression Qry1 = new QueryExpression(hil_businessmapping.EntityLogicalName);
                            ColumnSet Col1 = new ColumnSet(false);
                            Qry1.ColumnSet = Col1;
                            Qry1.Criteria = new FilterExpression(LogicalOperator.And);
                            Qry1.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                            Qry1.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, iPin.Id));
                            EntityCollection Colec1 = service.RetrieveMultiple(Qry1);
                            if (Colec1.Entities.Count > 0)
                            {
                                hil_businessmapping iBus = Colec1.Entities[0].ToEntity<hil_businessmapping>();
                                _CstA["hil_pincode"] = new EntityReference(hil_businessmapping.EntityLogicalName, iBus.Id);
                            }
                        }
                    }
                    if (bridge.InvoiceDate != null)
                    {
                        DateTime dtInvoice = DateTime.ParseExact(bridge.InvoiceDate, "yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
                        _CstA.hil_InvoiceDate = dtInvoice;
                    }
                    if (bridge.InvoiceNumber != null)
                    {
                        _CstA.hil_InvoiceNo = bridge.InvoiceNumber;
                    }
                    if (DummyAc != Guid.Empty)
                    {
                        _CstA.msdyn_Account = new EntityReference(Account.EntityLogicalName, DummyAc);
                    }
                    if (Stage.hil_ProductCategoryDivision != null && Stage.hil_ProductSubCategoryMG != null)
                    {
                        _CstA.hil_ProductCategory = Stage.hil_ProductCategoryDivision;
                        _CstA.hil_ProductSubcategory = Stage.hil_ProductSubCategoryMG;
                        _CstA.hil_productsubcategorymapping = new EntityReference(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, Stage.Id);
                    }
                    if (bridge.PROD_LOC != 0)
                    {
                        _CstA.hil_Product = new OptionSetValue(bridge.PROD_LOC);
                    }
                    if (bridge.PURCHASE_FROM != null)
                    {
                        _CstA["hil_purchasedfrom"] = bridge.PURCHASE_FROM;
                    }
                    if (bridge.RetailerPINCode != null)
                    {
                        _CstA["hil_retailerpincode"] = bridge.RetailerPINCode;
                    }
                    if (bridge.InvoiceValue > 0)
                    {
                        _CstA.hil_InvoiceValue = bridge.InvoiceValue;
                    }
                    _CstA.hil_CreateWarranty = false;
                    _CstA["hil_source"] = new OptionSetValue(1); // Consumer Mobile App
                    _CstA["statuscode"] = new OptionSetValue(910590000); // Pending for Approval

                    Guid CustomerAssetId = service.Create(_CstA);
                    if (CustomerAssetId != Guid.Empty)
                    {
                        AttachNotes(service, bridge.InvoiceAttachment, CustomerAssetId, bridge.FileType, "msdyn_customerasset");
                        bridge.status = true;
                    }
                }
                else
                {
                    bridge.ErrDesc = "01";
                    bridge.ErrCode = "WRONG OPERATION CODE";
                    bridge.status = false;
                }
                //OrganizationRequest req = new OrganizationRequest("hil_ConsumerApp_ProductRegistrationaddba877ca7ae811a95e000d3af06236");
                //{
                //    req["Method"] = bridge.Method;
                //    req["ContactGuId"] = bridge.CustomerGUID;
                //    req["ProductSerialNo"] = bridge.ProductSerialNumber;
                //    req["InvoiceDate"] = bridge.InvoiceDate;
                //    req["InvoiceNumber"] = bridge.InvoiceNumber;
                //    req["Base64String"] = bridge.InvoiceAttachment;
                //    req["PinCode"] = bridge.PinCode;
                //    req["FileType"] = bridge.FileType;
                //    req["ProductSubCategory"] = bridge.StagingProductSubCategory;
                //    req["Model"] = bridge.Model;
                //    req["DealerPin"] = bridge.DealerPin;
                //    req["ProductLocation"] = bridge.PROD_LOC;
                //    req["Model"] = bridge.Model;
                //    req["PurchasedFrom"] = bridge.PURCHASE_FROM;
                //};
                //OrganizationResponse response = service.Execute(req);
                ////bridge.ContGuid = (string)response["ContGuid"];
                //bridge.status = (bool)response["StatusCode"];
                //bridge.ErrCode = (string)response["ErrorCode"];
                //bridge.ErrDesc = (string)response["ErrorDescription"];
            }
            catch (Exception ex)
            {
                bridge.ErrDesc = ex.Message.ToUpper();
                bridge.ErrCode = ex.Message.ToUpper();
                bridge.status = false;
            }
            return (bridge);
        }

        public void AttachNotes(IOrganizationService service, string Notes, Guid Asset, int fileType, string entityType)
        {
            try
            {
                if ((Notes != null) && (fileType != null))//int can't be null
                {
                    if (fileType == 0)
                    {
                        byte[] NoteByte = Convert.FromBase64String(Notes);
                        Annotation An = new Annotation();
                        An.DocumentBody = Convert.ToBase64String(NoteByte);
                        An.MimeType = "image/jpeg";
                        An.FileName = "invoice.jpeg";
                        An.ObjectId = new EntityReference(entityType, Asset);
                        An.ObjectTypeCode = entityType;
                        service.Create(An);
                    }
                    else if (fileType == 1)
                    {
                        byte[] NoteByte = Convert.FromBase64String(Notes);
                        Annotation An = new Annotation();
                        An.DocumentBody = Convert.ToBase64String(NoteByte);
                        An.MimeType = "image/png";
                        An.FileName = "invoice.png";
                        An.ObjectId = new EntityReference(entityType, Asset);
                        An.ObjectTypeCode = entityType;
                        service.Create(An);
                    }
                    else if (fileType == 2)
                    {
                        byte[] NoteByte = Convert.FromBase64String(Notes);
                        Annotation An = new Annotation();
                        An.DocumentBody = Convert.ToBase64String(NoteByte);
                        An.MimeType = "application/pdf";
                        An.FileName = "invoice.pdf";
                        An.ObjectId = new EntityReference(entityType, Asset);
                        An.ObjectTypeCode = entityType;
                        service.Create(An);
                    }
                    else if (fileType == 3)
                    {
                        byte[] NoteByte = Convert.FromBase64String(Notes);
                        Annotation An = new Annotation();
                        An.DocumentBody = Convert.ToBase64String(NoteByte);
                        An.MimeType = "application/doc";
                        An.FileName = "invoice.doc";
                        An.ObjectId = new EntityReference(entityType, Asset);
                        An.ObjectTypeCode = entityType;
                        service.Create(An);
                    }
                    else if (fileType == 4)
                    {
                        byte[] NoteByte = Convert.FromBase64String(Notes);
                        Annotation An = new Annotation();
                        An.DocumentBody = Convert.ToBase64String(NoteByte);
                        An.MimeType = "application/tiff";
                        An.FileName = "invoice.tiff";
                        An.ObjectId = new EntityReference(entityType, Asset);
                        An.ObjectTypeCode = entityType;
                        service.Create(An);
                    }
                    else if (fileType == 5)
                    {
                        byte[] NoteByte = Convert.FromBase64String(Notes);
                        Annotation An = new Annotation();
                        An.DocumentBody = Convert.ToBase64String(NoteByte);
                        An.MimeType = "application/gif";
                        An.FileName = "invoice.gif";
                        An.ObjectId = new EntityReference(entityType, Asset);
                        An.ObjectTypeCode = entityType;
                        service.Create(An);
                    }
                    else if (fileType == 6)
                    {
                        byte[] NoteByte = Convert.FromBase64String(Notes);
                        Annotation An = new Annotation();
                        An.DocumentBody = Convert.ToBase64String(NoteByte);
                        An.MimeType = "application/bmp";
                        An.FileName = "invoice.bmp";
                        An.ObjectId = new EntityReference(entityType, Asset);
                        An.ObjectTypeCode = entityType;
                        service.Create(An);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error Attaching Notes" + ex.Message);
            }
        }
        public Guid GetGuidbyName98(String sEntityName, String sFieldName, String sFieldValue, IOrganizationService service, int iStatusCode = 0)
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
