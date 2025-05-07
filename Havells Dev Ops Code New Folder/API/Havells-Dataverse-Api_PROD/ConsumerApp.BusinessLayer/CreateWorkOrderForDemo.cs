using Microsoft.Xrm.Sdk;
using System;
using Microsoft.Xrm.Sdk.Query;
using System.Runtime.Serialization;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class CreateWorkOrderForDemo
    {
        [DataMember]
        public string Method { get; set; }
        [DataMember(IsRequired = false)]
        public string ContGuid { get; set; }
        [DataMember(IsRequired = false)]
        public string EmailID { get; set; }
        [DataMember(IsRequired = false)]
        public string Mob { get; set; }
        [DataMember(IsRequired = false)]
        public string SPinCode { get; set; }//This Is String Formatted GuId
        [DataMember(IsRequired = false)]
        public string FName { get; set; }
        [DataMember(IsRequired = false)]
        public string LName { get; set; }
        [DataMember]
        public string ProductSubCategory { get; set; }
        [DataMember]
        public string IncidentType { get; set; }
        [DataMember(IsRequired = false)]
        public bool status { get; set; }
        [DataMember(IsRequired = false)]
        public string ErrCode { get; set; }
        [DataMember(IsRequired = false)]
        public string ErrDesc { get; set; }
        [DataMember(IsRequired = false)]
        public string WorkOrderID { get; set; }
        [DataMember(IsRequired = false)]
        public string CustomerId { get; set; }
        [DataMember(IsRequired = false)]
        public string StagingProductCategory { get; set; }
        [DataMember(IsRequired = false)]
        public string ADDRESS_ID { get; set; }
        [DataMember]
        public string ServiceAttachment { get; set; }
        [DataMember]
        public int FileType { get; set; }
        public CreateWorkOrderForDemo SubmitInstallationDemoRequest(CreateWorkOrderForDemo bridge)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();//get org service obj for connection
            #region NEW CODE
            try
            {
                if (bridge.Method == "CreateWODemo")
                {
                    if (bridge.ContGuid != null)
                    {
                        Guid JobId = CreateWorkOrderForDemoMethod(service, bridge, new Guid(bridge.ContGuid));
                        if (JobId != Guid.Empty)
                        {
                            msdyn_workorder enJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, JobId, new ColumnSet("msdyn_name"));

                            bridge.status = true;
                            bridge.ErrCode = "SUCCESS";
                            bridge.ErrDesc = "SUCCESS";
                            bridge.CustomerId = "";
                            bridge.WorkOrderID = enJob.msdyn_name;
                            new ProductRegistration().AttachNotes(service, bridge.ServiceAttachment, JobId, bridge.FileType, "msdyn_workorder");

                        }
                        //else
                        //{
                        //    bridge.status = false;
                        //    bridge.ErrCode = "UNEXPECTED ERROR PLEASE CHECK CONNECTION AND TRY AGAIN";
                        //    bridge.ErrDesc = "UNEXPECTED ERROR PLEASE CHECK CONNECTION AND TRY AGAIN";
                        //    bridge.CustomerId = "UNEXPECTED ERROR PLEASE CHECK CONNECTION AND TRY AGAIN";
                        //    bridge.WorkOrderID = "UNEXPECTED ERROR PLEASE CHECK CONNECTION AND TRY AGAIN";
                        //}
                    }
                    else if (bridge.EmailID != null && bridge.Mob != null && bridge.FName != null && bridge.LName != null)// && bridge.SPinCode != null)
                    {
                        Guid guContId = GetThisCustomerIfExists(service, bridge);
                        if (guContId != Guid.Empty)
                        {
                            Guid JobId = CreateWorkOrderForDemoMethod(service, bridge, guContId);
                            if (JobId != Guid.Empty)
                            {
                                msdyn_workorder enJob = (msdyn_workorder)service.Retrieve(msdyn_workorder.EntityLogicalName, JobId, new ColumnSet("msdyn_name"));
                                bridge.status = true;
                                bridge.ErrCode = "SUCCESS";
                                bridge.ErrDesc = "SUCCESS";
                                bridge.CustomerId = guContId.ToString();
                                bridge.WorkOrderID = enJob.msdyn_name;//  JobId.ToString();
                                new ProductRegistration().AttachNotes(service, bridge.ServiceAttachment, JobId, bridge.FileType, "msdyn_workorder");

                            }
                            //else
                            //{
                            //    bridge.status = false;
                            //    bridge.ErrCode = "UNEXPECTED ERROR PLEASE CHECK CONNECTION AND TRY AGAIN";
                            //    bridge.ErrDesc = "UNEXPECTED ERROR PLEASE CHECK CONNECTION AND TRY AGAIN";
                            //    bridge.CustomerId = "UNEXPECTED ERROR PLEASE CHECK CONNECTION AND TRY AGAIN";
                            //    bridge.WorkOrderID = "UNEXPECTED ERROR PLEASE CHECK CONNECTION AND TRY AGAIN";
                            //}
                        }
                    }
                    else
                    {
                        bridge.status = false;
                        bridge.ErrCode = "MANDATORY DETAILS MISSING IN PAYLOAD";
                        bridge.ErrDesc = "MANDATORY DETAILS MISSING IN PAYLOAD";
                        bridge.CustomerId = "MANDATORY DETAILS MISSING IN PAYLOAD";
                        bridge.WorkOrderID = "MANDATORY DETAILS MISSING IN PAYLOAD";
                    }
                }
                else
                {
                    bridge.status = false;
                    bridge.ErrCode = "METHOD NAME DOESN'T MATCH";
                    bridge.ErrDesc = "METHOD NAME DOESN'T MATCH";
                    bridge.CustomerId = "METHOD NAME DOESN'T MATCH";
                    bridge.WorkOrderID = "METHOD NAME DOESN'T MATCH";
                }
            }
            catch (Exception ex)
            {
                bridge.status = false;
                bridge.ErrCode = ex.Message.ToUpper();
                bridge.ErrDesc = ex.Message.ToUpper();
                bridge.CustomerId = ex.Message.ToUpper();
                bridge.WorkOrderID = ex.Message.ToUpper();
            }
            #endregion
            #region UN-USED
            //OrganizationRequest req = new OrganizationRequest("hil_ConsumerApp_WorkOrderDemo95cd59a51789e811a95e000d3af05df5");
            //{
            //    req["Method"] = bridge.Method;
            //    req["ContactGuId"] = bridge.ContGuid;
            //    req["EmailId"] = bridge.EmailID;
            //    req["MobileNumber"] = bridge.Mob;
            //    req["ShipPinCode"] = bridge.SPinCode;
            //    req["FirstName"] = bridge.FName;
            //    req["LastName"] = bridge.LName;
            //    req["ProductSubCategory"] = bridge.ProductSubCategory;
            //    req["IncidentType"] = bridge.IncidentType;
            //};
            //OrganizationResponse response = service.Execute(req);
            //bridge.status = (bool)response["StatusCode"];
            //bridge.ErrCode = (string)response["ErrorCode"];
            //bridge.ErrDesc = (string)response["ErrorDescription"];
            //bridge.CustomerId = (string)response["CustomerGuId"];
            //bridge.WorkOrderID = (string)response["WorkOrderID"];
            #endregion
            return (bridge);
        }
        #region FIND CUSTOMER 
        public static Guid GetThisCustomerIfExists(IOrganizationService service, CreateWorkOrderForDemo bridge)
        {
            Guid ContactId = Guid.Empty;
            QueryExpression Query = new QueryExpression(Contact.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(false);
            Query.Criteria = new FilterExpression(LogicalOperator.Or);
            Query.Criteria.AddCondition(new ConditionExpression("mobilephone", ConditionOperator.Equal, bridge.Mob));
            Query.Criteria.AddCondition(new ConditionExpression("emailaddress1", ConditionOperator.Equal, bridge.EmailID));
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                foreach (Contact cnt in Found.Entities)
                {
                    ContactId = cnt.Id;
                }
            }
            else
            {
                Contact enContact = new Contact();
                enContact.FirstName = bridge.FName;
                enContact.LastName = bridge.LName;
                enContact.MobilePhone = bridge.Mob;
                enContact.EMailAddress1 = bridge.EmailID;
                Guid guidContactId = service.Create(enContact);
                if (guidContactId != Guid.Empty)
                {
                    hil_address enAddress = new hil_address();
                    enAddress.hil_AddressType = new OptionSetValue(1);
                    enAddress.hil_Customer = new EntityReference(Contact.EntityLogicalName, guidContactId);
                    enAddress.hil_PinCode = new EntityReference(hil_pincode.EntityLogicalName, new Guid(bridge.SPinCode));
                    hil_businessmapping iDetails = AddAddress.GetBusinessGeo(service, new Guid(bridge.SPinCode));
                    if (iDetails.Id != Guid.Empty)
                    {
                        enAddress.hil_SalesOffice = iDetails.hil_salesoffice;
                        enAddress.hil_Area = iDetails.hil_area;
                        enAddress.hil_Branch = iDetails.hil_branch;
                        enAddress.hil_Region = iDetails.hil_region;
                    }
                    service.Create(enAddress);
                }
            }
            return ContactId;
        }
        #endregion
        #region CREATE WORK ORDER
        public static Guid CreateWorkOrderForDemoMethod(IOrganizationService service, CreateWorkOrderForDemo bridge, Guid gCustomer)
        {
            msdyn_workorder enJob = new msdyn_workorder();
            Contact Cont = (Contact)service.Retrieve(Contact.EntityLogicalName, gCustomer, new ColumnSet("mobilephone", "emailaddress1"));
            Guid PriceList = ConsumerApp.BusinessLayer.CreateWorkOrder.GetGuidbyName(PriceLevel.EntityLogicalName, "name", "Default Price List", service);
            Guid ServiceAccount = ConsumerApp.BusinessLayer.CreateWorkOrder.GetGuidbyName(Account.EntityLogicalName, "name", "Dummy Customer", service);
            hil_callsubtype iCall = (hil_callsubtype)service.Retrieve(hil_callsubtype.EntityLogicalName, new Guid(bridge.IncidentType), new ColumnSet(true));
            enJob.hil_CustomerRef = new EntityReference(Contact.EntityLogicalName, gCustomer);
            enJob.hil_Email = Cont.EMailAddress1;
            enJob.hil_mobilenumber = Cont.MobilePhone;
            enJob.hil_CallSubType = new EntityReference(hil_callsubtype.EntityLogicalName, iCall.Id);
            if (PriceList != Guid.Empty)
            {
                enJob.msdyn_PriceList = new EntityReference(PriceLevel.EntityLogicalName, PriceList);
            }
            if (ServiceAccount != Guid.Empty)
            {
                enJob.msdyn_ServiceAccount = new EntityReference(Account.EntityLogicalName, ServiceAccount);
                enJob.msdyn_BillingAccount = new EntityReference(Account.EntityLogicalName, ServiceAccount);
            }
            if (iCall.hil_CallType != null)
            {
                enJob.hil_CallType = new EntityReference(hil_calltype.EntityLogicalName, iCall.hil_CallType.Id);
            }
            hil_stagingdivisonmaterialgroupmapping Stage = (hil_stagingdivisonmaterialgroupmapping)service.Retrieve(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, new Guid(bridge.ProductSubCategory), new ColumnSet("hil_productcategorydivision", "hil_productsubcategorymg"));
            if (Stage.hil_ProductCategoryDivision != null && Stage.hil_ProductSubCategoryMG != null)
            {
                enJob.hil_Productcategory = Stage.hil_ProductCategoryDivision;
                enJob.hil_ProductCatSubCatMapping = new EntityReference(hil_stagingdivisonmaterialgroupmapping.EntityLogicalName, Stage.Id);
                enJob.hil_ProductSubcategory = Stage.hil_ProductSubCategoryMG;
                Guid iNature = GetNature(service, iCall.hil_name, Stage.hil_ProductSubCategoryMG);
                if (iNature != Guid.Empty)
                    enJob.hil_natureofcomplaint = new EntityReference(hil_natureofcomplaint.EntityLogicalName, iNature);
            }
            enJob.hil_quantity = 1;
            enJob.hil_SourceofJob = new OptionSetValue(4);
            Guid JobId = service.Create(enJob);
            //enJob.hil_CustomerComplaintDescription = Desc;
            return JobId;
        }
        #endregion
        #region GET NATURE
        public static Guid GetNature(IOrganizationService service, string Call, EntityReference ProductSCat)
        {
            Guid Nature = new Guid();
            QueryExpression Query = new QueryExpression(hil_natureofcomplaint.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(true);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, Call));
            Query.Criteria.AddCondition(new ConditionExpression("hil_relatedproduct", ConditionOperator.Equal, ProductSCat.Id));
            Query.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                Nature = Found.Entities[0].Id;
            }
            return Nature;
        }
        #endregion
    }
}