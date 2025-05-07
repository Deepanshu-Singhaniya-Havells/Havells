using System;
using System.Collections.Generic;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Runtime.Serialization;

namespace TestApp
{
    public class IotServiceCall
    {
        public Guid CustomerGuid { get; set; }

        public IoTServiceCallRegistration IoTCreateServiceCallDealerPortal(IoTServiceCallRegistration serviceCalldata, IOrganizationService service)
        {
            IoTServiceCallRegistration objServiceCall;
            Guid customerGuid = Guid.Empty;
            Guid callSubTypeGuid = Guid.Empty;
            Guid serviceCallGuid = Guid.Empty;
            Entity lookupObj = null;
            EntityCollection entcoll;
            QueryExpression Query;
            string customerFullName = string.Empty;
            string customerMobileNumber = string.Empty;
            string customerEmail = string.Empty;
            Guid customerAssetGuid = Guid.Empty;
            DateTime? invoiceDate = null;
            string modelName = string.Empty;
            EntityReference erProductsubcategorymapping = null;
            EntityReference erProductCategory = null;
            EntityReference erProductsubcategory = null;

            EntityReference erNatureOfComplaint = null;
            EntityReference erCustomerAsset = null;
            bool continueFlag = false;
            string fullAddress = string.Empty;
            try
            {
                //IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (serviceCalldata.CustomerMobleNo == string.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer Nobile Number is required." };
                    }
                    if (serviceCalldata.CustomerGuid == Guid.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer Guid is required." };
                    }
                    if (serviceCalldata.NOCGuid == Guid.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Nature of Complaint is required." };
                    }
                    if (serviceCalldata.AddressGuid == Guid.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Consumer Address is required." };
                    }
                    if (serviceCalldata.ProductCategoryGuid == Guid.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Product Category is required." };
                    }
                    if (serviceCalldata.ProductSubCategoryGuid == Guid.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Product Sub Category is required." };
                    }

                    Query = new QueryExpression("hil_address");
                    Query.ColumnSet = new ColumnSet("hil_addressid");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_addressid", ConditionOperator.Equal, serviceCalldata.AddressGuid);
                    Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, serviceCalldata.CustomerGuid);
                    entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count == 0)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Address does not belong to Customer." };
                    }

                    Query = new QueryExpression("contact");
                    Query.ColumnSet = new ColumnSet("fullname", "emailaddress1", "mobilephone");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, serviceCalldata.CustomerMobleNo);
                    Query.Criteria.AddCondition("contactid", ConditionOperator.Equal, serviceCalldata.CustomerGuid);
                    entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count == 0)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer/Mobile No. does not exist." };
                    }
                    else
                    {
                        lookupObj = entcoll.Entities[0];
                        customerFullName = lookupObj.GetAttributeValue<string>("fullname");
                        customerEmail = lookupObj.GetAttributeValue<string>("emailaddress1");
                        customerMobileNumber = lookupObj.GetAttributeValue<string>("mobilephone"); // N
                    }
                    if (serviceCalldata.ChiefComplaint == string.Empty || serviceCalldata.ChiefComplaint == null || serviceCalldata.ChiefComplaint.Trim().Length == 0)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer's Chief Complaint is required." };
                    }
                    //Case 1 Serial Number Exists
                    if (serviceCalldata.AssetGuid != Guid.Empty)
                    {
                        Entity ent = service.Retrieve("msdyn_customerasset", serviceCalldata.AssetGuid,new ColumnSet("msdyn_name", "hil_customer", "hil_productsubcategorymapping", "hil_productcategory", "hil_productsubcategory", "msdyn_customerassetid", "hil_invoicedate"));
                        if (ent != null)
                        {
                            erCustomerAsset = ent.ToEntityReference();
                            modelName = ent.GetAttributeValue<string>("msdyn_name");
                            invoiceDate = ent.GetAttributeValue<DateTime>("hil_invoicedate");
                            erProductCategory = ent.GetAttributeValue<EntityReference>("hil_productcategory");
                            erProductsubcategory = ent.GetAttributeValue<EntityReference>("hil_productsubcategory");
                            erProductsubcategorymapping = ent.GetAttributeValue<EntityReference>("hil_productsubcategorymapping");
                            continueFlag = true;
                        }
                        else
                        {
                            return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer Asset Serial Number does not exist." };
                        }
                    }
                    //Case 2 Product Category 
                    else if (serviceCalldata.ProductCategoryGuid != Guid.Empty)
                    {
                        erProductCategory = new EntityReference("product", serviceCalldata.ProductCategoryGuid);
                        erProductsubcategory = new EntityReference("product", serviceCalldata.ProductSubCategoryGuid);
                        modelName = string.Empty;
                        continueFlag = true;
                    }

                    if (!continueFlag)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Asset Serial Number/Product Category is required to proceed." };
                    }

                    if (serviceCalldata.SourceOfJob != 6)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Invalid Source of Job!!!" };
                    }

                    #region Get Nature of Complaint
                    string fetchXML = string.Empty;

                    if (serviceCalldata.NOCGuid != Guid.Empty)
                    {
                        fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>";
                        fetchXML += "<entity name='hil_natureofcomplaint'>";
                        fetchXML += "<attribute name='hil_callsubtype' />";
                        fetchXML += "<attribute name='hil_natureofcomplaintid' />";
                        fetchXML += "<order attribute='createdon' descending='false' />";
                        fetchXML += "<filter type='and'>";
                        fetchXML += "<condition attribute='hil_relatedproduct' operator='eq' value='{" + erProductsubcategory.Id + "}' />";
                        fetchXML += "<condition attribute='hil_natureofcomplaintid' operator='eq' value='{" + serviceCalldata.NOCGuid + "}' />";
                        fetchXML += "</filter>";
                        fetchXML += "</entity>";
                        fetchXML += "</fetch>";
                        entcoll = service.RetrieveMultiple(new FetchExpression(fetchXML));
                        if (entcoll.Entities.Count > 0)
                        {
                            erNatureOfComplaint = entcoll.Entities[0].ToEntityReference();
                            if (entcoll.Entities[0].Attributes.Contains("hil_callsubtype"))
                            {
                                callSubTypeGuid = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_callsubtype").Id;
                            }
                        }
                        else
                        {
                            return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "NOC does not match with Product Sub Category." };
                        }
                    }
                    #endregion
                    #region Create Service Call
                    objServiceCall = new IoTServiceCallRegistration();
                    objServiceCall = serviceCalldata;
                    Entity enWorkorder = new Entity("msdyn_workorder");

                    if (serviceCalldata.CustomerGuid != Guid.Empty)
                    {
                        enWorkorder["hil_customerref"] = new EntityReference("contact", serviceCalldata.CustomerGuid);
                    }
                    enWorkorder["hil_customername"] = customerFullName;
                    enWorkorder["hil_mobilenumber"] = customerMobileNumber;
                    enWorkorder["hil_email"] = customerEmail;

                    if (serviceCalldata.PreferredPartOfDay > 0 && serviceCalldata.PreferredPartOfDay < 4)
                    {
                        enWorkorder["hil_preferredtime"] = new OptionSetValue(serviceCalldata.PreferredPartOfDay);
                    }

                    if (serviceCalldata.PreferredDate != null && serviceCalldata.PreferredDate.Trim().Length > 0)
                    {
                        string _date = serviceCalldata.PreferredDate;
                        DateTime dtInvoice = new DateTime(Convert.ToInt32(_date.Substring(6, 4)), Convert.ToInt32(_date.Substring(0, 2)), Convert.ToInt32(_date.Substring(3, 2)));
                        enWorkorder["hil_preferreddate"] = dtInvoice;
                    }

                    if (serviceCalldata.AddressGuid != Guid.Empty)
                    {
                        enWorkorder["hil_address"] = new EntityReference("hil_address", serviceCalldata.AddressGuid);
                    }

                    if (erCustomerAsset != null)
                    {
                        enWorkorder["msdyn_customerasset"] = erCustomerAsset;
                    }

                    if (modelName != string.Empty)
                    {
                        enWorkorder["hil_modelname"] = modelName;
                    }

                    if (erProductCategory != null)
                    {
                        enWorkorder["hil_productcategory"] = erProductCategory;
                    }
                    if (erProductsubcategory != null)
                    {
                        enWorkorder["hil_productsubcategory"] = erProductsubcategory;
                    }

                    Query = new QueryExpression("hil_stagingdivisonmaterialgroupmapping");
                    Query.ColumnSet = new ColumnSet("hil_stagingdivisonmaterialgroupmappingid");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_productcategorydivision", ConditionOperator.Equal, erProductCategory.Id);
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_productsubcategorymg", ConditionOperator.Equal, erProductsubcategory.Id);
                    EntityCollection ec = service.RetrieveMultiple(Query);
                    if (ec.Entities.Count > 0)
                    {
                        enWorkorder["hil_productcatsubcatmapping"] = ec.Entities[0].ToEntityReference();
                    }

                    EntityCollection entCol;
                    Query = new QueryExpression("hil_consumertype");
                    Query.ColumnSet = new ColumnSet("hil_consumertypeid");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "B2C");
                    entCol = service.RetrieveMultiple(Query);
                    if (entCol.Entities.Count > 0)
                    {
                        enWorkorder["hil_consumertype"] = entCol.Entities[0].ToEntityReference();
                    }

                    Query = new QueryExpression("hil_consumercategory");
                    Query.ColumnSet = new ColumnSet("hil_consumercategoryid");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "End User");
                    entCol = service.RetrieveMultiple(Query);
                    if (entCol.Entities.Count > 0)
                    {
                        enWorkorder["hil_consumercategory"] = entCol.Entities[0].ToEntityReference();
                    }

                    if (erNatureOfComplaint != null)
                    {
                        enWorkorder["hil_natureofcomplaint"] = erNatureOfComplaint;
                    }
                    if (callSubTypeGuid != Guid.Empty)
                    {
                        enWorkorder["hil_callsubtype"] = new EntityReference("hil_callsubtype", callSubTypeGuid);
                    }
                    enWorkorder["hil_quantity"] = 1;
                    enWorkorder["hil_customercomplaintdescription"] = serviceCalldata.ChiefComplaint;
                    enWorkorder["hil_callertype"] = new OptionSetValue(910590001);// {SourceofJob:"Customer"}

                    enWorkorder["hil_sourceofjob"] = new OptionSetValue(serviceCalldata.SourceOfJob); // SourceofJob:[{"6": "Dealer Portal"}]

                    enWorkorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {ServiceAccount:"Dummy Account"}
                    enWorkorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {BillingAccount:"Dummy Account"}
                    //enWorkorder["msdyn_primaryincidenttype"] = new EntityReference("msdyn_incidenttype", new Guid("0F5E8009-3BFD-E811-A94C-000D3AF0694E")); // {Primary Incident Type:"Installation -Decorative FAN CF"}

                    serviceCallGuid = service.Create(enWorkorder);
                    if (serviceCallGuid != Guid.Empty)
                    {
                        objServiceCall.JobGuid = serviceCallGuid;
                        objServiceCall.JobId = service.Retrieve("msdyn_workorder", serviceCallGuid, new ColumnSet("msdyn_name")).GetAttributeValue<string>("msdyn_name");
                        objServiceCall.StatusCode = "200";
                        objServiceCall.StatusDescription = "OK";
                    }
                    else
                    {
                        objServiceCall.StatusCode = "204";
                        objServiceCall.StatusDescription = "FAILURE !!! Something went wrong";
                    }
                    return objServiceCall;
                    #endregion
                }
                else
                {
                    objServiceCall = new IoTServiceCallRegistration { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    return objServiceCall;
                }
            }
            catch (Exception ex)
            {
                objServiceCall = new IoTServiceCallRegistration { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                return objServiceCall;
            }
        }
    }

    public class IoTServiceCallRegistration
    {

        public string SerialNumber { get; set; }

        public string ProductModelNumber { get; set; }

        public Guid NOCGuid { get; set; }

        public string NOCName { get; set; }

        public Guid ProductCategoryGuid { get; set; }


        public Guid ProductSubCategoryGuid { get; set; }


        public string ChiefComplaint { get; set; }

        public Guid AddressGuid { get; set; }

        public Guid AssetGuid { get; set; }

        public string CustomerMobleNo { get; set; }

        public Guid CustomerGuid { get; set; }

        public Guid JobGuid { get; set; }

        public string JobId { get; set; }

        public string ImageBase64String { get; set; }

        public int ImageType { get; set; }

        public int SourceOfJob { get; set; }

        public string PreferredDate { get; set; }

        public int PreferredPartOfDay { get; set; }

        public string StatusCode { get; set; }


        public string StatusDescription { get; set; }
    }

    public class IoTServiceCallResult
    {

        public string JobId { get; set; }


        public Guid JobGuid { get; set; }


        public string CallSubType { get; set; }


        public string JobLoggedon { get; set; }


        public string JobStatus { get; set; }


        public string JobAssignedTo { get; set; }


        public string CustomerAsset { get; set; }


        public string ProductCategory { get; set; }


        public string NatureOfComplaint { get; set; }


        public string JobClosedOn { get; set; }


        public string CustomerName { get; set; }


        public string ServiceAddress { get; set; }


        public string Product { get; set; }


        public string ChiefComplaint { get; set; }

        public string PreferredDate { get; set; }

        public int PreferredPartOfDay { get; set; }

        public string PreferredPartOfDayName { get; set; }

        public string StatusCode { get; set; }


        public string StatusDescription { get; set; }
    }

    public class IoTNatureofComplaint
    {

        public string SerialNumber { get; set; }


        public Guid ProductSubCategoryId { get; set; }


        public string Name { get; set; }

        public Guid Guid { get; set; }

        public string StatusCode { get; set; }

        public string StatusDescription { get; set; }
    }
}
