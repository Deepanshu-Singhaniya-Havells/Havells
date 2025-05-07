using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System.Runtime.Serialization;
using System.Security.Cryptography;
using System.IO;
using System.Text.RegularExpressions;
using System.Linq;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class IotServiceCall
    {
        [DataMember]
        public Guid CustomerGuid { get; set; }

        public List<IoTServiceCallResult> GetIoTServiceCalls(IotServiceCall job)
        {
            List<IoTServiceCallResult> jobList = new List<IoTServiceCallResult>();
            IoTServiceCallResult objJobOutput;

            try
            {
                if (job.CustomerGuid.ToString().Trim().Length == 0)
                {
                    objJobOutput = new IoTServiceCallResult { StatusCode = "204", StatusDescription = "Customer GUID is required." };
                    jobList.Add(objJobOutput);
                    return jobList;
                }
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    QueryExpression query = new QueryExpression()
                    {
                        EntityName = msdyn_workorder.EntityLogicalName,
                        ColumnSet = new ColumnSet(true)
                    };
                    FilterExpression filterExpression = new FilterExpression(LogicalOperator.And);
                    filterExpression.Conditions.Add(new ConditionExpression("hil_customerref", ConditionOperator.Equal, job.CustomerGuid));
                    query.Criteria.AddFilter(filterExpression);
                    query.AddOrder("createdon", OrderType.Descending);

                    EntityCollection collection = service.RetrieveMultiple(query);

                    if (collection.Entities != null && collection.Entities.Count > 0)
                    {
                        foreach (Entity item in collection.Entities)
                        {
                            IoTServiceCallResult jobObj = new IoTServiceCallResult();
                            if (item.Attributes.Contains("msdyn_name"))
                            {
                                jobObj.JobId = item.GetAttributeValue<string>("msdyn_name");
                            }
                            if (item.Attributes.Contains("msdyn_name"))
                            {
                                jobObj.JobGuid = item.Id;
                            }
                            if (item.Attributes.Contains("hil_callsubtype"))
                            {
                                jobObj.CallSubType = item.GetAttributeValue<EntityReference>("hil_callsubtype").Name;
                            }
                            if (item.Attributes.Contains("createdon"))
                            {
                                jobObj.JobLoggedon = item.GetAttributeValue<DateTime>("createdon").AddMinutes(330).ToString();
                            }
                            if (item.Attributes.Contains("msdyn_substatus"))
                            {
                                jobObj.JobStatus = item.GetAttributeValue<EntityReference>("msdyn_substatus").Name;
                            }
                            if (item.Attributes.Contains("hil_owneraccount"))
                            {
                                jobObj.JobAssignedTo = item.GetAttributeValue<EntityReference>("hil_owneraccount").Name;
                            }
                            if (item.Attributes.Contains("msdyn_customerasset"))
                            {
                                jobObj.CustomerAsset = item.GetAttributeValue<EntityReference>("msdyn_customerasset").Name;
                            }
                            if (item.Attributes.Contains("hil_productcategory"))
                            {
                                jobObj.ProductCategory = item.GetAttributeValue<EntityReference>("hil_productcategory").Name;
                            }
                            if (item.Attributes.Contains("hil_natureofcomplaint"))
                            {
                                jobObj.NatureOfComplaint = item.GetAttributeValue<EntityReference>("hil_natureofcomplaint").Name;
                            }
                            if (item.Attributes.Contains("hil_jobclosuredon"))
                            {
                                jobObj.JobClosedOn = item.GetAttributeValue<DateTime>("hil_jobclosuredon").AddMinutes(330).ToString();
                            }
                            if (item.Attributes.Contains("hil_customerref"))
                            {
                                jobObj.CustomerName = item.GetAttributeValue<EntityReference>("hil_customerref").Name;
                            }
                            if (item.Attributes.Contains("hil_fulladdress"))
                            {
                                jobObj.ServiceAddress = item.GetAttributeValue<string>("hil_fulladdress");
                            }
                            if (item.Attributes.Contains("msdyn_customerasset"))
                            {
                                Entity ec = service.Retrieve("msdyn_customerasset", item.GetAttributeValue<EntityReference>("msdyn_customerasset").Id, new ColumnSet("hil_modelname"));
                                if (ec != null)
                                {
                                    jobObj.Product = ec.GetAttributeValue<string>("hil_modelname");
                                }
                            }
                            if (item.Attributes.Contains("hil_customercomplaintdescription"))
                            {
                                jobObj.ChiefComplaint = item.GetAttributeValue<string>("hil_customercomplaintdescription");
                            }
                            if (item.Attributes.Contains("hil_preferredtime"))
                            {
                                jobObj.PreferredPartOfDay = item.GetAttributeValue<OptionSetValue>("hil_preferredtime").Value;
                                jobObj.PreferredPartOfDayName = item.FormattedValues["hil_preferredtime"].ToString();
                            }
                            if (item.Attributes.Contains("hil_preferreddate"))
                            {
                                jobObj.PreferredDate = item.GetAttributeValue<DateTime>("hil_preferreddate").AddMinutes(330).ToShortDateString();
                            }
                            jobObj.StatusCode = "200";
                            jobObj.StatusDescription = "OK";
                            jobList.Add(jobObj);
                        }
                    }
                    return jobList;
                }
                else
                {
                    objJobOutput = new IoTServiceCallResult { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    jobList.Add(objJobOutput);
                }
            }
            catch (Exception ex)
            {
                objJobOutput = new IoTServiceCallResult { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                jobList.Add(objJobOutput);
            }
            return jobList;
        }
        public IoTServiceCallRegistration IoTCreateServiceCall(IoTServiceCallRegistration serviceCalldata)
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
            EntityReference erProductCategory = null;
            EntityReference erProductsubcategory = null;
            EntityReference erProductsubcategorymapping = null;
            EntityReference erNatureOfComplaint = null;
            EntityReference erCustomerAsset = null;
            bool continueFlag = false;
            string fullAddress = string.Empty;
            Regex Regex_MobileNo = new Regex("^[6-9]\\d{9}$");
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (string.IsNullOrWhiteSpace(serviceCalldata.CustomerMobleNo))
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer Mobile Number is required." };
                    }
                    else if (!Regex_MobileNo.IsMatch(serviceCalldata.CustomerMobleNo))
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Invalid Customer Mobile Number." };
                    }
                    if (serviceCalldata.CustomerGuid == Guid.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer Guid is required." };
                    }
                    if (serviceCalldata.NOCGuid == Guid.Empty)
                    {
                        if (serviceCalldata.SourceOfJob != 22)
                        {
                            return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Nature of Complaint is required." };
                        }
                    }
                    if (serviceCalldata.AddressGuid == Guid.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Consumer Address is required." };
                    }

                    if (string.IsNullOrWhiteSpace(serviceCalldata.SerialNumber))
                    {
                        if (serviceCalldata.SourceOfJob == 22)
                        {
                            if (!string.IsNullOrEmpty(serviceCalldata.ProductModelNumber))
                            {
                                Query = new QueryExpression("product");
                                Query.ColumnSet = new ColumnSet("hil_division", "hil_materialgroup", "productnumber");
                                Query.Criteria.AddCondition("productnumber", ConditionOperator.Equal, serviceCalldata.ProductModelNumber);
                                EntityCollection dataCollection = service.RetrieveMultiple(Query);
                                if (dataCollection.Entities.Count > 0)
                                {
                                    serviceCalldata.ProductCategoryGuid = dataCollection.Entities[0].GetAttributeValue<EntityReference>("hil_division").Id;
                                    serviceCalldata.ProductSubCategoryGuid = dataCollection.Entities[0].GetAttributeValue<EntityReference>("hil_materialgroup").Id;
                                }
                                else
                                {
                                    return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Invalid ProductModelNumber." };
                                }
                            }
                            else if (serviceCalldata.ProductCategoryGuid == Guid.Empty || serviceCalldata.ProductSubCategoryGuid == Guid.Empty)
                            {
                                return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Serial Number or Product Modelnumber or Product Category and Product SubCategory is required" };
                            }
                        }
                        if (serviceCalldata.ProductCategoryGuid == Guid.Empty)
                        {
                            return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Product Category is required." };
                        }
                        else if (serviceCalldata.ProductSubCategoryGuid == Guid.Empty)
                        {
                            return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Product Sub Category is required." };
                        }
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
                    //else {
                    //    fullAddress = entcoll.Entities[0].GetAttributeValue<string>("");
                    //}

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
                    if (string.IsNullOrWhiteSpace(serviceCalldata.ChiefComplaint))
                    {
                        if (serviceCalldata.SourceOfJob != 22)
                        {
                            return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer's Chief Complaint is required." };
                        }
                    }
                    //Case 1 Serial Number Exists
                    if (serviceCalldata.SerialNumber != null)
                    {
                        string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_customerasset'>
                            <attribute name='msdyn_name' />
                            <attribute name='hil_customer' />
                            <attribute name='hil_productsubcategorymapping' />
                            <attribute name='hil_productcategory' />
                            <attribute name='hil_productsubcategory' />
                            <attribute name='msdyn_customerassetid' />
                            <attribute name='hil_invoicedate' />
                        <order attribute='msdyn_name' descending='false' />
	                    <filter type='and'>
                            <condition attribute='hil_customer' operator='eq' value='{" + serviceCalldata.CustomerGuid.ToString() + @"}' />
                            <condition attribute='msdyn_name' operator='eq' value='" + serviceCalldata.SerialNumber + @"' />
                        </filter>
                        </entity>
                        </fetch>";
                        entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        if (entcoll.Entities.Count > 0)
                        {
                            erCustomerAsset = entcoll.Entities[0].ToEntityReference();
                            modelName = entcoll.Entities[0].GetAttributeValue<string>("msdyn_name");
                            invoiceDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_invoicedate");
                            erProductCategory = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_productcategory");
                            erProductsubcategory = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_productsubcategory");
                            erProductsubcategorymapping = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_productsubcategorymapping");
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
                    int[] ArraySourceOfJob = new int[] { 12, 13, 16, 22 };
                    // Input<12> for Whatsapp, <13> for IoT Platform, <16> for Chatbot and <22> for OneWebsite.
                    if (!ArraySourceOfJob.Contains(serviceCalldata.SourceOfJob))
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
                        fetchXML += "<condition attribute='statecode' operator='eq' value='0' />";
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
                    else if (serviceCalldata.SourceOfJob == 22)
                    {
                        string[] Callsubtype = new string[] { "I", "B", "D" };
                        if (Callsubtype.Contains(serviceCalldata.CallSubType))
                        {
                            fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>";
                            fetchXML += "<entity name='hil_natureofcomplaint'>";
                            fetchXML += "<attribute name='hil_callsubtype' />";
                            fetchXML += "<attribute name='hil_natureofcomplaintid' />";
                            fetchXML += "<order attribute='createdon' descending='false' />";
                            fetchXML += "<filter type='and'>";
                            fetchXML += "<condition attribute='statecode' operator='eq' value='0' />";
                            fetchXML += "<condition attribute='hil_relatedproduct' operator='eq' value='{" + erProductsubcategory.Id + "}' />";
                            fetchXML += "<condition attribute='hil_callsubtype' operator='in'>";
                            fetchXML += "<value uiname='Demo' uitype='hil_callsubtype'>{AE1B2B71-3C0B-E911-A94E-000D3AF06CD4}</value>";
                            fetchXML += "<value uiname='Installation' uitype='hil_callsubtype'>{E3129D79-3C0B-E911-A94E-000D3AF06CD4}</value></condition>";
                            fetchXML += "</filter>";
                            fetchXML += "</entity>";
                            fetchXML += "</fetch>";
                            entcoll = service.RetrieveMultiple(new FetchExpression(fetchXML));
                            if (entcoll.Entities.Count > 0)
                            {
                                if (serviceCalldata.CallSubType == "D")
                                {
                                    callSubTypeGuid = new Guid("AE1B2B71-3C0B-E911-A94E-000D3AF06CD4");//Demo
                                }
                                else
                                {
                                    callSubTypeGuid = new Guid("E3129D79-3C0B-E911-A94E-000D3AF06CD4");//Installation
                                }
                            }
                        }
                        else
                        {
                            return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Invalid Call Sub Type." };
                        }
                    }
                    #endregion
                    #region Create Service Call
                    Int32 warrantyStatus = 2;
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
                    //if (invoiceDate != null)
                    //{
                    //    enWorkorder["hil_purchasedate"] = invoiceDate;
                    //}
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
                    //enWorkorder["hil_warrantystatus"] = new OptionSetValue(warrantyStatus);
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

                    enWorkorder["hil_sourceofjob"] = new OptionSetValue(serviceCalldata.SourceOfJob); // SourceofJob:[{"12": "WhatsApp"} ,{"13","IoT Platform"},{"16","Chatbot"}]

                    enWorkorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {ServiceAccount:"Dummy Account"}
                    enWorkorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {BillingAccount:"Dummy Account"}
                    //enWorkorder["msdyn_primaryincidenttype"] = new EntityReference("msdyn_incidenttype", new Guid("0F5E8009-3BFD-E811-A94C-000D3AF0694E")); // {Primary Incident Type:"Installation&nbsp;-Decorative&nbsp;FAN&nbsp;CF"}

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

        //public IoTServiceCallRegistration IoTCreateServiceCall(IoTServiceCallRegistration serviceCalldata)
        //{
        //    IoTServiceCallRegistration objServiceCall;
        //    Guid customerGuid = Guid.Empty;
        //    Guid callSubTypeGuid = Guid.Empty;
        //    Guid serviceCallGuid = Guid.Empty;
        //    Entity lookupObj = null;
        //    EntityCollection entcoll;
        //    QueryExpression Query;
        //    string customerFullName = string.Empty;
        //    string customerMobileNumber = string.Empty;
        //    string customerEmail = string.Empty;
        //    Guid customerAssetGuid = Guid.Empty;
        //    DateTime? invoiceDate = null;
        //    string modelName = string.Empty;
        //    EntityReference erProductCategory = null;
        //    EntityReference erProductsubcategory = null;
        //    EntityReference erProductsubcategorymapping = null;
        //    EntityReference erNatureOfComplaint = null;
        //    EntityReference erCustomerAsset = null;
        //    bool continueFlag = false;
        //    string fullAddress = string.Empty;

        //    try
        //    {
        //        IOrganizationService service = ConnectToCRM.GetOrgService();
        //        if (service != null)
        //        {
        //            if (serviceCalldata.CustomerMobleNo == string.Empty)
        //            {
        //                return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer Nobile Number is required." };
        //            }
        //            if (serviceCalldata.CustomerGuid == Guid.Empty)
        //            {
        //                return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer Guid is required." };
        //            }
        //            if (serviceCalldata.NOCGuid == Guid.Empty)
        //            {
        //                return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Nature of Complaint is required." };
        //            }
        //            if (serviceCalldata.AddressGuid == Guid.Empty)
        //            {
        //                return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Consumer Address is required." };
        //            }

        //            if (serviceCalldata.SerialNumber == null || serviceCalldata.SerialNumber == string.Empty || serviceCalldata.SerialNumber.Trim().Length == 0)
        //            {
        //                if (serviceCalldata.ProductCategoryGuid == Guid.Empty)
        //                {
        //                    return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Product Category is required." };
        //                }
        //                else if (serviceCalldata.ProductSubCategoryGuid == Guid.Empty)
        //                {
        //                    return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Product Sub Category is required." };
        //                }
        //            }
        //            Query = new QueryExpression("hil_address");
        //            Query.ColumnSet = new ColumnSet("hil_addressid");
        //            Query.Criteria = new FilterExpression(LogicalOperator.And);
        //            Query.Criteria.AddCondition("hil_addressid", ConditionOperator.Equal, serviceCalldata.AddressGuid);
        //            Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, serviceCalldata.CustomerGuid);
        //            entcoll = service.RetrieveMultiple(Query);
        //            if (entcoll.Entities.Count == 0)
        //            {
        //                return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Address does not belong to Customer." };
        //            }
        //            //else {
        //            //    fullAddress = entcoll.Entities[0].GetAttributeValue<string>("");
        //            //}

        //            Query = new QueryExpression("contact");
        //            Query.ColumnSet = new ColumnSet("fullname", "emailaddress1", "mobilephone");
        //            Query.Criteria = new FilterExpression(LogicalOperator.And);
        //            Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, serviceCalldata.CustomerMobleNo);
        //            Query.Criteria.AddCondition("contactid", ConditionOperator.Equal, serviceCalldata.CustomerGuid);
        //            entcoll = service.RetrieveMultiple(Query);
        //            if (entcoll.Entities.Count == 0)
        //            {
        //                return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer/Mobile No. does not exist." };
        //            }
        //            else
        //            {
        //                lookupObj = entcoll.Entities[0];
        //                customerFullName = lookupObj.GetAttributeValue<string>("fullname");
        //                customerEmail = lookupObj.GetAttributeValue<string>("emailaddress1");
        //                customerMobileNumber = lookupObj.GetAttributeValue<string>("mobilephone"); // N
        //            }
        //            if (serviceCalldata.ChiefComplaint == string.Empty || serviceCalldata.ChiefComplaint == null || serviceCalldata.ChiefComplaint.Trim().Length == 0)
        //            {
        //                return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer's Chief Complaint is required." };
        //            }
        //            //Case 1 Serial Number Exists
        //            if (serviceCalldata.SerialNumber != null)
        //            {
        //                string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                <entity name='msdyn_customerasset'>
        //                    <attribute name='msdyn_name' />
        //                    <attribute name='hil_customer' />
        //                    <attribute name='hil_productsubcategorymapping' />
        //                    <attribute name='hil_productcategory' />
        //                    <attribute name='hil_productsubcategory' />
        //                    <attribute name='msdyn_customerassetid' />
        //                    <attribute name='hil_invoicedate' />
        //                <order attribute='msdyn_name' descending='false' />
        //             <filter type='and'>
        //                    <condition attribute='hil_customer' operator='eq' value='{" + serviceCalldata.CustomerGuid.ToString() + @"}' />
        //                    <condition attribute='msdyn_name' operator='eq' value='" + serviceCalldata.SerialNumber + @"' />
        //                </filter>
        //                </entity>
        //                </fetch>";
        //                entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
        //                if (entcoll.Entities.Count > 0)
        //                {
        //                    erCustomerAsset = entcoll.Entities[0].ToEntityReference();
        //                    modelName = entcoll.Entities[0].GetAttributeValue<string>("msdyn_name");
        //                    invoiceDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_invoicedate");
        //                    erProductCategory = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_productcategory");
        //                    erProductsubcategory = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_productsubcategory");
        //                    erProductsubcategorymapping = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_productsubcategorymapping");
        //                    continueFlag = true;
        //                }
        //                else
        //                {
        //                    return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer Asset Serial Number does not exist." };
        //                }
        //            }
        //            //Case 2 Product Category 
        //            else if (serviceCalldata.ProductCategoryGuid != Guid.Empty)
        //            {
        //                erProductCategory = new EntityReference("product", serviceCalldata.ProductCategoryGuid);
        //                erProductsubcategory = new EntityReference("product", serviceCalldata.ProductSubCategoryGuid);
        //                modelName = string.Empty;
        //                continueFlag = true;
        //            }

        //            if (!continueFlag)
        //            {
        //                return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Asset Serial Number/Product Category is required to proceed." };
        //            }
        //            if (serviceCalldata.SourceOfJob != 12 && serviceCalldata.SourceOfJob != 13 && serviceCalldata.SourceOfJob != 16)
        //            {
        //                return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Invalid Source of Job!!! Please Input <12> for Whatsapp for Service <13> for IoT Platform <16> for Chatbot." };
        //            }

        //            #region Get Nature of Complaint
        //            string fetchXML = string.Empty;

        //            if (serviceCalldata.NOCGuid != Guid.Empty)
        //            {
        //                fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>";
        //                fetchXML += "<entity name='hil_natureofcomplaint'>";
        //                fetchXML += "<attribute name='hil_callsubtype' />";
        //                fetchXML += "<attribute name='hil_natureofcomplaintid' />";
        //                fetchXML += "<order attribute='createdon' descending='false' />";
        //                fetchXML += "<filter type='and'>";
        //                fetchXML += "<condition attribute='statecode' operator='eq' value='0' />";
        //                fetchXML += "<condition attribute='hil_relatedproduct' operator='eq' value='{" + erProductsubcategory.Id + "}' />";
        //                fetchXML += "<condition attribute='hil_natureofcomplaintid' operator='eq' value='{" + serviceCalldata.NOCGuid + "}' />";
        //                fetchXML += "</filter>";
        //                fetchXML += "</entity>";
        //                fetchXML += "</fetch>";
        //                entcoll = service.RetrieveMultiple(new FetchExpression(fetchXML));
        //                if (entcoll.Entities.Count > 0)
        //                {
        //                    erNatureOfComplaint = entcoll.Entities[0].ToEntityReference();
        //                    if (entcoll.Entities[0].Attributes.Contains("hil_callsubtype"))
        //                    {
        //                        callSubTypeGuid = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_callsubtype").Id;
        //                    }
        //                }
        //                else
        //                {
        //                    return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "NOC does not match with Product Sub Category." };
        //                }
        //            }
        //            #endregion
        //            #region Create Service Call
        //            Int32 warrantyStatus = 2;
        //            objServiceCall = new IoTServiceCallRegistration();
        //            objServiceCall = serviceCalldata;
        //            Entity enWorkorder = new Entity("msdyn_workorder");

        //            if (serviceCalldata.CustomerGuid != Guid.Empty)
        //            {
        //                enWorkorder["hil_customerref"] = new EntityReference("contact", serviceCalldata.CustomerGuid);
        //            }
        //            enWorkorder["hil_customername"] = customerFullName;
        //            enWorkorder["hil_mobilenumber"] = customerMobileNumber;
        //            enWorkorder["hil_email"] = customerEmail;

        //            if (serviceCalldata.PreferredPartOfDay > 0 && serviceCalldata.PreferredPartOfDay < 4)
        //            {
        //                enWorkorder["hil_preferredtime"] = new OptionSetValue(serviceCalldata.PreferredPartOfDay);
        //            }

        //            if (serviceCalldata.PreferredDate != null && serviceCalldata.PreferredDate.Trim().Length > 0)
        //            {
        //                string _date = serviceCalldata.PreferredDate;
        //                DateTime dtInvoice = new DateTime(Convert.ToInt32(_date.Substring(6, 4)), Convert.ToInt32(_date.Substring(0, 2)), Convert.ToInt32(_date.Substring(3, 2)));
        //                enWorkorder["hil_preferreddate"] = dtInvoice;
        //            }

        //            if (serviceCalldata.AddressGuid != Guid.Empty)
        //            {
        //                enWorkorder["hil_address"] = new EntityReference("hil_address", serviceCalldata.AddressGuid);
        //            }

        //            if (erCustomerAsset != null)
        //            {
        //                enWorkorder["msdyn_customerasset"] = erCustomerAsset;
        //            }
        //            //if (invoiceDate != null)
        //            //{
        //            //    enWorkorder["hil_purchasedate"] = invoiceDate;
        //            //}
        //            if (modelName != string.Empty)
        //            {
        //                enWorkorder["hil_modelname"] = modelName;
        //            }

        //            if (erProductCategory != null)
        //            {
        //                enWorkorder["hil_productcategory"] = erProductCategory;
        //            }
        //            if (erProductsubcategory != null)
        //            {
        //                enWorkorder["hil_productsubcategory"] = erProductsubcategory;
        //            }

        //            Query = new QueryExpression("hil_stagingdivisonmaterialgroupmapping");
        //            Query.ColumnSet = new ColumnSet("hil_stagingdivisonmaterialgroupmappingid");
        //            Query.Criteria = new FilterExpression(LogicalOperator.And);
        //            Query.Criteria.AddCondition("hil_productcategorydivision", ConditionOperator.Equal, erProductCategory.Id);
        //            Query.Criteria = new FilterExpression(LogicalOperator.And);
        //            Query.Criteria.AddCondition("hil_productsubcategorymg", ConditionOperator.Equal, erProductsubcategory.Id);
        //            EntityCollection ec = service.RetrieveMultiple(Query);
        //            if (ec.Entities.Count > 0)
        //            {
        //                enWorkorder["hil_productcatsubcatmapping"] = ec.Entities[0].ToEntityReference();
        //            }

        //            EntityCollection entCol;
        //            //enWorkorder["hil_warrantystatus"] = new OptionSetValue(warrantyStatus);
        //            Query = new QueryExpression("hil_consumertype");
        //            Query.ColumnSet = new ColumnSet("hil_consumertypeid");
        //            Query.Criteria = new FilterExpression(LogicalOperator.And);
        //            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "B2C");
        //            entCol = service.RetrieveMultiple(Query);
        //            if (entCol.Entities.Count > 0)
        //            {
        //                enWorkorder["hil_consumertype"] = entCol.Entities[0].ToEntityReference();
        //            }

        //            Query = new QueryExpression("hil_consumercategory");
        //            Query.ColumnSet = new ColumnSet("hil_consumercategoryid");
        //            Query.Criteria = new FilterExpression(LogicalOperator.And);
        //            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "End User");
        //            entCol = service.RetrieveMultiple(Query);
        //            if (entCol.Entities.Count > 0)
        //            {
        //                enWorkorder["hil_consumercategory"] = entCol.Entities[0].ToEntityReference();
        //            }

        //            if (erNatureOfComplaint != null)
        //            {
        //                enWorkorder["hil_natureofcomplaint"] = erNatureOfComplaint;
        //            }
        //            if (callSubTypeGuid != Guid.Empty)
        //            {
        //                enWorkorder["hil_callsubtype"] = new EntityReference("hil_callsubtype", callSubTypeGuid);
        //            }
        //            enWorkorder["hil_quantity"] = 1;
        //            enWorkorder["hil_customercomplaintdescription"] = serviceCalldata.ChiefComplaint;
        //            enWorkorder["hil_callertype"] = new OptionSetValue(910590001);// {SourceofJob:"Customer"}

        //            enWorkorder["hil_sourceofjob"] = new OptionSetValue(serviceCalldata.SourceOfJob); // SourceofJob:[{"12": "WhatsApp"} ,{"13","IoT Platform"},{"16","Chatbot"}]

        //            enWorkorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {ServiceAccount:"Dummy Account"}
        //            enWorkorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {BillingAccount:"Dummy Account"}
        //            //enWorkorder["msdyn_primaryincidenttype"] = new EntityReference("msdyn_incidenttype", new Guid("0F5E8009-3BFD-E811-A94C-000D3AF0694E")); // {Primary Incident Type:"Installation -Decorative FAN CF"}

        //            serviceCallGuid = service.Create(enWorkorder);
        //            if (serviceCallGuid != Guid.Empty)
        //            {
        //                objServiceCall.JobGuid = serviceCallGuid;
        //                objServiceCall.JobId = service.Retrieve("msdyn_workorder", serviceCallGuid, new ColumnSet("msdyn_name")).GetAttributeValue<string>("msdyn_name");
        //                objServiceCall.StatusCode = "200";
        //                objServiceCall.StatusDescription = "OK";
        //            }
        //            else
        //            {
        //                objServiceCall.StatusCode = "204";
        //                objServiceCall.StatusDescription = "FAILURE !!! Something went wrong";
        //            }
        //            return objServiceCall;
        //            #endregion
        //        }
        //        else
        //        {
        //            objServiceCall = new IoTServiceCallRegistration { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
        //            return objServiceCall;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        objServiceCall = new IoTServiceCallRegistration { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
        //        return objServiceCall;
        //    }
        //}

        public IoTServiceCallRegistration IoTCreateServiceCallWhatsapp(IoTServiceCallRegistration serviceCalldata)
        {
            IoTServiceCallRegistration objServiceCall;
            Guid customerGuid = Guid.Empty;
            Guid callSubTypeGuid = new Guid("6560565a-3c0b-e911-a94e-000d3af06cd4"); //BREAKDOWN
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
            EntityReference erProductCategory = new EntityReference("product", new Guid("B96F747D-16FA-E811-A94C-000D3AF0677F")); //OTHERS
            EntityReference erProductsubcategory = new EntityReference("product", new Guid("4AABBB57-A85E-EA11-A811-000D3AF057DD")); //OTHERS
            EntityReference erProductsubcategorymapping = null;
            EntityReference erNatureOfComplaint = new EntityReference("hil_natureofcomplaint", new Guid("6DB33D1A-8E5F-EA11-A811-000D3AF05D7B")); //BREAKDOWN
            EntityReference erCustomerAsset = null;
            bool continueFlag = false;
            string fullAddress = string.Empty;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
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
                    if (serviceCalldata.NOCName == string.Empty || serviceCalldata.NOCName == null || serviceCalldata.NOCName.Trim().Length == 0)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Nature of Complaint is required." };
                    }
                    if (serviceCalldata.AddressGuid == Guid.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Consumer Address is required." };
                    }

                    lookupObj = service.Retrieve("contact", serviceCalldata.CustomerGuid, new ColumnSet("fullname", "emailaddress1", "mobilephone"));
                    if (lookupObj == null)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Consumer Guid does not exist." };
                    }
                    else
                    {
                        customerFullName = lookupObj.GetAttributeValue<string>("fullname");
                        customerEmail = lookupObj.GetAttributeValue<string>("emailaddress1");
                        customerMobileNumber = lookupObj.GetAttributeValue<string>("mobilephone"); // N
                    }
                    if (serviceCalldata.ChiefComplaint == string.Empty || serviceCalldata.ChiefComplaint == null || serviceCalldata.ChiefComplaint.Trim().Length == 0)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer's Chief Complaint is required." };
                    }
                    //Case 1 Serial Number Exists
                    if (serviceCalldata.SerialNumber != null)
                    {
                        string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_customerasset'>
                            <attribute name='msdyn_name' />
                            <attribute name='hil_customer' />
                            <attribute name='hil_productsubcategorymapping' />
                            <attribute name='hil_productcategory' />
                            <attribute name='hil_productsubcategory' />
                            <attribute name='msdyn_customerassetid' />
                            <attribute name='hil_invoicedate' />
                        <order attribute='msdyn_name' descending='false' />
	                    <filter type='and'>
                            <condition attribute='hil_customer' operator='eq' value='{" + serviceCalldata.CustomerGuid.ToString() + @"}' />
                            <condition attribute='msdyn_name' operator='eq' value='" + serviceCalldata.SerialNumber + @"' />
                        </filter>
                        </entity>
                        </fetch>";
                        entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        if (entcoll.Entities.Count > 0)
                        {
                            erCustomerAsset = entcoll.Entities[0].ToEntityReference();
                            modelName = entcoll.Entities[0].GetAttributeValue<string>("msdyn_name");
                            invoiceDate = entcoll.Entities[0].GetAttributeValue<DateTime>("hil_invoicedate");
                            erProductCategory = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_productcategory");
                            erProductsubcategory = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_productsubcategory");
                            erProductsubcategorymapping = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_productsubcategorymapping");
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

                        string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='hil_whatsappproductdivisionconfig'>
                            <attribute name='hil_productmaterialgroup' />
	                        <filter type='and'>
                                <condition attribute='hil_productdivision' operator='eq' value='{" + serviceCalldata.ProductCategoryGuid + @"}' />
                            </filter>
                            </entity>
                            </fetch>";
                        entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        if (entcoll.Entities.Count > 0)
                        {
                            modelName = string.Empty;
                            erProductsubcategory = entcoll.Entities[0].GetAttributeValue<EntityReference>("hil_productmaterialgroup");
                            continueFlag = true;
                        }
                        else
                        {
                            return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "No sub Category mapped with Product Category you have selected." };
                        }
                    }
                    else
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Product Category is required to proceed." };
                    }

                    if (!continueFlag)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Asset Serial Number/Product Model Number/Product Category is required to proceed." };
                    }
                    if (serviceCalldata.SourceOfJob != 12 && serviceCalldata.SourceOfJob != 13)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Invalid Source of Job!!! Please Input <12> for Whatsapp for Service <13> for IoT Platform." };
                    }

                    lookupObj = service.Retrieve("hil_address", serviceCalldata.AddressGuid, new ColumnSet("hil_fulladdress"));
                    if (lookupObj == null)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Consumer Address does not exist." };
                    }
                    else
                    {
                        fullAddress = lookupObj.GetAttributeValue<string>("hil_fulladdress");
                    }
                    #region Get Nature of Complaint
                    string fetchXML = string.Empty;

                    if (serviceCalldata.NOCName.ToUpper().IndexOf("INSTALL") >= 0 || serviceCalldata.ChiefComplaint.ToUpper().IndexOf("INSTALL") >= 0)
                    {
                        fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>";
                        fetchXML += "<entity name='hil_natureofcomplaint'>";
                        fetchXML += "<attribute name='hil_callsubtype' />";
                        fetchXML += "<attribute name='hil_natureofcomplaintid' />";
                        fetchXML += "<order attribute='createdon' descending='false' />";
                        fetchXML += "<filter type='and'>";
                        fetchXML += "<condition attribute='hil_relatedproduct' operator='eq' value='{" + erProductsubcategory.Id + "}' />";
                        fetchXML += "<condition attribute='hil_callsubtypename' operator='like' value='%install%' />";
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
                    }
                    else
                    {
                        fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>";
                        fetchXML += "<entity name='hil_natureofcomplaint'>";
                        fetchXML += "<attribute name='hil_callsubtype' />";
                        fetchXML += "<attribute name='hil_natureofcomplaintid' />";
                        fetchXML += "<order attribute='createdon' descending='false' />";
                        fetchXML += "<filter type='and'>";
                        fetchXML += "<condition attribute='hil_relatedproduct' operator='eq' value='{" + erProductsubcategory.Id + "}' />";
                        fetchXML += "<condition attribute='hil_callsubtypename' operator='like' value='%break%' />";
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
                    }
                    #endregion
                    #region Create Service Call
                    Int32 warrantyStatus = 2;
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

                    if (serviceCalldata.AddressGuid != Guid.Empty)
                    {
                        enWorkorder["hil_address"] = new EntityReference("hil_address", serviceCalldata.AddressGuid);
                    }

                    if (erCustomerAsset != null)
                    {
                        enWorkorder["msdyn_customerasset"] = erCustomerAsset;
                    }
                    //if (invoiceDate != null)
                    //{
                    //    enWorkorder["hil_purchasedate"] = invoiceDate;
                    //}
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
                    //enWorkorder["hil_warrantystatus"] = new OptionSetValue(warrantyStatus);
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

                    enWorkorder["hil_sourceofjob"] = new OptionSetValue(serviceCalldata.SourceOfJob); // SourceofJob:[{"12": "WhatsApp"} ,{"13","IoT Platform"}]

                    enWorkorder["msdyn_serviceaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {ServiceAccount:"Dummy Account"}
                    enWorkorder["msdyn_billingaccount"] = new EntityReference("account", new Guid("B8168E04-7D0A-E911-A94F-000D3AF00F43")); // {BillingAccount:"Dummy Account"}
                    serviceCallGuid = service.Create(enWorkorder);
                    if (serviceCallGuid != Guid.Empty)
                    {
                        Guid addressId = Guid.Empty;
                        addressId = service.Retrieve("msdyn_workorder", serviceCallGuid, new ColumnSet("hil_address")).GetAttributeValue<EntityReference>("hil_address").Id;
                        if (addressId != serviceCalldata.AddressGuid)
                        {
                            msdyn_workorder jobObj = new msdyn_workorder();
                            jobObj.Id = serviceCallGuid;
                            jobObj.hil_Address = new EntityReference("hil_address", serviceCalldata.AddressGuid);
                            jobObj.hil_FullAddress = fullAddress;
                            service.Update(jobObj);
                        }
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

        public List<IoTNatureofComplaint> GetIoTNatureOfComplaints(IoTNatureofComplaint natureOfComplaint)
        {
            IoTNatureofComplaint objNatureOfComplaint;
            List<IoTNatureofComplaint> lstNatureOfComplaint = new List<IoTNatureofComplaint>();
            EntityCollection entcoll;
            QueryExpression Query;

            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (natureOfComplaint.SerialNumber.Trim().Length == 0)
                    {
                        objNatureOfComplaint = new IoTNatureofComplaint { StatusCode = "204", StatusDescription = "Product Serial Number is required." };
                        lstNatureOfComplaint.Add(objNatureOfComplaint);
                        return lstNatureOfComplaint;
                    }

                    Query = new QueryExpression("msdyn_customerasset");
                    Query.ColumnSet = new ColumnSet("msdyn_name");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, natureOfComplaint.SerialNumber);
                    Query.TopCount = 1;
                    entcoll = service.RetrieveMultiple(Query);

                    if (entcoll.Entities.Count == 0)
                    {
                        objNatureOfComplaint = new IoTNatureofComplaint { StatusCode = "204", StatusDescription = "Product Serial Number does not exist." };
                        lstNatureOfComplaint.Add(objNatureOfComplaint);
                        return lstNatureOfComplaint;
                    }
                    else
                    {
                        string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                        <entity name='hil_natureofcomplaint'>
                        <attribute name='hil_name' />
                        <attribute name='hil_natureofcomplaintid' />
                        <link-entity name='product' from='productid' to='hil_relatedproduct' link-type='inner' alias='ae'>
                            <link-entity name='msdyn_customerasset' from='hil_productsubcategory' to='productid' link-type='inner' alias='af'>
                                <filter type='and'>
                                    <condition attribute='msdyn_name' operator='eq' value='" + natureOfComplaint.SerialNumber + @"' />
                                </filter>
                            </link-entity>
                        </link-entity>
                        <link-entity name='hil_callsubtype' from='hil_callsubtypeid' to='hil_callsubtype' link-type='inner' alias='ac'>
                        <attribute name='hil_index' />
                        <order attribute='hil_index' descending='false' />
                        <filter type='and'>
                            <condition attribute='hil_index' operator='ne' value='99' />
                            <condition attribute='statecode' operator='eq' value='0' />
                        </filter>
                        </link-entity>
                        </entity>
                        </fetch>";
                        entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        if (entcoll.Entities.Count == 0)
                        {
                            objNatureOfComplaint = new IoTNatureofComplaint { StatusCode = "204", StatusDescription = "No Nature of Complaint is mapped with Serial Number." };
                            lstNatureOfComplaint.Add(objNatureOfComplaint);
                        }
                        else
                        {
                            foreach (Entity ent in entcoll.Entities)
                            {
                                objNatureOfComplaint = new IoTNatureofComplaint();
                                objNatureOfComplaint.Guid = ent.GetAttributeValue<Guid>("hil_natureofcomplaintid");
                                objNatureOfComplaint.Name = ent.GetAttributeValue<string>("hil_name");
                                objNatureOfComplaint.SerialNumber = natureOfComplaint.SerialNumber;
                                objNatureOfComplaint.StatusCode = "200";
                                objNatureOfComplaint.StatusDescription = "OK";
                                lstNatureOfComplaint.Add(objNatureOfComplaint);
                            }
                        }
                        return lstNatureOfComplaint;
                    }
                }
                else
                {
                    objNatureOfComplaint = new IoTNatureofComplaint { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    lstNatureOfComplaint.Add(objNatureOfComplaint);
                    return lstNatureOfComplaint;
                }
            }
            catch (Exception ex)
            {
                objNatureOfComplaint = new IoTNatureofComplaint { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                lstNatureOfComplaint.Add(objNatureOfComplaint);
                return lstNatureOfComplaint;
            }
        }

        public List<IoTNatureofComplaint> GetIoTNatureOfComplaintsByProdSubCatg(IoTNatureofComplaint natureOfComplaint)
        {
            IoTNatureofComplaint objNatureOfComplaint;
            List<IoTNatureofComplaint> lstNatureOfComplaint = new List<IoTNatureofComplaint>();
            EntityCollection entcoll;
            QueryExpression Query;

            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (natureOfComplaint.ProductSubCategoryId == Guid.Empty)
                    {
                        objNatureOfComplaint = new IoTNatureofComplaint { StatusCode = "204", StatusDescription = "Product Subcategory is required." };
                        lstNatureOfComplaint.Add(objNatureOfComplaint);
                        return lstNatureOfComplaint;
                    }
                    string fetchXml = string.Empty;
                    if (natureOfComplaint.Source == null)
                    {
                        fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                            <entity name='hil_natureofcomplaint'>
                            <attribute name='hil_name' />
                            <attribute name='hil_natureofcomplaintid' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_relatedproduct' operator='eq' value='{" + natureOfComplaint.ProductSubCategoryId + @"}' />
                            </filter>
                            </entity>
                            </fetch>";
                    }
                    else
                    {
                        fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                            <entity name='hil_natureofcomplaint'>
                            <attribute name='hil_name' />
                            <attribute name='hil_natureofcomplaintid' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                                <condition attribute='statecode' operator='eq' value='0' />
                                <condition attribute='hil_relatedproduct' operator='eq' value='{" + natureOfComplaint.ProductSubCategoryId + @"}' />
                                <condition attribute='hil_callsubtype' operator='in'>
                                <value uiname='AMC Call' uitype='hil_callsubtype'>{55A71A52-3C0B-E911-A94E-000D3AF06CD4}</value>
                                <value uiname='Breakdown' uitype='hil_callsubtype'>{6560565A-3C0B-E911-A94E-000D3AF06CD4}</value>
                                <value uiname='Demo' uitype='hil_callsubtype'>{AE1B2B71-3C0B-E911-A94E-000D3AF06CD4}</value>
                                <value uiname='Installation' uitype='hil_callsubtype'>{E3129D79-3C0B-E911-A94E-000D3AF06CD4}</value>
                                <value uiname='PMS' uitype='hil_callsubtype'>{E2129D79-3C0B-E911-A94E-000D3AF06CD4}</value>
                                </condition>
                            </filter>
                            </entity>
                            </fetch>";
                    }
                    entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (entcoll.Entities.Count == 0)
                    {
                        objNatureOfComplaint = new IoTNatureofComplaint { StatusCode = "204", StatusDescription = "No Nature of Complaint is mapped with Serial Number." };
                        lstNatureOfComplaint.Add(objNatureOfComplaint);
                    }
                    else
                    {
                        foreach (Entity ent in entcoll.Entities)
                        {
                            objNatureOfComplaint = new IoTNatureofComplaint();
                            objNatureOfComplaint.Guid = ent.GetAttributeValue<Guid>("hil_natureofcomplaintid");
                            objNatureOfComplaint.Name = ent.GetAttributeValue<string>("hil_name");
                            objNatureOfComplaint.SerialNumber = natureOfComplaint.SerialNumber;
                            objNatureOfComplaint.ProductSubCategoryId = natureOfComplaint.ProductSubCategoryId;
                            objNatureOfComplaint.StatusCode = "200";
                            objNatureOfComplaint.StatusDescription = "OK";
                            lstNatureOfComplaint.Add(objNatureOfComplaint);
                        }
                    }
                    return lstNatureOfComplaint;
                }
                else
                {
                    objNatureOfComplaint = new IoTNatureofComplaint { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    lstNatureOfComplaint.Add(objNatureOfComplaint);
                    return lstNatureOfComplaint;
                }
            }
            catch (Exception ex)
            {
                objNatureOfComplaint = new IoTNatureofComplaint { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                lstNatureOfComplaint.Add(objNatureOfComplaint);
                return lstNatureOfComplaint;
            }
        }

        public CancelJobResponse CancelServiceJob(CancelJobRequest reqParam)
        {
            CancelJobResponse res = new CancelJobResponse();
            try
            {
                if (reqParam.JobGuid == string.Empty || reqParam.JobGuid == null)
                {
                    res = new CancelJobResponse { StatusCode = "204", StatusDescription = "Job Id is required" };
                    return res;
                }
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    string _cancelStatusId = "1527FA6C-FA0F-E911-A94E-000D3AF060A1";

                    QueryExpression Query = new QueryExpression("msdyn_workorder");
                    Query.ColumnSet = new ColumnSet("msdyn_substatus");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("msdyn_workorderid", ConditionOperator.Equal, new Guid(reqParam.JobGuid));
                    EntityCollection entCols = service.RetrieveMultiple(Query);
                    if (entCols.Entities.Count > 0)
                    {
                        EntityReference _erSubstatus = entCols.Entities[0].GetAttributeValue<EntityReference>("msdyn_substatus");
                        if (_erSubstatus.Id.ToString().ToUpper() == _cancelStatusId)
                        {
                            res = new CancelJobResponse { StatusCode = "204", StatusDescription = "This Job is already Cancelled" };
                            return res;
                        }
                        string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_statustransitionmatrix'>
                        <attribute name='hil_statustransitionmatrixid' />
                        <filter type='and'>
                            <condition attribute='hil_basestatus' operator='eq' value='" + _erSubstatus.Id + @"' />
                            <filter type='or'>
                            <condition attribute='hil_statustransition1' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition2' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition3' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition4' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition5' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition6' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition7' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition8' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition9' operator='eq' value='{" + _cancelStatusId + @"}' />
                            <condition attribute='hil_statustransition10' operator='eq' value='{" + _cancelStatusId + @"}' />
                            </filter>
                        </filter>
                        </entity>
                        </fetch>";
                        entCols = service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entCols.Entities.Count == 0)
                        {
                            res = new CancelJobResponse { StatusCode = "204", StatusDescription = "This Job cannot be Cancelled, please contact with Customer Care." };
                            return res;
                        }
                    }
                    else
                    {
                        res = new CancelJobResponse { StatusCode = "204", StatusDescription = "Invalid Job ID." };
                        return res;
                    }

                    Query = new QueryExpression("msdyn_workordersubstatus");
                    Query.ColumnSet = new ColumnSet(false);
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("msdyn_name", ConditionOperator.Equal, "Canceled");
                    entCols = service.RetrieveMultiple(Query);
                    if (entCols.Entities.Count > 0)
                    {
                        Entity _jobUpdate = new Entity("msdyn_workorder");
                        _jobUpdate.Id = new Guid(reqParam.JobGuid);
                        _jobUpdate["hil_closureremarks"] = "Job cancelled by Customer via Sync App";
                        _jobUpdate["hil_jobcancelreason"] = new OptionSetValue(7);
                        _jobUpdate["msdyn_substatus"] = new EntityReference(entCols.EntityName, entCols[0].Id);

                        DateTime _today = DateTime.Now;

                        _today = _today.AddHours(5);
                        _today = _today.AddMinutes(30);

                        _jobUpdate["hil_jobclosureon"] = _today;
                        _jobUpdate["msdyn_timeclosed"] = _today;

                        try
                        {
                            service.Update(_jobUpdate);
                        }
                        catch (Exception ex)
                        {
                            res = new CancelJobResponse { StatusCode = "204", StatusDescription = "This Job cannot be Cancelled, please contact with Customer Care. \n" + ex.Message };
                            return res;
                        }
                        res.StatusCode = "200";
                        res.StatusDescription = "Job is Cancelled";
                    }
                    else
                    {
                        res = new CancelJobResponse { StatusCode = "204", StatusDescription = "Job Sub status(Cancel) is not defined in System." };
                        return res;
                    }
                }
                else
                {
                    res = new CancelJobResponse { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                }
            }
            catch (Exception ex)
            {
                res = new CancelJobResponse { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
            }
            return res;
        }

        public IoTServiceCallRegistration IoTCreateServiceCallDealerPortal(IoTServiceCallRegistration serviceCalldata)
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
                IOrganizationService service = ConnectToCRM.GetOrgService();
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
                        Entity ent = service.Retrieve("msdyn_customerasset", serviceCalldata.AssetGuid, new ColumnSet("msdyn_name", "hil_customer", "hil_productsubcategorymapping", "hil_productcategory", "hil_productsubcategory", "msdyn_customerassetid", "hil_invoicedate"));
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

                    string _dealerInfo = (serviceCalldata.DealerCode + "|" + serviceCalldata.DealerName);
                    enWorkorder["hil_newserialnumber"] = _dealerInfo.Length > 99 ? _dealerInfo.Substring(0, 99) : _dealerInfo; //Dealer Code/Name 

                    serviceCallGuid = service.Create(enWorkorder);
                    if (serviceCallGuid != Guid.Empty)
                    {
                        if (serviceCalldata.DealerCode != null && serviceCalldata.DealerCode != "" && serviceCalldata.DealerCode.Trim().Length > 0)
                        {
                            Query = new QueryExpression("hil_jobsextension");
                            Query.ColumnSet = new ColumnSet("hil_dealercode", "hil_dealername");
                            Query.Criteria = new FilterExpression(LogicalOperator.And);
                            Query.Criteria.AddCondition("hil_jobs", ConditionOperator.Equal, serviceCallGuid);
                            EntityCollection entColExt = service.RetrieveMultiple(Query);
                            if (entColExt.Entities.Count > 0)
                            {
                                Entity _entExt = entColExt.Entities[0];
                                _entExt["hil_dealercode"] = serviceCalldata.DealerCode;
                                _entExt["hil_dealername"] = serviceCalldata.DealerName;
                                try
                                {
                                    service.Update(_entExt);
                                }
                                catch (Exception ex)
                                {
                                    objServiceCall = new IoTServiceCallRegistration { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                                    return objServiceCall;
                                }
                            }
                        }
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

        public IoTServiceCallRegistration IoTCreateServiceVoiceBot(IoTServiceCallRegistration serviceCalldata)
        {
            IoTServiceCallRegistration objServiceCall;

            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (serviceCalldata.CustomerMobleNo == string.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer Nobile Number is required." };
                    }
                    if (serviceCalldata.CustomerName == null || serviceCalldata.CustomerName == string.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Customer Name is required." };
                    }
                    if (serviceCalldata.NOCName == string.Empty || serviceCalldata.NOCName == null || serviceCalldata.NOCName.Trim().Length == 0)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Nature of Complaint is required." };
                    }

                    if (serviceCalldata.Pincode == string.Empty || serviceCalldata.Pincode == null)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Pincode is required." };
                    }
                    if (serviceCalldata.AddressLine1 == string.Empty || serviceCalldata.AddressLine1 == null)
                    {
                        //return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Consumer Address is required." };
                        serviceCalldata.AddressLine1 = "No Address: " + serviceCalldata.Pincode;
                    }
                    if (serviceCalldata.ProductCategoryGuid == Guid.Empty)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Product Category is required to proceed." };
                    }
                    if (serviceCalldata.ChiefComplaint == string.Empty || serviceCalldata.ChiefComplaint == null)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Product Category is required to proceed." };
                    }
                    if (serviceCalldata.SourceOfJob != 18)
                    {
                        return new IoTServiceCallRegistration { StatusCode = "204", StatusDescription = "Invalid Source of Job!!! Please Input 18." };
                    }
                    Entity enServiceReq = new Entity("hil_servicerequestvoicebot");

                    enServiceReq["hil_addressline1"] = serviceCalldata.AddressLine1;
                    enServiceReq["hil_chiefcomplaint"] = serviceCalldata.ChiefComplaint;
                    enServiceReq["hil_customername"] = serviceCalldata.CustomerName;
                    enServiceReq["hil_name"] = serviceCalldata.CustomerMobleNo;
                    enServiceReq["hil_noc"] = serviceCalldata.NOCName;
                    enServiceReq["hil_pincode"] = serviceCalldata.Pincode;
                    enServiceReq["hil_preferredlanguage"] = Convert.ToInt32(serviceCalldata.PreferredLanguage);
                    enServiceReq["hil_productcategory"] = serviceCalldata.ProductCategoryGuid.ToString();
                    enServiceReq["hil_sourceofrequest"] = new OptionSetValue(serviceCalldata.SourceOfJob);

                    try
                    {
                        service.Create(enServiceReq);
                        objServiceCall = new IoTServiceCallRegistration { StatusCode = "200", StatusDescription = "OK" };
                    }
                    catch (Exception ex)
                    {
                        objServiceCall = new IoTServiceCallRegistration { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                        return objServiceCall;
                    }

                    return objServiceCall;
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

        public AuthenticateConsumer AuthenticateConsumerAMC(AuthenticateConsumer requestParam)
        {
            IOrganizationService _service = ConnectToCRM.GetOrgService();
            AuthenticateConsumer _retValue = new AuthenticateConsumer()
            {
                LoginUserId = requestParam.LoginUserId,
                SourceType = requestParam.SourceType
            };
            Entity _userSession = null;
            try
            {
                if (requestParam.SourceType == string.Empty && requestParam.SourceType == null)// != "5")
                {
                    _retValue.StatusCode = "400";
                    _retValue.StatusDescription = "Access Denied!!! Source Type is mandatory";
                    return _retValue;
                }
                if (requestParam.LoginUserId == string.Empty || requestParam.LoginUserId == null)
                {
                    _retValue.StatusCode = "400";
                    _retValue.StatusDescription = "Access Denied!!! Source Code or Mobile Number is mandatory";
                    return _retValue;
                }
                if (_service != null)
                {
                    int expTime = 0;
                    string _portalSURL = string.Empty;
                    if (!CheckSourceType(_service, requestParam.SourceType, out expTime, out _portalSURL))
                    {
                        _retValue.StatusCode = "400";
                        _retValue.StatusDescription = "Access Denied!!! API is not extended to Source Type: " + requestParam.SourceType;
                        return _retValue;
                    }
                    else
                    {
                        QueryExpression queryExp = new QueryExpression("hil_consumerloginsession");
                        queryExp.ColumnSet = new ColumnSet("hil_expiredon", "hil_consumerloginsessionid");
                        ConditionExpression condExp = new ConditionExpression("hil_name", ConditionOperator.Equal, requestParam.LoginUserId);
                        queryExp.Criteria.AddCondition(condExp);
                        condExp = new ConditionExpression("statecode", ConditionOperator.Equal, 0);
                        queryExp.Criteria.AddCondition(condExp);
                        queryExp.AddOrder("hil_expiredon", OrderType.Descending);

                        EntityCollection entCol = _service.RetrieveMultiple(queryExp);
                        if (entCol.Entities.Count > 0) //No Active session found
                        {
                            foreach (Entity entity in entCol.Entities)
                            {
                                SetStateRequest setStateRequest = new SetStateRequest()
                                {
                                    EntityMoniker = new EntityReference
                                    {
                                        Id = entCol.Entities[0].Id,
                                        LogicalName = "hil_consumerloginsession",
                                    },
                                    State = new OptionSetValue(1), //Inactive
                                    Status = new OptionSetValue(2) //Inactive
                                };
                                _service.Execute(setStateRequest);
                            }
                        }
                        _userSession = new Entity("hil_consumerloginsession");
                        _userSession["hil_name"] = requestParam.LoginUserId;
                        _userSession["hil_origin"] = requestParam.SourceType;
                        expTime = expTime + 330;
                        _userSession["hil_expiredon"] = DateTime.Now.AddMinutes(expTime);
                        _retValue.SessionId = _service.Create(_userSession).ToString();

                        string _postfixURL = EncryptAES256("SessionId=" + _retValue.SessionId + "&MobileNumber=" + requestParam.LoginUserId + "&SourceOrigin=" + requestParam.SourceType);
                        _retValue.RedirectUrl = _portalSURL + _postfixURL;
                        _retValue.StatusDescription = "Success";
                        _retValue.StatusCode = "200";
                    }
                }
                else
                {
                    _retValue.StatusCode = "503";
                    _retValue.StatusDescription = "D365 Service is unavailable.";
                }
            }
            catch (Exception ex)
            {
                _retValue.StatusCode = "500";
                _retValue.StatusDescription = "D365 Internal Server Error : " + ex.Message;
            }
            return _retValue;
        }

        public ValidateSessionResponse ValidateSessionDetails(ValidateSessionRequest requestParam)
        {
            ValidateSessionResponse _retValue = new ValidateSessionResponse();
            try
            {
                IOrganizationService _service = ConnectToCRM.GetOrgService();

                if (requestParam.SessionId == string.Empty && requestParam.SessionId == null)
                {
                    _retValue.StatusCode = "400";
                    _retValue.StatusDescription = "Access Denied!!! Please input Session Id.";
                    return _retValue;
                }
                QueryExpression queryExp = new QueryExpression("hil_consumerloginsession");
                queryExp.ColumnSet = new ColumnSet("hil_expiredon", "hil_consumerloginsessionid");
                ConditionExpression condExp = new ConditionExpression("hil_consumerloginsessionid", ConditionOperator.Equal, requestParam.SessionId);
                queryExp.Criteria.AddCondition(condExp);
                condExp = new ConditionExpression("statecode", ConditionOperator.Equal, 0);
                queryExp.Criteria.AddCondition(condExp);

                EntityCollection entCol = _service.RetrieveMultiple(queryExp);

                if (entCol.Entities.Count == 1)
                {
                    DateTime expDate = entCol[0].GetAttributeValue<DateTime>("hil_expiredon");
                    if (expDate > DateTime.Now)
                    {
                        if (requestParam.KeepSessionLive)
                        {
                            Entity _userSession = new Entity("hil_consumerloginsession", entCol.Entities[0].Id);
                            _userSession["hil_expiredon"] = DateTime.Now.AddMinutes(350);
                            _service.Update(_userSession);
                        }
                        _retValue.StatusCode = "200";
                        _retValue.StatusDescription = "Session Id is Valid";
                    }
                    else
                    {
                        _retValue.StatusCode = "400";
                        _retValue.StatusDescription = "Access Denied!!! Session has been expired";
                        SetStateRequest setStateRequest = new SetStateRequest()
                        {
                            EntityMoniker = new EntityReference
                            {
                                Id = entCol.Entities[0].Id,
                                LogicalName = "hil_consumerloginsession",
                            },
                            State = new OptionSetValue(1), //Inactive
                            Status = new OptionSetValue(2) //Inactive
                        };
                        _service.Execute(setStateRequest);
                    }
                }
                else
                {
                    _retValue.StatusCode = "400";
                    _retValue.StatusDescription = "Invalid Session Id.";
                }
            }
            catch (Exception ex)
            {
                _retValue.StatusCode = "400";
                _retValue.StatusDescription = "ERROR!!! " + ex.Message;
            }
            return _retValue;
        }

        public bool CheckSourceType(IOrganizationService service, string sourceOrigin, out int expTime, out string portalURL)
        {
            QueryExpression qryExp = new QueryExpression("hil_integrationsource");
            qryExp.ColumnSet = new ColumnSet("hil_deeplinkingallowed", "hil_sessiontimeout", "hil_amcportalurl");
            qryExp.Criteria.AddCondition(new ConditionExpression("hil_code", ConditionOperator.Equal, sourceOrigin));
            qryExp.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            EntityCollection entCol = service.RetrieveMultiple(qryExp);
            if (entCol.Entities.Count != 1)
            {
                expTime = 0;
                portalURL = string.Empty;
                return false;
            }
            else
            {
                expTime = entCol[0].GetAttributeValue<int>("hil_sessiontimeout");
                portalURL = entCol[0].GetAttributeValue<string>("hil_amcportalurl");
                return entCol[0].GetAttributeValue<bool>("hil_deeplinkingallowed");
            }
        }
        private string EncryptAES256(string plainText)
        {
            string Key = "DklsdvkfsDlkslsdsdnv234djSDAjkd1";
            byte[] key32 = Encoding.UTF8.GetBytes(Key);
            byte[] IV16 = Encoding.UTF8.GetBytes(Key.Substring(0, 16)); if (plainText == null || plainText.Length <= 0)
            {
                throw new ArgumentNullException("plainText");
            }
            byte[] encrypted;
            using (AesManaged aesAlg = new AesManaged())
            {
                aesAlg.KeySize = 256;
                aesAlg.Padding = PaddingMode.PKCS7;
                aesAlg.IV = IV16;
                aesAlg.Key = key32;
                ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
                using (MemoryStream msEncrypt = new MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                        {
                            //Write all data to the stream.
                            swEncrypt.Write(plainText);
                        }
                        encrypted = msEncrypt.ToArray();
                    }
                }
            }
            return Convert.ToBase64String(encrypted);
        }

        public ValidateProductInstallation ValidateProductInstallation(ValidateProductInstallation _data)
        {
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (string.IsNullOrEmpty(_data.productCode) || string.IsNullOrWhiteSpace(_data.productCode))
                    {
                        return new ValidateProductInstallation() { StatusCode = "204", StatusDescription = "SKU Code is required." };
                    }

                    QueryExpression Query = new QueryExpression("product");
                    Query.ColumnSet = new ColumnSet("hil_installationrequired");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("productnumber", ConditionOperator.Equal, _data.productCode);
                    EntityCollection entcoll = service.RetrieveMultiple(Query);
                    if (entcoll.Entities.Count > 0)
                    {
                        return new ValidateProductInstallation { StatusCode = "200", StatusDescription = "OK", productCode = _data.productCode, isInstallable = entcoll.Entities[0].Attributes.Contains("hil_installationrequired") ? entcoll.Entities[0].GetAttributeValue<bool>("hil_installationrequired") : false };
                    }
                    else
                    {
                        return new ValidateProductInstallation { StatusCode = "204", StatusDescription = "SKU/Model Code does not exist.", productCode = _data.productCode };
                    }
                }
                else
                {
                    return new ValidateProductInstallation { StatusCode = "503", StatusDescription = "D365 Service Unavailable", productCode = _data.productCode };
                }
            }
            catch (Exception ex)
            {
                return new ValidateProductInstallation { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper(), productCode = _data.productCode };
            }
        }
    }

    [DataContract]
    //public class IoTServiceCallRegistration
    //{
    //    [DataMember]
    //    public string SerialNumber { get; set; }
    //    [DataMember]
    //    public string ProductModelNumber { get; set; }
    //    [DataMember]
    //    public Guid NOCGuid { get; set; }
    //    [DataMember]
    //    public string NOCName { get; set; }
    //    [DataMember]
    //    public Guid ProductCategoryGuid { get; set; }

    //    [DataMember]
    //    public Guid ProductSubCategoryGuid { get; set; }

    //    [DataMember]
    //    public string ChiefComplaint { get; set; }
    //    [DataMember]
    //    public Guid AddressGuid { get; set; }
    //    [DataMember]
    //    public Guid AssetGuid { get; set; }
    //    [DataMember]
    //    public string CustomerMobleNo { get; set; }
    //    [DataMember]
    //    public Guid CustomerGuid { get; set; }
    //    [DataMember]
    //    public Guid JobGuid { get; set; }
    //    [DataMember]
    //    public string JobId { get; set; }
    //    [DataMember]
    //    public string ImageBase64String { get; set; }
    //    [DataMember]
    //    public int ImageType { get; set; }
    //    [DataMember]
    //    public int SourceOfJob { get; set; }
    //    [DataMember]
    //    public string PreferredDate { get; set; }
    //    [DataMember]
    //    public int PreferredPartOfDay { get; set; }
    //    [DataMember]
    //    public string StatusCode { get; set; }
    //    [DataMember]
    //    public string StatusDescription { get; set; }
    //    [DataMember]
    //    public string DealerCode { get; set; }
    //    [DataMember]
    //    public string DealerName { get; set; }
    //    [DataMember]
    //    public string CustomerName { get; set; }
    //    [DataMember]
    //    public string AddressLine1 { get; set; }
    //    [DataMember]
    //    public string Pincode { get; set; }
    //    [DataMember]
    //    public string PreferredLanguage { get; set; }
    //}

    public class IoTServiceCallRegistration
    {
        [DataMember]
        public string SerialNumber { get; set; }
        [DataMember]
        public string ProductModelNumber { get; set; }
        [DataMember]
        public Guid NOCGuid { get; set; }
        [DataMember]
        public string NOCName { get; set; }
        [DataMember]
        public Guid ProductCategoryGuid { get; set; }

        [DataMember]
        public Guid ProductSubCategoryGuid { get; set; }

        [DataMember]
        public string ChiefComplaint { get; set; }
        [DataMember]
        public Guid AddressGuid { get; set; }
        [DataMember]
        public Guid AssetGuid { get; set; }
        [DataMember]
        public string CustomerMobleNo { get; set; }
        [DataMember]
        public Guid CustomerGuid { get; set; }
        [DataMember]
        public Guid JobGuid { get; set; }
        [DataMember]
        public string JobId { get; set; }
        [DataMember]
        public string ImageBase64String { get; set; }
        [DataMember]
        public int ImageType { get; set; }
        [DataMember]
        public int SourceOfJob { get; set; }
        [DataMember]
        public string PreferredDate { get; set; }
        [DataMember]
        public int PreferredPartOfDay { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
        [DataMember]
        public string DealerCode { get; set; }
        [DataMember]
        public string DealerName { get; set; }
        [DataMember]
        public string CustomerName { get; set; }
        [DataMember]
        public string AddressLine1 { get; set; }
        [DataMember]
        public string Pincode { get; set; }
        [DataMember]
        public string PreferredLanguage { get; set; }
        [DataMember]
        public string CallSubType { get; set; }

    }
    //[DataContract]
    //public class IoTServiceCallRegistration
    //{
    //    [DataMember]
    //    public string SerialNumber { get; set; }
    //    [DataMember]
    //    public string ProductModelNumber { get; set; }
    //    [DataMember]
    //    public Guid NOCGuid { get; set; }
    //    [DataMember]
    //    public string NOCName { get; set; }
    //    [DataMember]
    //    public Guid ProductCategoryGuid { get; set; }

    //    [DataMember]
    //    public Guid ProductSubCategoryGuid { get; set; }

    //    [DataMember]
    //    public string ChiefComplaint { get; set; }
    //    [DataMember]
    //    public Guid AddressGuid { get; set; }
    //    [DataMember]
    //    public Guid AssetGuid { get; set; }
    //    [DataMember]
    //    public string CustomerMobleNo { get; set; }
    //    [DataMember]
    //    public Guid CustomerGuid { get; set; }
    //    [DataMember]
    //    public Guid JobGuid { get; set; }
    //    [DataMember]
    //    public string JobId { get; set; }
    //    [DataMember]
    //    public string ImageBase64String { get; set; }
    //    [DataMember]
    //    public int ImageType { get; set; }
    //    [DataMember]
    //    public int SourceOfJob { get; set; }
    //    [DataMember]
    //    public string PreferredDate { get; set; }
    //    [DataMember]
    //    public int PreferredPartOfDay { get; set; }
    //    [DataMember]
    //    public string StatusCode { get; set; }
    //    [DataMember]
    //    public string StatusDescription { get; set; }
    //    [DataMember]
    //    public string DealerCode { get; set; }
    //    [DataMember]
    //    public string DealerName { get; set; }
    //}

    [DataContract]
    public class IoTServiceCallResult
    {
        [DataMember]
        public string JobId { get; set; }

        [DataMember]
        public Guid JobGuid { get; set; }

        [DataMember]
        public string CallSubType { get; set; }

        [DataMember]
        public string JobLoggedon { get; set; }

        [DataMember]
        public string JobStatus { get; set; }

        [DataMember]
        public string JobAssignedTo { get; set; }

        [DataMember]
        public string CustomerAsset { get; set; }

        [DataMember]
        public string ProductCategory { get; set; }

        [DataMember]
        public string NatureOfComplaint { get; set; }

        [DataMember]
        public string JobClosedOn { get; set; }

        [DataMember]
        public string CustomerName { get; set; }

        [DataMember]
        public string ServiceAddress { get; set; }

        [DataMember]
        public string Product { get; set; }

        [DataMember]
        public string ChiefComplaint { get; set; }
        [DataMember]
        public string PreferredDate { get; set; }
        [DataMember]
        public int PreferredPartOfDay { get; set; }
        [DataMember]
        public string PreferredPartOfDayName { get; set; }
        [DataMember]
        public string StatusCode { get; set; }

        [DataMember]
        public string StatusDescription { get; set; }
    }

    [DataContract]
    public class IoTNatureofComplaint
    {
        [DataMember]
        public string SerialNumber { get; set; }
        [DataMember]
        public Guid ProductSubCategoryId { get; set; }

        [DataMember]
        public string Name { get; set; }
        [DataMember]
        public Guid Guid { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
        [DataMember]
        public string Source { get; set; }
    }

    [DataContract]
    public class CancelJobRequest
    {
        [DataMember]
        public string JobGuid { get; set; }
    }
    [DataContract]
    public class CancelJobResponse
    {
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
    }
    public class AuthenticateConsumer
    {
        [DataMember]
        public string LoginUserId { get; set; }
        [DataMember]
        public string SourceType { get; set; }

        [DataMember]
        public string SessionId { get; set; }
        [DataMember]
        public string RedirectUrl { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
    }
    [DataContract]
    public class ValidateSessionRequest
    {
        [DataMember]
        public string JWTToken { get; set; }
        [DataMember]
        public string MobileNumber { get; set; }
        [DataMember]
        public string SourceType { get; set; }
        [DataMember]
        public string SourceCode { get; set; }
        [DataMember]
        public string SessionId { get; set; }
        [DataMember]
        public bool KeepSessionLive { get; set; }
    }
    [DataContract]
    public class ValidateSessionResponse
    {
        [DataMember]
        public string JWTToken { get; set; }
        [DataMember]
        public string MobileNumber { get; set; }
        [DataMember]
        public string SourceType { get; set; }
        [DataMember]
        public string SourceCode { get; set; }
        [DataMember]
        public string SessionId { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
    }
    [DataContract]
    public class ValidateProductInstallation
    {
        [DataMember]
        public string productCode { get; set; }
        [DataMember]
        public bool isInstallable { get; set; }
        [DataMember]
        public string StatusCode { get; set; }
        [DataMember]
        public string StatusDescription { get; set; }
    }
}
