using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using ConsumerApp.BusinessLayer;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class CreateEnquiry
    {

        [DataMember]
        public string EnquiryType { get; set; }

        [DataMember]
        public string productDivisionID { get; set; }

        [DataMember]
        public string Name { get; set; }

        [DataMember]
        public string MobNumber { get; set; }

        [DataMember]
        public string Email { get; set; }

        [DataMember]
        public string Pincode { get; set; }

        [DataMember]
        public string Remarks { get; set; }

        [DataMember]
        public string UTMCode { get; set; }

        [DataMember]
        public string CampaignCode { get; set; }

        //public ReturnInfo CreateEnquiryEntry(CreateEnquiry obj_Enquiry)
        //{
        //    IOrganizationService service = ConnectToCRM.GetOrgService(); //get org service obj for connection
        //    ReturnInfo obj_return = new ReturnInfo();
        //    Guid _enquiryId = Guid.Empty;
        //    Guid _productId = Guid.Empty;
        //    string _stringProductId = string.Empty;
        //    string _selectedDivision = "";
        //    string _selectedDivisionName = "";
        //    string _dataPacket;
        //    Guid campaignGuId = Guid.Empty;

        //    try
        //    {
        //        if (obj_Enquiry.EnquiryType == null || obj_Enquiry.EnquiryType.Trim().Length == 0)
        //        {
        //            obj_return.ErrorCode = "FAILURE";
        //            obj_return.ErrorDescription = "Enquiry Type is required.";
        //        }
        //        //else if (obj_Enquiry.EnquiryType != "910590000" && obj_Enquiry.EnquiryType != "910590001" && obj_Enquiry.EnquiryType != "910590002" && obj_Enquiry.EnquiryType != "910590003")
        //        //{
        //        //    obj_return.ErrorCode = "FAILURE";
        //        //    obj_return.ErrorDescription = "Invalid Enquiry Type.";
        //        //}
        //        else if (obj_Enquiry.productDivisionID == null || obj_Enquiry.productDivisionID.Trim().Length == 0)
        //        {
        //            obj_return.ErrorCode = "FAILURE";
        //            obj_return.ErrorDescription = "Product Division is required.";
        //        }
        //        else if (obj_Enquiry.Name == null || obj_Enquiry.Name.Trim().Length == 0)
        //        {
        //            obj_return.ErrorCode = "FAILURE";
        //            obj_return.ErrorDescription = "Customer Name is required.";
        //        }
        //        else if (obj_Enquiry.MobNumber == null || obj_Enquiry.MobNumber.Trim().Length == 0)
        //        {
        //            obj_return.ErrorCode = "FAILURE";
        //            obj_return.ErrorDescription = "Mobile Number is required.";
        //        }
        //        else if (obj_Enquiry.Email == null)
        //        {
        //            obj_return.ErrorCode = "FAILURE";
        //            obj_return.ErrorDescription = "Email is required.";
        //        }
        //        else if (obj_Enquiry.Pincode == null || obj_Enquiry.Pincode.Trim().Length == 0)
        //        {
        //            obj_return.ErrorCode = "FAILURE";
        //            obj_return.ErrorDescription = "Pincode is PIN Code.";
        //        }
        //        else
        //        {
        //            //Preparing Request Data Packet
        //            _dataPacket = JsonConvert.SerializeObject(obj_Enquiry);

        //            Entity obj_createEnquiryRecord = new Entity("lead");

        //            //if (obj_Enquiry.EnquiryType == "910590000" || obj_Enquiry.EnquiryType == "910590001" || obj_Enquiry.EnquiryType == "910590002" || obj_Enquiry.EnquiryType == "910590003")
        //            //{
        //            obj_createEnquiryRecord.Attributes["hil_leadtype"] = new OptionSetValue(Convert.ToInt32(obj_Enquiry.EnquiryType));
        //            //}

        //            string[] productDivisionID;

        //            if (obj_Enquiry.productDivisionID.IndexOf(',') < 0)
        //            {
        //                obj_Enquiry.productDivisionID = obj_Enquiry.productDivisionID + ",";
        //            }
        //            productDivisionID = obj_Enquiry.productDivisionID.Split(',');
        //            //Product Selection
        //            obj_createEnquiryRecord.Attributes["hil_productenquirynumber"] = new OptionSetValue(productDivisionID.Length > 1 ? 910590001 : 910590000);

        //            for (int i = 0; i < productDivisionID.Length; i++)
        //            {
        //                if (productDivisionID[i].ToString().Trim().Length > 0)
        //                {
        //                    QueryExpression qsProduct = new QueryExpression("product");
        //                    qsProduct.ColumnSet = new ColumnSet("hil_sapcode", "name");
        //                    ConditionExpression cond8 = new ConditionExpression("hil_sapcode", ConditionOperator.Equal, productDivisionID[i]);
        //                    qsProduct.Criteria.AddCondition(cond8);
        //                    ConditionExpression cond80 = new ConditionExpression("hil_hierarchylevel", ConditionOperator.Equal, "2"); //{"Hierarchy Level" : Division} 
        //                    qsProduct.Criteria.AddCondition(cond80);
        //                    EntityCollection collct_product = service.RetrieveMultiple(qsProduct);
        //                    if (collct_product.Entities.Count > 0)
        //                    {
        //                        foreach (Entity product_record in collct_product.Entities)
        //                        {
        //                            _productId = product_record.Id;
        //                            _stringProductId = Convert.ToString(_productId);
        //                            _selectedDivisionName += product_record.Attributes["name"].ToString() + ",";
        //                        }
        //                    }
        //                    _selectedDivision += _stringProductId + ",";
        //                }
        //            }
        //            if (_selectedDivision.EndsWith(","))
        //            {
        //                _selectedDivision = _selectedDivision.Substring(0, _selectedDivision.Length - 1);
        //                _selectedDivisionName = _selectedDivisionName.Substring(0, _selectedDivisionName.Length - 1);
        //            }

        //            obj_createEnquiryRecord.Attributes["hil_selecteddivisionsname"] = _selectedDivisionName;
        //            obj_createEnquiryRecord.Attributes["hil_selecteddivisions"] = _selectedDivision;
        //            obj_createEnquiryRecord.Attributes["firstname"] = obj_Enquiry.Name;
        //            obj_createEnquiryRecord.Attributes["mobilephone"] = obj_Enquiry.MobNumber;
        //            obj_createEnquiryRecord.Attributes["emailaddress1"] = obj_Enquiry.Email;
        //            obj_createEnquiryRecord.Attributes["leadsourcecode"] = new OptionSetValue(8);
        //            obj_createEnquiryRecord.Attributes["hil_jsondatapacket"] = _dataPacket;
        //            obj_createEnquiryRecord.Attributes["description"] = obj_Enquiry.Remarks;
        //            obj_createEnquiryRecord.Attributes["hil_utmcode"] = obj_Enquiry.UTMCode;

        //            if (obj_Enquiry.CampaignCode != null)
        //            {
        //                QueryExpression qsCampaign = new QueryExpression("campaign");
        //                qsCampaign.ColumnSet = new ColumnSet("codename");
        //                ConditionExpression condCampaign = new ConditionExpression("codename", ConditionOperator.Equal, obj_Enquiry.CampaignCode);
        //                qsCampaign.Criteria.AddCondition(condCampaign);
        //                EntityCollection collct_Campaign = service.RetrieveMultiple(qsCampaign);
        //                if (collct_Campaign.Entities.Count == 1)
        //                {
        //                    foreach (Entity Campaign_record in collct_Campaign.Entities)
        //                    {
        //                        campaignGuId = Campaign_record.Id;
        //                    }
        //                }
        //                if (campaignGuId != Guid.Empty)
        //                {
        //                    obj_createEnquiryRecord.Attributes["campaignid"] = new EntityReference("campaign", campaignGuId);
        //                }
        //            }

        //            QueryExpression qsPincode = new QueryExpression("hil_pincode");
        //            qsPincode.ColumnSet = new ColumnSet("hil_name");
        //            ConditionExpression condPincode = new ConditionExpression("hil_name", ConditionOperator.Equal, obj_Enquiry.Pincode);
        //            qsPincode.Criteria.AddCondition(condPincode);
        //            EntityCollection collct_pincode = service.RetrieveMultiple(qsPincode);
        //            if (collct_pincode.Entities.Count == 1)
        //            {
        //                foreach (Entity pincode_record in collct_pincode.Entities)
        //                {
        //                    Guid _pincodeId = pincode_record.Id;
        //                    obj_createEnquiryRecord.Attributes["hil_pincode"] = new EntityReference("hil_pincode", _pincodeId);
        //                }
        //            }

        //            Guid _primaryContact = Guid.Empty;
        //            QueryExpression qsprimaryContact = new QueryExpression("contact");
        //            qsprimaryContact.ColumnSet = new ColumnSet("firstname");
        //            ConditionExpression cond1 = new ConditionExpression("mobilephone", ConditionOperator.Equal, obj_Enquiry.MobNumber);
        //            qsprimaryContact.Criteria.AddCondition(cond1);
        //            EntityCollection collct_primaryContact = service.RetrieveMultiple(qsprimaryContact);
        //            if (collct_primaryContact.Entities.Count > 0)
        //            {
        //                foreach (Entity primaryContact_record in collct_primaryContact.Entities)
        //                {
        //                    _primaryContact = primaryContact_record.Id;
        //                }
        //            }
        //            else
        //            {
        //                Entity obj_contact = new Entity("contact");
        //                obj_contact.Attributes["firstname"] = obj_Enquiry.Name;
        //                obj_contact.Attributes["mobilephone"] = obj_Enquiry.MobNumber;
        //                _primaryContact = service.Create(obj_contact);
        //            }
        //            obj_createEnquiryRecord.Attributes["parentcontactid"] = new EntityReference("contact", _primaryContact);

        //            string _enquiryNo = string.Empty;

        //            _enquiryId = service.Create(obj_createEnquiryRecord);

        //            if (_enquiryId != Guid.Empty)
        //            {
        //                Entity obj_leadProduct = new Entity("hil_leadproduct");
        //                string[] divisions;

        //                if (_selectedDivision.IndexOf(',') < 0)
        //                {
        //                    _selectedDivision = _selectedDivision + ",";
        //                }

        //                divisions = _selectedDivision.Split(',');

        //                for (int i = 0; i < divisions.Length; i++)
        //                {
        //                    if (divisions[i] != null && divisions[i].ToString().Trim().Length > 0)
        //                    {
        //                        obj_leadProduct.Attributes["hil_enquiry"] = new EntityReference("lead", _enquiryId);
        //                        obj_leadProduct.Attributes["hil_productdivision"] = new EntityReference("product", Guid.Parse(divisions[i]));
        //                        service.Create(obj_leadProduct);
        //                    }
        //                }
        //                Entity obj_lead1 = (Entity)service.Retrieve("lead", _enquiryId, new ColumnSet("hil_leadstatus", "hil_ticketnumber"));
        //                _enquiryNo = obj_lead1.GetAttributeValue<string>("hil_ticketnumber");

        //                obj_lead1.Attributes["hil_leadstatus"] = new OptionSetValue(2);
        //                service.Update(obj_lead1);
        //            }
        //            obj_return.ErrorCode = "SUCCESS";
        //            obj_return.ErrorDescription = _enquiryNo;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        obj_return.ErrorCode = "FAILURE";
        //        obj_return.ErrorDescription = "ENQUIRY : " + ex.Message.ToUpper();
        //    }
        //    return obj_return;
        //}

        public ReturnInfoVar CreateEnquiryEntry(CreateEnquiry obj_Enquiry)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService(); //get org service obj for connection
            ReturnInfoVar obj_return = new ReturnInfoVar();
            Guid _enquiryId = Guid.Empty;
            Guid _productId = Guid.Empty;
            string _stringProductId = string.Empty;
            string _selectedDivision = "";
            string _selectedDivisionName = "";
            string _dataPacket;
            Guid campaignGuId = Guid.Empty;

            try
            {
                if (obj_Enquiry.EnquiryType == null || obj_Enquiry.EnquiryType.Trim().Length == 0)
                {
                    obj_return.ErrorCode = "FAILURE";
                    obj_return.ErrorDescription = "Enquiry Type is required.";
                }
                else if (obj_Enquiry.productDivisionID == null || obj_Enquiry.productDivisionID.Trim().Length == 0)
                {
                    obj_return.ErrorCode = "FAILURE";
                    obj_return.ErrorDescription = "Product Division is required.";
                }
                else if (obj_Enquiry.Name == null || obj_Enquiry.Name.Trim().Length == 0)
                {
                    obj_return.ErrorCode = "FAILURE";
                    obj_return.ErrorDescription = "Customer Name is required.";
                }
                else if (obj_Enquiry.MobNumber == null || obj_Enquiry.MobNumber.Trim().Length == 0)
                {
                    obj_return.ErrorCode = "FAILURE";
                    obj_return.ErrorDescription = "Mobile Number is required.";
                }
                else if (obj_Enquiry.Email == null)
                {
                    obj_return.ErrorCode = "FAILURE";
                    obj_return.ErrorDescription = "Email is required.";
                }
                else if (obj_Enquiry.Pincode == null || obj_Enquiry.Pincode.Trim().Length == 0)
                {
                    obj_return.ErrorCode = "FAILURE";
                    obj_return.ErrorDescription = "Pincode is PIN Code.";
                }
                else
                {
                    Guid _pincodeId = Guid.Empty;

                    QueryExpression qsPincode = new QueryExpression("hil_pincode");
                    qsPincode.ColumnSet = new ColumnSet("hil_name");
                    ConditionExpression condPincode = new ConditionExpression("hil_name", ConditionOperator.Equal, obj_Enquiry.Pincode);
                    qsPincode.Criteria.AddCondition(condPincode);
                    EntityCollection collct_pincode = service.RetrieveMultiple(qsPincode);
                    if (collct_pincode.Entities.Count == 0)
                    {
                        obj_return.ErrorCode = "FAILURE";
                        obj_return.ErrorDescription = "Invalid Pincode.";
                    }
                    else
                    {
                        _pincodeId = collct_pincode.Entities[0].Id;

                    }

                    //Preparing Request Data Packet
                    _dataPacket = JsonConvert.SerializeObject(obj_Enquiry);
                    Entity obj_createEnquiryRecord = new Entity("lead");
                    //if (obj_Enquiry.EnquiryType == "910590000" || obj_Enquiry.EnquiryType == "910590001" || obj_Enquiry.EnquiryType == "910590002" || obj_Enquiry.EnquiryType == "910590003")
                    //{
                    obj_createEnquiryRecord.Attributes["hil_leadtype"] = new OptionSetValue(Convert.ToInt32(obj_Enquiry.EnquiryType));
                    //}
                    string[] productDivisionID;
                    if (obj_Enquiry.productDivisionID.IndexOf(',') < 0)
                    {
                        obj_Enquiry.productDivisionID = obj_Enquiry.productDivisionID + ",";
                    }
                    productDivisionID = obj_Enquiry.productDivisionID.Split(',');
                    productDivisionID = productDivisionID.Distinct().ToArray();

                    //Product Selection
                    obj_createEnquiryRecord.Attributes["hil_productenquirynumber"] = new OptionSetValue(productDivisionID.Length > 1 ? 910590001 : 910590000);

                    for (int i = 0; i < productDivisionID.Length; i++)
                    {
                        if (productDivisionID[i].ToString().Trim().Length > 0)
                        {
                            QueryExpression qsProduct = new QueryExpression("product");
                            qsProduct.ColumnSet = new ColumnSet("hil_sapcode", "name");
                            ConditionExpression cond8 = new ConditionExpression("hil_sapcode", ConditionOperator.Equal, productDivisionID[i]);
                            qsProduct.Criteria.AddCondition(cond8);
                            cond8 = new ConditionExpression("hil_hierarchylevel", ConditionOperator.Equal, 2);
                            qsProduct.Criteria.AddCondition(cond8);
                            cond8 = new ConditionExpression("hil_claimcategory", ConditionOperator.Null);
                            qsProduct.Criteria.AddCondition(cond8);
                            EntityCollection collct_product = service.RetrieveMultiple(qsProduct);
                            if (collct_product.Entities.Count > 0)
                            {
                                foreach (Entity product_record in collct_product.Entities)
                                {
                                    _productId = product_record.Id;
                                    _stringProductId = Convert.ToString(_productId);
                                    _selectedDivisionName += product_record.Attributes["name"].ToString() + ",";
                                }
                            }
                            _selectedDivision += _stringProductId + ",";
                        }
                    }
                    if (_selectedDivision.EndsWith(","))
                    {
                        _selectedDivision = _selectedDivision.Substring(0, _selectedDivision.Length - 1);
                        _selectedDivisionName = _selectedDivisionName.Substring(0, _selectedDivisionName.Length - 1);
                    }

                    obj_createEnquiryRecord.Attributes["hil_selecteddivisionsname"] = _selectedDivisionName;
                    obj_createEnquiryRecord.Attributes["hil_selecteddivisions"] = _selectedDivision;
                    obj_createEnquiryRecord.Attributes["firstname"] = obj_Enquiry.Name;
                    obj_createEnquiryRecord.Attributes["mobilephone"] = obj_Enquiry.MobNumber;
                    obj_createEnquiryRecord.Attributes["emailaddress1"] = obj_Enquiry.Email;
                    obj_createEnquiryRecord.Attributes["leadsourcecode"] = new OptionSetValue(8);
                    obj_createEnquiryRecord.Attributes["hil_jsondatapacket"] = _dataPacket;
                    obj_createEnquiryRecord.Attributes["description"] = obj_Enquiry.Remarks;
                    obj_createEnquiryRecord.Attributes["hil_utmcode"] = obj_Enquiry.UTMCode;

                    if (obj_Enquiry.CampaignCode != null)
                    {
                        QueryExpression qsCampaign = new QueryExpression("campaign");
                        qsCampaign.ColumnSet = new ColumnSet("codename");
                        ConditionExpression condCampaign = new ConditionExpression("codename", ConditionOperator.Equal, obj_Enquiry.CampaignCode);
                        qsCampaign.Criteria.AddCondition(condCampaign);
                        EntityCollection collct_Campaign = service.RetrieveMultiple(qsCampaign);
                        if (collct_Campaign.Entities.Count == 1)
                        {
                            foreach (Entity Campaign_record in collct_Campaign.Entities)
                            {
                                campaignGuId = Campaign_record.Id;
                            }
                        }
                        if (campaignGuId != Guid.Empty)
                        {
                            obj_createEnquiryRecord.Attributes["campaignid"] = new EntityReference("campaign", campaignGuId);
                        }
                    }

                    if (_pincodeId != Guid.Empty)
                    {
                        obj_createEnquiryRecord.Attributes["hil_pincode"] = new EntityReference("hil_pincode", _pincodeId);
                    }

                    Guid _primaryContact = Guid.Empty;
                    QueryExpression qsprimaryContact = new QueryExpression("contact");
                    qsprimaryContact.ColumnSet = new ColumnSet("firstname");
                    ConditionExpression cond1 = new ConditionExpression("mobilephone", ConditionOperator.Equal, obj_Enquiry.MobNumber);
                    qsprimaryContact.Criteria.AddCondition(cond1);
                    EntityCollection collct_primaryContact = service.RetrieveMultiple(qsprimaryContact);
                    if (collct_primaryContact.Entities.Count > 0)
                    {
                        foreach (Entity primaryContact_record in collct_primaryContact.Entities)
                        {
                            _primaryContact = primaryContact_record.Id;
                        }
                    }
                    else
                    {
                        Entity obj_contact = new Entity("contact");
                        obj_contact.Attributes["firstname"] = obj_Enquiry.Name;
                        obj_contact.Attributes["mobilephone"] = obj_Enquiry.MobNumber;
                        _primaryContact = service.Create(obj_contact);
                    }

                    obj_createEnquiryRecord.Attributes["parentcontactid"] = new EntityReference("contact", _primaryContact);

                    string _enquiryNo = string.Empty;

                    _enquiryId = service.Create(obj_createEnquiryRecord);

                    if (_enquiryId != Guid.Empty)
                    {
                        Entity obj_leadProduct = new Entity("hil_leadproduct");
                        string[] divisions;

                        if (_selectedDivision.IndexOf(',') < 0)
                        {
                            _selectedDivision = _selectedDivision + ",";
                        }

                        divisions = _selectedDivision.Split(',');

                        for (int i = 0; i < divisions.Length; i++)
                        {
                            if (divisions[i] != null && divisions[i].ToString().Trim().Length > 0)
                            {
                                obj_leadProduct.Attributes["hil_enquiry"] = new EntityReference("lead", _enquiryId);
                                obj_leadProduct.Attributes["hil_productdivision"] = new EntityReference("product", Guid.Parse(divisions[i]));
                                service.Create(obj_leadProduct);
                            }
                        }
                        Entity obj_lead1 = (Entity)service.Retrieve("lead", _enquiryId, new ColumnSet("hil_leadstatus", "hil_ticketnumber"));
                        _enquiryNo = obj_lead1.GetAttributeValue<string>("hil_ticketnumber");

                        QueryExpression qrEnquiries = new QueryExpression("lead");
                        qrEnquiries.ColumnSet = new ColumnSet("hil_ticketnumber");
                        ConditionExpression condEnquiry = new ConditionExpression("hil_originatinglead", ConditionOperator.Equal, _enquiryId);
                        qrEnquiries.Criteria.AddCondition(condEnquiry);
                        EntityCollection entColEnquiry = service.RetrieveMultiple(qrEnquiries);
                        foreach (Entity ent in entColEnquiry.Entities)
                        {
                            _enquiryNo += "," + ent.GetAttributeValue<string>("hil_ticketnumber");
                        }

                        obj_lead1.Attributes["hil_leadstatus"] = new OptionSetValue(2);
                        service.Update(obj_lead1);
                    }
                    obj_return.Status = "SUCCESS";
                    obj_return.EnquiryNo = _enquiryNo;
                }
            }
            catch (Exception ex)
            {
                obj_return.ErrorCode = "FAILURE";
                obj_return.ErrorDescription = "ENQUIRY : " + ex.Message.ToUpper();
            }
            return obj_return;
        }

        public List<EnquiryDetails> GetSalesEnquiry(EnquiryStatus _enquiryStatus)
        {
            List<EnquiryDetails> lstEnquiryDetails = new List<EnquiryDetails>();
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    bool _paramCheck = false;

                    string _filterStr = @"<filter type='and'><condition attribute='statecode' operator='eq' value='0' />";
                    if (_enquiryStatus.FromDate != null && _enquiryStatus.ToDate != null)
                    {
                        _filterStr += @"<condition attribute='createdon' operator='on-or-after' value='" + _enquiryStatus.FromDate + @"' />
                        <condition attribute='createdon' operator='on-or-before' value='" + _enquiryStatus.ToDate + @"' />";
                        _paramCheck = true;
                    }
                    if (_enquiryStatus.MobileNo != null && _enquiryStatus.MobileNo != "" && _enquiryStatus.MobileNo.Trim().Length > 0)
                    {
                        _filterStr += @"<condition attribute='mobilephone' operator='eq' value='" + _enquiryStatus.MobileNo + @"' />";
                        _paramCheck = true;
                    }
                    if (_enquiryStatus.EnquiryId != null && _enquiryStatus.EnquiryId != "" && _enquiryStatus.EnquiryId.Trim().Length > 0)
                    {
                        _filterStr += @"<condition attribute='hil_ticketnumber' operator='eq' value='" + _enquiryStatus.EnquiryId + @"' />";
                        _paramCheck = true;
                    }
                    _filterStr += @"</filter>";
                    if (!_paramCheck)
                    {
                        lstEnquiryDetails.Add(new EnquiryDetails() { ErrorCode = "204", Remarks = "ERROR!!! Please input From Date/ToDate/Mobileno/EnquiryId." });
                        return lstEnquiryDetails;
                    }
                    string _fetchxml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='lead'>
                        <attribute name='fullname' />
                        <attribute name='createdon' />
                        <attribute name='hil_selecteddivisionsname' />
                        <attribute name='hil_productdivision' />
                        <attribute name='leadid' />
                        <attribute name='hil_ticketnumber' />
                        <attribute name='hil_leadtype' />
                        <attribute name='description' />
                        <attribute name='hil_pincode' />
                        <attribute name='hil_leadstatus' />
                        <order attribute='createdon' descending='true' />" + _filterStr + @"
                      </entity>
                    </fetch>";

                    EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchxml));
                    if (entCol.Entities.Count > 0)
                    {
                        foreach (Entity ent in entCol.Entities)
                        {
                            lstEnquiryDetails.Add(new EnquiryDetails()
                            {
                                CreatedOn = ent.Contains("createdon") ? ent.GetAttributeValue<DateTime>("createdon").ToString() : "",
                                CustomerName = ent.Contains("fullname") ? ent.GetAttributeValue<string>("fullname").ToString() : "",
                                DivisionName = ent.Contains("hil_productdivision") ? ent.GetAttributeValue<EntityReference>("hil_productdivision").Name.ToString() : "",
                                //DivisionName = ent.Contains("hil_selecteddivisionsname") ? ent.GetAttributeValue<string>("hil_selecteddivisionsname").ToString() : "",
                                Remarks = ent.Contains("description") ? ent.GetAttributeValue<string>("description").ToString() : "",
                                EnquiryStatus = ent.Contains("hil_leadstatus") ? ent.FormattedValues["hil_leadstatus"].ToString() : "",
                                EnquiryType = ent.Contains("hil_leadtype") ? ent.FormattedValues["hil_leadtype"].ToString() : "",
                                PinCode = ent.Contains("hil_pincode") ? ent.GetAttributeValue<EntityReference>("hil_pincode").Name : "",
                                EnquiryId = ent.Contains("hil_ticketnumber") ? ent.GetAttributeValue<string>("hil_ticketnumber").ToString() : "",
                                ErrorCode = "Success"
                            });
                        }
                    }
                    else
                    {
                        lstEnquiryDetails.Add(new EnquiryDetails() { ErrorCode = "204", Remarks = "No Data found." });
                    }
                    return lstEnquiryDetails;
                }
                else
                {
                    lstEnquiryDetails.Add(new EnquiryDetails() { ErrorCode = "204", Remarks = "D365 Service Unavailable" });
                    return lstEnquiryDetails;
                }
            }
            catch (Exception ex)
            {
                lstEnquiryDetails.Add(new EnquiryDetails() { ErrorCode = "204", Remarks = "D365 Internal Server Error : " + ex.Message });
                return lstEnquiryDetails;
            }
        }
        public List<EnquiryType> GetEnquiryTypes(EnquiryType _enquiryCategory)
        {
            List<EnquiryType> lstEnquiryType = new List<EnquiryType>();
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    string _fetchxml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='hil_enquirytype'>
                            <attribute name='hil_enquirytypename' />
                            <attribute name='hil_enquirytypecode' />
                            <order attribute='hil_enquirytypename' descending='false' />
                            <filter type='and'>
                              <condition attribute='statecode' operator='eq' value='0' />
                              <condition attribute='hil_enquirycategory' operator='eq' value='" + _enquiryCategory.EnquiryCategory + @"' />
                            </filter>
                          </entity>
                        </fetch>";
                    EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchxml));
                    if (entCol.Entities.Count > 0)
                    {
                        foreach (Entity ent in entCol.Entities)
                        {
                            lstEnquiryType.Add(new EnquiryType()
                            {
                                EnquiryTypeName = ent.GetAttributeValue<string>("hil_enquirytypename"),
                                EnquiryTypeCode = ent.GetAttributeValue<Int32>("hil_enquirytypecode").ToString(),
                                EnquiryCategory = _enquiryCategory.EnquiryCategory
                            });
                        }
                    }
                    else
                    {
                        lstEnquiryType.Add(new EnquiryType() { EnquiryTypeCode = "ERROR", EnquiryTypeName = "No Data found." });
                    }
                    return lstEnquiryType;
                }
                else
                {
                    lstEnquiryType.Add(new EnquiryType() { EnquiryTypeCode = "ERROR", EnquiryTypeName = "D365 Service Unavailable" });
                    return lstEnquiryType;
                }
            }
            catch (Exception ex)
            {
                lstEnquiryType.Add(new EnquiryType() { EnquiryTypeCode = "ERROR", EnquiryTypeName = "D365 Internal Server Error : " + ex.Message });
                return lstEnquiryType;
            }
        }

        public List<EnquiryProductType> GetEnquiryProductTypes(EnquiryProductType _enquiryCategory)
        {
            List<EnquiryProductType> lstEnquiryProdType = new List<EnquiryProductType>();
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    string _fetchxml = string.Empty;

                    if (_enquiryCategory.EnquiryCategory == "1")
                    {
                        _fetchxml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='hil_typeofproduct'>
                            <attribute name='hil_typeofproductid' />
                            <attribute name='hil_name' />
                            <attribute name='hil_typeofenquiry' />
                            <attribute name='hil_index' />
                            <attribute name='hil_code' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                              <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            <link-entity name='hil_enquirytype' from='hil_enquirytypeid' to='hil_typeofenquiry' visible='false' link-type='outer' alias='eq'>
                              <attribute name='hil_enquirytypename' />
                              <attribute name='hil_enquirytypecode' />
                            </link-entity>
                          </entity>
                        </fetch>";
                        EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchxml));
                        if (entCol.Entities.Count > 0)
                        {
                            foreach (Entity ent in entCol.Entities)
                            {
                                lstEnquiryProdType.Add(new EnquiryProductType()
                                {
                                    EnquiryProductCode = ent.GetAttributeValue<string>("hil_code"),
                                    EnquiryProductName = ent.GetAttributeValue<string>("hil_name"),
                                    EnquiryTypeName = ent.GetAttributeValue<AliasedValue>("eq.hil_enquirytypename").Value.ToString(),
                                    EnquiryTypeCode = ent.GetAttributeValue<AliasedValue>("eq.hil_enquirytypecode").Value.ToString(),
                                });
                            }
                        }
                        else
                        {
                            lstEnquiryProdType.Add(new EnquiryProductType() { EnquiryProductCode = "ERROR", EnquiryProductName = "No Data found." });
                        }
                    }
                    else if (_enquiryCategory.EnquiryCategory == "2")
                    {
                        _fetchxml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='product'>
                            <attribute name='name' />
                            <attribute name='productnumber' />
                            <attribute name='hil_sapcode' />
                            <order attribute='name' descending='false' />
                            <filter type='and'>
                              <condition attribute='hil_hierarchylevel' operator='eq' value='2' />
                              <condition attribute='producttypecode' operator='eq' value='1' />
                            </filter>
                          </entity>
                        </fetch>";
                        EntityCollection entCol = service.RetrieveMultiple(new FetchExpression(_fetchxml));
                        if (entCol.Entities.Count > 0)
                        {
                            foreach (Entity ent in entCol.Entities)
                            {
                                lstEnquiryProdType.Add(new EnquiryProductType()
                                {
                                    EnquiryProductCode = ent.GetAttributeValue<string>("hil_sapcode"),
                                    EnquiryProductName = ent.GetAttributeValue<string>("name"),
                                });
                            }
                        }
                        else
                        {
                            lstEnquiryProdType.Add(new EnquiryProductType() { EnquiryProductCode = "ERROR", EnquiryProductName = "No Data found." });
                        }

                    }
                    return lstEnquiryProdType;
                }
                else
                {
                    lstEnquiryProdType.Add(new EnquiryProductType() { EnquiryProductCode = "ERROR", EnquiryProductName = "D365 Service Unavailable" });
                    return lstEnquiryProdType;
                }
            }
            catch (Exception ex)
            {
                lstEnquiryProdType.Add(new EnquiryProductType() { EnquiryProductCode = "ERROR", EnquiryProductName = "D365 Internal Server Error : " + ex.Message });
                return lstEnquiryProdType;
            }
        }
    }
    [DataContract]
    public class ReturnInfoVar
    {
        [DataMember(IsRequired = false)]
        public string ErrorCode { get; set; }

        [DataMember(IsRequired = false)]
        public string ErrorDescription { get; set; }

        [DataMember(IsRequired = false)]
        public Guid CustomerGuid { get; set; }

        [DataMember(IsRequired = false)]
        public string Status { get; set; }

        [DataMember(IsRequired = false)]
        public string EnquiryNo { get; set; }
    }
    [DataContract]
    public class EnquiryType
    {
        [DataMember]
        public string EnquiryCategory { get; set; }
        [DataMember]
        public string EnquiryTypeName { get; set; }
        [DataMember]
        public string EnquiryTypeCode { get; set; }
    }

    [DataContract]
    public class EnquiryProductType
    {
        [DataMember]
        public string EnquiryCategory { get; set; }
        [DataMember]
        public string EnquiryTypeName { get; set; }
        [DataMember]
        public string EnquiryTypeCode { get; set; }
        [DataMember]
        public string EnquiryProductCode { get; set; }
        [DataMember]
        public string EnquiryProductName { get; set; }

    }


    [DataContract]
    public class EnquiryDetails
    {
        [DataMember]
        public string CustomerName { get; set; }
        [DataMember]
        public string CreatedOn { get; set; }
        [DataMember]
        public string DivisionName { get; set; }
        [DataMember]
        public string Remarks { get; set; }
        [DataMember]
        public string EnquiryType { get; set; }
        [DataMember]
        public string PinCode { get; set; }
        [DataMember]
        public string EnquirySource { get; set; }
        [DataMember]
        public string EnquiryStatus { get; set; }
        [DataMember]
        public string ErrorCode { get; set; }
        [DataMember]
        public string EnquiryId { get; set; }
    }
}
