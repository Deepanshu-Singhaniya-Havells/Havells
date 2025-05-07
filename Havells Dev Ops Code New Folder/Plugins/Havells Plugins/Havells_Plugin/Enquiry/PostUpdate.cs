using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Havells_Plugin.HelperIntegration;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;

namespace Havells_Plugin.Enquiry
{
    public class PostUpdate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            string _divisionSAPCode = string.Empty;
            QueryExpression qsCampaign;
            EntityCollection collect_Campaign;

            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == "lead" && context.MessageName.ToUpper() == "UPDATE")
                {
                    tracingService.Trace("1");
                    Entity entity = (Entity)context.InputParameters["Target"];
                    Lead PostImageLead = ((Entity)context.PostEntityImages["image"]).ToEntity<Lead>();
                    tracingService.Trace("12");
                    if (entity.Contains("hil_leadstatus") && entity.GetAttributeValue<OptionSetValue>("hil_leadstatus").Value == 2)
                    {
                        tracingService.Trace("2");
                        if (PostImageLead.Contains("hil_productenquirynumber") && PostImageLead.GetAttributeValue<OptionSetValue>("hil_productenquirynumber").Value == 910590000) //Enquiry Type : Single
                        {
                            tracingService.Trace("3");
                            EMS_Request eMS_Request = new EMS_Request();
                            Data data = new Data();
                            data.TicketNo = PostImageLead["hil_ticketnumber"].ToString();
                            data.TicketParentId = PostImageLead["hil_ticketnumber"].ToString();
                            data.FirstName = PostImageLead.FirstName != null ? PostImageLead.FirstName : string.Empty;
                            data.LastName = PostImageLead.LastName != null ? PostImageLead.LastName : string.Empty;
                            data.Remarks = PostImageLead.Description != null ? PostImageLead.Description : string.Empty;
                            if (PostImageLead.CampaignId != null)
                            {
                                qsCampaign = new QueryExpression("campaign");
                                qsCampaign.ColumnSet = new ColumnSet("codename", "name");
                                qsCampaign.Criteria.AddCondition(new ConditionExpression("campaignid", ConditionOperator.Equal, PostImageLead.CampaignId.Id));
                                collect_Campaign = service.RetrieveMultiple(qsCampaign);
                                if (collect_Campaign.Entities.Count > 0)
                                {
                                    data.CampaignCode = collect_Campaign.Entities[0].GetAttributeValue<string>("codename");
                                    data.CampaignDesc = collect_Campaign.Entities[0].GetAttributeValue<string>("name");
                                }
                            }

                            string Consumer_address = PostImageLead.Address1_Line1 != null ? PostImageLead.Address1_Line1 + ", " : "";

                            if (PostImageLead.Address1_Line2 != null)
                                Consumer_address = Consumer_address + PostImageLead.Address1_Line2 + ", ";
                            if (PostImageLead.Address1_Line3 != null)
                                Consumer_address = Consumer_address + PostImageLead.Address1_Line3 + ", ";
                            if (PostImageLead.hil_City != null)
                                Consumer_address = Consumer_address + PostImageLead.hil_City.Name + ", ";
                            if (PostImageLead.hil_State != null)
                                Consumer_address = Consumer_address + PostImageLead.hil_State.Name + ", ";
                            if (PostImageLead.hil_PinCode != null)
                                Consumer_address = Consumer_address + PostImageLead.hil_PinCode.Name + " ";
                            if (PostImageLead.hil_Region != null)
                                Consumer_address = Consumer_address + "Region: " + PostImageLead.hil_Region.Name + ", ";
                            if (PostImageLead.hil_LandMark != null)
                                Consumer_address = Consumer_address + "Landmark: " + PostImageLead.hil_LandMark + ", ";

                            data.Address = Consumer_address;
                            data.MobileNo = PostImageLead.MobilePhone != null ? PostImageLead.MobilePhone : string.Empty;
                            data.Email = PostImageLead.EMailAddress1 != null ? PostImageLead.EMailAddress1 : string.Empty;
                            data.PinCode = PostImageLead.hil_PinCode != null ? PostImageLead.hil_PinCode.Name : string.Empty;

                            _divisionSAPCode = GetDivisionSAPCode(service, PostImageLead.hil_ProductDivision.Id);
                            data.DivisionCode = _divisionSAPCode;

                            //if (_divisionSAPCode!=string.Empty)
                            //    data.DivisionCode = _divisionSAPCode;
                            //else
                            //    throw new InvalidPluginExecutionException("Product Division SAP Code is not defined.");

                            data.EmployeeType = "EMP";
                            //data.EnquiryType = PostImageLead.hil_LeadType != null ? PostImageLead.FormattedValues["hil_leadtype"] : string.Empty;
                            data.EnquiryType = PostImageLead.hil_LeadType != null ? PostImageLead.hil_LeadType.Value.ToString() : string.Empty;
                            data.Interest = PostImageLead.Contains("hil_interest") ? PostImageLead.FormattedValues["hil_interest"] : string.Empty;
                            eMS_Request.Data = data;
                            tracingService.Trace("4");
                            EMS_Response eMS_Response = PostEnquiry(service, eMS_Request);
                            if (eMS_Response.Results != null && eMS_Response.Results.Count > 0)
                            {
                                tracingService.Trace("5");
                                Lead leadUpdateObj = new Lead();
                                leadUpdateObj.Id = entity.Id;
                                leadUpdateObj["hil_message"] = eMS_Response.Results[0].ErrorMessage + " 1 Division Code " + _divisionSAPCode + " ParentLeadNo:" + PostImageLead["hil_ticketnumber"].ToString();
                                leadUpdateObj["hil_jsondatapacket"] = eMS_Response.DataPacket;
                                service.Update(leadUpdateObj);
                            }
                            tracingService.Trace("6");
                        }
                        else
                        {
                            QueryExpression qsLeadProduct = new QueryExpression("hil_leadproduct");
                            qsLeadProduct.ColumnSet = new ColumnSet("hil_enquiry", "hil_productdivision");
                            qsLeadProduct.Criteria.AddCondition(new ConditionExpression("hil_enquiry", ConditionOperator.Equal, entity.Id));
                            EntityCollection collect_LeadProduct = service.RetrieveMultiple(qsLeadProduct);

                            if (collect_LeadProduct.Entities.Count == 0)
                            {
                                throw new InvalidPluginExecutionException("Select At least One product");
                            }

                            int countflag = 1;
                            Guid _originatingLeadId = Guid.Empty;
                            string _originatingLeadNo = string.Empty;
                            int leadSource = 8;
                            foreach (Entity lead in collect_LeadProduct.Entities)
                            {
                                if (countflag == 1)
                                {
                                    if (lead.Contains("hil_productdivision"))
                                    {
                                        _originatingLeadId = entity.Id;

                                        leadSource = PostImageLead.GetAttributeValue<OptionSetValue>("leadsourcecode").Value;
                                        EMS_Request eMS_Request = new EMS_Request();
                                        Data data = new Data();
                                        data.TicketNo = PostImageLead["hil_ticketnumber"].ToString();
                                        _originatingLeadNo = PostImageLead["hil_ticketnumber"].ToString();

                                        data.TicketParentId = _originatingLeadNo;
                                        data.FirstName = PostImageLead.FirstName != null ? PostImageLead.FirstName : string.Empty;
                                        data.LastName = PostImageLead.LastName != null ? PostImageLead.LastName : string.Empty;
                                        data.Remarks = PostImageLead.Description != null ? PostImageLead.Description : string.Empty;

                                        qsCampaign = new QueryExpression("campaign");
                                        qsCampaign.ColumnSet = new ColumnSet("codename", "name");
                                        qsCampaign.Criteria.AddCondition(new ConditionExpression("campaignid", ConditionOperator.Equal, PostImageLead.CampaignId.Id));
                                        collect_Campaign = service.RetrieveMultiple(qsCampaign);
                                        if (collect_Campaign.Entities.Count > 0)
                                        {
                                            data.CampaignCode = collect_Campaign.Entities[0].GetAttributeValue<string>("codename");
                                            data.CampaignDesc = collect_Campaign.Entities[0].GetAttributeValue<string>("name");
                                        }

                                        string Consumer_address = PostImageLead.Address1_Line1 != null ? PostImageLead.Address1_Line1 + ", " : "";

                                        if (PostImageLead.Address1_Line2 != null)
                                            Consumer_address = Consumer_address + PostImageLead.Address1_Line2 + ", ";
                                        if (PostImageLead.Address1_Line3 != null)
                                            Consumer_address = Consumer_address + PostImageLead.Address1_Line3 + ", ";
                                        if (PostImageLead.hil_City != null)
                                            Consumer_address = Consumer_address + PostImageLead.hil_City.Name + ", ";
                                        if (PostImageLead.hil_State != null)
                                            Consumer_address = Consumer_address + PostImageLead.hil_State.Name + ", ";
                                        if (PostImageLead.hil_PinCode != null)
                                            Consumer_address = Consumer_address + PostImageLead.hil_PinCode.Name + " ";
                                        if (PostImageLead.hil_Region != null)
                                            Consumer_address = Consumer_address + "Region: " + PostImageLead.hil_Region.Name + ", ";
                                        if (PostImageLead.hil_LandMark != null)
                                            Consumer_address = Consumer_address + "Landmark: " + PostImageLead.hil_LandMark + ", ";

                                        data.Address = Consumer_address;
                                        data.MobileNo = PostImageLead.MobilePhone != null ? PostImageLead.MobilePhone : string.Empty;
                                        data.Email = PostImageLead.EMailAddress1 != null ? PostImageLead.EMailAddress1 : string.Empty;
                                        data.PinCode = PostImageLead.hil_PinCode != null ? PostImageLead.hil_PinCode.Name : string.Empty;

                                        _divisionSAPCode = string.Empty;
                                        _divisionSAPCode = GetDivisionSAPCode(service, lead.GetAttributeValue<EntityReference>("hil_productdivision").Id);
                                        //data.DivisionCode = lead.Contains("hil_productdivision") ? lead.GetAttributeValue<EntityReference>("hil_productdivision").Name : string.Empty;
                                        data.DivisionCode = _divisionSAPCode;
                                        data.EmployeeType = "EMP";
                                        //data.EnquiryType = PostImageLead.hil_LeadType != null ? PostImageLead.FormattedValues["hil_leadtype"] : string.Empty;
                                        data.EnquiryType = PostImageLead.hil_LeadType != null ? PostImageLead.hil_LeadType.Value.ToString() : string.Empty;
                                        data.Interest = PostImageLead.Contains("hil_interest") ? PostImageLead.FormattedValues["hil_interest"] : string.Empty;
                                        eMS_Request.Data = data;

                                        EMS_Response eMS_Response = PostEnquiry(service, eMS_Request);
                                        if (eMS_Response.Results != null && eMS_Response.Results.Count > 0)
                                        {
                                            Lead leadUpdateObj = new Lead();
                                            leadUpdateObj.Id = entity.Id;
                                            leadUpdateObj["hil_message"] = eMS_Response.Results[0].ErrorMessage + " 2 Division Code " + _divisionSAPCode + " ParentLeadNo:" + _originatingLeadNo;
                                            leadUpdateObj["hil_productdivision"] = new EntityReference("product", lead.GetAttributeValue<EntityReference>("hil_productdivision").Id);
                                            leadUpdateObj["hil_jsondatapacket"] = eMS_Response.DataPacket;
                                            //leadUpdateObj["hil_originatinglead"] = new EntityReference("lead", entity.Id);
                                            service.Update(leadUpdateObj);
                                            countflag++;
                                        }
                                    }
                                }
                                else
                                {
                                    Lead CreateleadObj = new Lead();

                                    if (lead.Contains("hil_productdivision"))
                                    {
                                        CreateleadObj["hil_productdivision"] = new EntityReference("product", lead.GetAttributeValue<EntityReference>("hil_productdivision").Id);
                                    }
                                    if (PostImageLead.Contains("subject"))
                                    {
                                        CreateleadObj["subject"] = PostImageLead.GetAttributeValue<string>("subject");
                                    }
                                    if (PostImageLead.Contains("firstname"))
                                    {
                                        CreateleadObj["firstname"] = PostImageLead.GetAttributeValue<string>("firstname");
                                    }
                                    if (PostImageLead.Contains("lastname"))
                                    {
                                        CreateleadObj["lastname"] = PostImageLead.GetAttributeValue<string>("lastname");
                                    }
                                    if (PostImageLead.Contains("mobilephone"))
                                    {
                                        CreateleadObj["mobilephone"] = PostImageLead.GetAttributeValue<string>("mobilephone");
                                    }
                                    if (PostImageLead.Contains("hil_alternatemobileno"))
                                    {
                                        CreateleadObj["hil_alternatemobileno"] = PostImageLead.GetAttributeValue<string>("hil_alternatemobileno");
                                    }
                                    if (PostImageLead.Contains("emailaddress1"))
                                    {
                                        CreateleadObj["emailaddress1"] = PostImageLead.GetAttributeValue<string>("emailaddress1");
                                    }
                                    if (PostImageLead.Contains("hil_productenquirynumber"))
                                    {
                                        CreateleadObj["hil_productenquirynumber"] = new OptionSetValue(PostImageLead.GetAttributeValue<OptionSetValue>("hil_productenquirynumber").Value);
                                    }
                                    if (PostImageLead.Contains("hil_leadtype"))
                                    {
                                        CreateleadObj["hil_leadtype"] = new OptionSetValue(PostImageLead.GetAttributeValue<OptionSetValue>("hil_leadtype").Value);
                                    }
                                    if (PostImageLead.Contains("prioritycode"))
                                    {
                                        CreateleadObj["prioritycode"] = new OptionSetValue(PostImageLead.GetAttributeValue<OptionSetValue>("prioritycode").Value);
                                    }
                                    if (PostImageLead.Contains("address1_line1"))
                                    {
                                        CreateleadObj["address1_line1"] = PostImageLead.GetAttributeValue<string>("address1_line1");
                                    }
                                    if (PostImageLead.Contains("address1_line2"))
                                    {
                                        CreateleadObj["address1_line2"] = PostImageLead.GetAttributeValue<string>("address1_line2");
                                    }
                                    if (PostImageLead.Contains("address1_line3"))
                                    {
                                        CreateleadObj["address1_line3"] = PostImageLead.GetAttributeValue<string>("address1_line3");
                                    }
                                    if (PostImageLead.Contains("hil_landmark"))
                                    {
                                        CreateleadObj["hil_landmark"] = PostImageLead.GetAttributeValue<string>("hil_landmark");
                                    }
                                    if (PostImageLead.Contains("hil_utmcode"))
                                    {
                                        CreateleadObj["hil_utmcode"] = PostImageLead.GetAttributeValue<string>("hil_utmcode");
                                    }

                                    if (PostImageLead.Contains("hil_pincode"))
                                    {
                                        CreateleadObj["hil_pincode"] = new EntityReference("hil_pincode", PostImageLead.GetAttributeValue<EntityReference>("hil_pincode").Id);
                                    }
                                    if (PostImageLead.Contains("hil_salesoffice"))
                                    {
                                        CreateleadObj["hil_salesoffice"] = new EntityReference("hil_salesoffice", PostImageLead.GetAttributeValue<EntityReference>("hil_salesoffice").Id);
                                    }
                                    if (PostImageLead.Contains("hil_city"))
                                    {
                                        CreateleadObj["hil_city"] = new EntityReference("hil_city", PostImageLead.GetAttributeValue<EntityReference>("hil_city").Id);
                                    }
                                    if (PostImageLead.Contains("hil_district"))
                                    {
                                        CreateleadObj["hil_district"] = new EntityReference("hil_district", PostImageLead.GetAttributeValue<EntityReference>("hil_district").Id);
                                    }
                                    if (PostImageLead.Contains("hil_state"))
                                    {
                                        CreateleadObj["hil_state"] = new EntityReference("hil_state", PostImageLead.GetAttributeValue<EntityReference>("hil_state").Id);
                                    }
                                    if (PostImageLead.Contains("hil_country"))
                                    {
                                        CreateleadObj["hil_country"] = new EntityReference("hil_country", PostImageLead.GetAttributeValue<EntityReference>("hil_country").Id);
                                    }
                                    if (PostImageLead.Contains("hil_region"))
                                    {
                                        CreateleadObj["hil_region"] = new EntityReference("hil_region", PostImageLead.GetAttributeValue<EntityReference>("hil_region").Id);
                                    }
                                    if (PostImageLead.Contains("hil_branchoffice"))
                                    {
                                        CreateleadObj["hil_branchoffice"] = new EntityReference("hil_branch", PostImageLead.GetAttributeValue<EntityReference>("hil_branchoffice").Id);
                                    }
                                    if (PostImageLead.Contains("preferredcontactmethodcode"))
                                    {
                                        CreateleadObj["preferredcontactmethodcode"] = new OptionSetValue(PostImageLead.GetAttributeValue<OptionSetValue>("preferredcontactmethodcode").Value);
                                    }
                                    if (PostImageLead.Contains("hil_preferredtimeofcontact"))
                                    {
                                        CreateleadObj["hil_preferredtimeofcontact"] = PostImageLead.GetAttributeValue<DateTime>("hil_preferredtimeofcontact");
                                    }
                                    if (PostImageLead.Contains("campaignid"))
                                    {
                                        CreateleadObj["campaignid"] = PostImageLead.CampaignId;
                                    }
                                    //Creating Contact 
                                    Guid _primaryContact = Guid.Empty;
                                    QueryExpression qsprimaryContact = new QueryExpression("contact");
                                    qsprimaryContact.ColumnSet = new ColumnSet("firstname");
                                    ConditionExpression cond1 = new ConditionExpression("mobilephone", ConditionOperator.Equal, PostImageLead.GetAttributeValue<string>("mobilephone"));
                                    qsprimaryContact.Criteria.AddCondition(cond1);
                                    EntityCollection collct_primaryContact = service.RetrieveMultiple(qsprimaryContact);
                                    if (collct_primaryContact.Entities.Count > 0)
                                    {
                                        foreach (Entity primaryContact_record in collct_primaryContact.Entities)
                                        {
                                            _primaryContact = primaryContact_record.Id;
                                        }
                                        CreateleadObj["parentcontactid"] = new EntityReference("contact", _primaryContact);
                                    }
                                    //Updating Parent Lead Id
                                    if (_originatingLeadId != Guid.Empty)
                                    {
                                        CreateleadObj["hil_originatinglead"] = new EntityReference("lead", _originatingLeadId);
                                    }
                                    CreateleadObj["leadsourcecode"] = new OptionSetValue(leadSource); //Originated from where
                                    CreateleadObj["hil_leadstatus"] = new OptionSetValue(2); //Submit
                                    if (PostImageLead.Contains("description"))
                                    {
                                        CreateleadObj["description"] = PostImageLead.GetAttributeValue<string>("description");
                                    }

                                    Guid leadid = service.Create(CreateleadObj);

                                    Lead leadcreated = (Lead)service.Retrieve(Lead.EntityLogicalName, leadid, new ColumnSet(new string[] { "hil_ticketnumber", "hil_productdivision" }));

                                    EMS_Request eMS_Request = new EMS_Request();
                                    Data data = new Data();
                                    data.TicketNo = leadcreated["hil_ticketnumber"].ToString();
                                    data.TicketParentId = _originatingLeadNo;

                                    data.FirstName = PostImageLead.FirstName != null ? PostImageLead.FirstName : string.Empty;
                                    data.LastName = PostImageLead.LastName != null ? PostImageLead.LastName : string.Empty;
                                    data.Remarks = PostImageLead.Description != null ? PostImageLead.Description : string.Empty;

                                    qsCampaign = new QueryExpression("campaign");
                                    qsCampaign.ColumnSet = new ColumnSet("codename", "name");
                                    qsCampaign.Criteria.AddCondition(new ConditionExpression("campaignid", ConditionOperator.Equal, PostImageLead.CampaignId.Id));
                                    collect_Campaign = service.RetrieveMultiple(qsCampaign);
                                    if (collect_Campaign.Entities.Count > 0)
                                    {
                                        data.CampaignCode = collect_Campaign.Entities[0].GetAttributeValue<string>("codename");
                                        data.CampaignDesc = collect_Campaign.Entities[0].GetAttributeValue<string>("name");
                                    }

                                    string Consumer_address = PostImageLead.Address1_Line1 != null ? PostImageLead.Address1_Line1 + ", " : "";

                                    if (PostImageLead.Address1_Line2 != null)
                                        Consumer_address = Consumer_address + PostImageLead.Address1_Line2 + ", ";
                                    if (PostImageLead.Address1_Line3 != null)
                                        Consumer_address = Consumer_address + PostImageLead.Address1_Line3 + ", ";
                                    if (PostImageLead.hil_City != null)
                                        Consumer_address = Consumer_address + PostImageLead.hil_City.Name + ", ";
                                    if (PostImageLead.hil_State != null)
                                        Consumer_address = Consumer_address + PostImageLead.hil_State.Name + ", ";
                                    if (PostImageLead.hil_PinCode != null)
                                        Consumer_address = Consumer_address + PostImageLead.hil_PinCode.Name + " ";
                                    if (PostImageLead.hil_Region != null)
                                        Consumer_address = Consumer_address + "Region: " + PostImageLead.hil_Region.Name + ", ";
                                    if (PostImageLead.hil_LandMark != null)
                                        Consumer_address = Consumer_address + "Landmark: " + PostImageLead.hil_LandMark + ", ";

                                    data.Address = Consumer_address;
                                    data.MobileNo = PostImageLead.MobilePhone != null ? PostImageLead.MobilePhone : string.Empty;
                                    data.Email = PostImageLead.EMailAddress1 != null ? PostImageLead.EMailAddress1 : string.Empty;
                                    data.PinCode = PostImageLead.hil_PinCode != null ? PostImageLead.hil_PinCode.Name : string.Empty;

                                    _divisionSAPCode = string.Empty;
                                    _divisionSAPCode = GetDivisionSAPCode(service, lead.GetAttributeValue<EntityReference>("hil_productdivision").Id);
                                    data.DivisionCode = _divisionSAPCode;
                                    //data.DivisionCode = lead.Contains("hil_productdivision") ? lead.GetAttributeValue<EntityReference>("hil_productdivision").Name : string.Empty;

                                    data.EmployeeType = "EMP";
                                    //data.EnquiryType = PostImageLead.hil_LeadType != null ? PostImageLead.FormattedValues["hil_leadtype"] : string.Empty;
                                    data.EnquiryType = PostImageLead.hil_LeadType != null ? PostImageLead.hil_LeadType.Value.ToString() : string.Empty;
                                    data.Interest = PostImageLead.Contains("hil_interest") ? PostImageLead.FormattedValues["hil_interest"] : string.Empty;
                                    eMS_Request.Data = data;

                                    EMS_Response eMS_Response = PostEnquiry(service, eMS_Request);
                                    if (eMS_Response.Results != null && eMS_Response.Results.Count > 0)
                                    {
                                        Lead leadUpdateObj = new Lead();
                                        leadUpdateObj.Id = leadid;
                                        leadUpdateObj["hil_message"] = eMS_Response.Results[0].ErrorMessage + " 3 Division Code " + _divisionSAPCode + " ParentLeadNo:" + _originatingLeadNo;
                                        leadUpdateObj["hil_jsondatapacket"] = eMS_Response.DataPacket;
                                        service.Update(leadUpdateObj);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException("Error in Havells_Plugin.Enquiry.PostUpdate " + e.Message);
            }
        }
        public string GetDivisionSAPCode(IOrganizationService service, Guid divisionGUID)
        {
            string _divisionSAPCode = string.Empty;
            try
            {
                QueryExpression qsProduct = new QueryExpression("product");
                qsProduct.ColumnSet = new ColumnSet("hil_sapcode");
                qsProduct.Criteria.AddCondition(new ConditionExpression("productid", ConditionOperator.Equal, divisionGUID));
                EntityCollection collect_Product = service.RetrieveMultiple(qsProduct);
                if (collect_Product.Entities.Count > 0)
                {
                    if (collect_Product.Entities[0].Attributes.Contains("hil_sapcode"))
                    {
                        _divisionSAPCode = collect_Product.Entities[0].Attributes["hil_sapcode"].ToString();
                    }
                }
            }
            catch { }
            return _divisionSAPCode;
        }
        public EMS_Response PostEnquiry(IOrganizationService service, EMS_Request eMS_Request)
        {
            EMS_Response resp = new EMS_Response();
            using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
            {
                String sUserName = String.Empty;
                String sPassword = String.Empty;
                var obj2 = from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
                           where _IConfig.hil_name == "EMS_Credentials"
                           select new { _IConfig };
                foreach (var iobj2 in obj2)
                {
                    if (iobj2._IConfig.hil_Username != String.Empty)
                        sUserName = iobj2._IConfig.hil_Username;
                    if (iobj2._IConfig.hil_Password != String.Empty)
                        sPassword = iobj2._IConfig.hil_Password;
                }
                var obj = from _IConfig in orgContext.CreateQuery<hil_integrationconfiguration>()
                          where _IConfig.hil_name == "EMS_Integration"
                          select new { _IConfig.hil_Url };

                foreach (var iobj in obj)
                {
                    if (iobj.hil_Url != null)
                    {
                        try
                        {
                            var Json = JsonConvert.SerializeObject(eMS_Request);
                            //throw new InvalidPluginExecutionException("Error while EMS Response: " + Json);
                            WebRequest request = WebRequest.Create(iobj.hil_Url);
                            string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
                            request.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
                            // Set the Method property of the request to POST.  
                            request.Method = "POST";
                            byte[] byteArray = Encoding.UTF8.GetBytes(Json);
                            // Set the ContentType property of the WebRequest.  
                            request.ContentType = "application/x-www-form-urlencoded";
                            // Set the ContentLength property of the WebRequest.  
                            request.ContentLength = byteArray.Length;
                            // Get the request stream.  
                            Stream dataStream = request.GetRequestStream();
                            // Write the data to the request stream.  
                            dataStream.Write(byteArray, 0, byteArray.Length);
                            // Close the Stream object.  
                            dataStream.Close();
                            // Get the response.  
                            WebResponse response = request.GetResponse();
                            // Display the status.  
                            Console.WriteLine(((HttpWebResponse)response).StatusDescription);
                            // Get the stream containing content returned by the server.  
                            dataStream = response.GetResponseStream();
                            // Open the stream using a StreamReader for easy access.  
                            StreamReader reader = new StreamReader(dataStream);
                            // Read the content.  
                            string responseFromServer = reader.ReadToEnd();
                            try
                            {
                                resp = JsonConvert.DeserializeObject<EMS_Response>(responseFromServer);
                                resp.DataPacket = Json;
                            }
                            catch (InvalidPluginExecutionException e)
                            {
                                resp.DataPacket = Json;
                                throw new InvalidPluginExecutionException("Error while EMS Response: " + e.Message);
                            }
                        }
                        catch (InvalidPluginExecutionException e)
                        {
                            throw new InvalidPluginExecutionException("Error while Interfacing Data to SAP: " + e.Message);
                        }
                    }
                }
            }
            return resp;
        }
    }
}
