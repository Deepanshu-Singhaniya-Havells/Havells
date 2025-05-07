using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Net;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel.Description;
using System.Configuration;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class AuthenticateCustomer
    {
        [DataMember]
        public string UName { get; set; }
        [DataMember]
        public string Pwd { get; set; }
        [DataMember]
        public string Method { get; set; }
        [DataMember(IsRequired = false)]
        public string ContGuid { get; set; }
        [DataMember(IsRequired = false)]
        public bool status { get; set; }
        [DataMember(IsRequired = false)]
        public bool IfValidated { get; set; }
        [DataMember(IsRequired = false)]
        public string FirstName { get; set; }
        [DataMember(IsRequired = false)]
        public string LastName { get; set; }
        [DataMember(IsRequired = false)]
        public string MobileNumber { get; set; }
        [DataMember(IsRequired = false)]
        public bool NewAPK { get; set; }
        [DataMember(IsRequired = false)]
        public string PinCode { get; set; }
        [DataMember(IsRequired = false)]
        public int ERROR_CODE { get; set; }
        public AuthenticateCustomer ValidateThisCustomer(AuthenticateCustomer bridge)
        {
            NewAPK = true;
            if (!NewAPK) throw new Exception();
            IOrganizationService service = ConnectToCRM.GetOrgService();//get org service obj for connection
            try
            {
                if (bridge.Method == "CustomerLogin")
                {
                    QueryExpression Query = new QueryExpression(Contact.EntityLogicalName);
                    Query.ColumnSet = new ColumnSet("emailaddress1", "hil_password", "firstname", "lastname", "mobilephone");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("emailaddress1", ConditionOperator.Equal, bridge.UName);
                    EntityCollection Found = service.RetrieveMultiple(Query);
                    if (Found.Entities.Count == 1)
                    {
                        Contact Cont = Found.Entities[0].ToEntity<Contact>();
                        if (Cont.hil_Password == bridge.Pwd)
                        {
                            bridge.ContGuid = Cont.ContactId.Value.ToString();
                            if (Cont.MobilePhone != null)
                                bridge.MobileNumber = Cont.MobilePhone;
                            bridge.IfValidated = true;
                            if (Cont.LastName != null)
                                bridge.LastName = Cont.LastName;
                            if (Cont.FirstName != null)
                                bridge.FirstName = Cont.FirstName;
                            bridge.PinCode = GetPermanentAddressPinCode(service, Cont.Id);
                            bridge.status = true;
                            bridge.ERROR_CODE = 0;
                        }
                        else
                        {
                            bridge.ContGuid = "NOT FOUND";
                            bridge.MobileNumber = "NOT FOUND";
                            bridge.IfValidated = false;
                            bridge.LastName = "NOT FOUND";
                            bridge.FirstName = "NOT FOUND";
                            bridge.PinCode = "NOT FOUND";
                            bridge.status = false;
                            bridge.ERROR_CODE = 2;
                        }
                    }
                    else
                    {
                        QueryExpression Query1 = new QueryExpression(Contact.EntityLogicalName);
                        Query1.ColumnSet = new ColumnSet("emailaddress1", "hil_password", "firstname", "lastname", "mobilephone");
                        Query1.Criteria = new FilterExpression(LogicalOperator.And);
                        Query1.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, bridge.UName);
                        EntityCollection Found1 = service.RetrieveMultiple(Query1);
                        if (Found1.Entities.Count == 1)
                        {
                            Contact Cont1 = Found1.Entities[0].ToEntity<Contact>();
                            if (Cont1.hil_Password == bridge.Pwd)
                            {
                                bridge.ContGuid = Cont1.Id.ToString();
                                if (Cont1.MobilePhone != null)
                                    bridge.MobileNumber = Cont1.MobilePhone;
                                bridge.IfValidated = true;
                                if (Cont1.LastName != null)
                                    bridge.LastName = Cont1.LastName;
                                if (Cont1.FirstName != null)
                                    bridge.FirstName = Cont1.FirstName;
                                bridge.PinCode = GetPermanentAddressPinCode(service, Cont1.Id);
                                bridge.status = true;
                                bridge.ERROR_CODE = 0;
                            }
                            else
                            {
                                bridge.ContGuid = "NOT FOUND";
                                bridge.MobileNumber = "NOT FOUND";
                                bridge.IfValidated = false;
                                bridge.LastName = "NOT FOUND";
                                bridge.FirstName = "NOT FOUND";
                                bridge.PinCode = "NOT FOUND";
                                bridge.status = false;
                                bridge.ERROR_CODE = 2;
                            }
                        }
                        else
                        {
                            bridge.ContGuid = "NOT FOUND";
                            bridge.MobileNumber = "NOT FOUND";
                            bridge.IfValidated = false;
                            bridge.LastName = "NOT FOUND";
                            bridge.FirstName = "NOT FOUND";
                            bridge.PinCode = "NOT FOUND";
                            bridge.status = false;
                            bridge.ERROR_CODE = 1;
                        }
                    }
                    if (bridge.MobileNumber != "" && bridge.MobileNumber.Trim().Length > 0) {
                        SMSService.sendSMS(service, bridge.MobileNumber, "V1_D365_ConsumerAppMigration");
                    }
                    //OrganizationRequest req = new OrganizationRequest("hil_ConsumerApp_Validated2dd12249a73e811a95a000d3af068d4");
                    //req["Method"] = bridge.Method;
                    //req["UserName"] = bridge.UName;
                    //req["Password"] = bridge.Pwd;
                    //OrganizationResponse response = service.Execute(req);
                    //bridge.ContGuid = (string)response["ContGuid"];
                    //bridge.IfValidated = (bool)response["IfValidated"];
                    //bridge.status = (bool)response["Status"];
                    //bridge.FirstName =(string)response["FirstName"];
                    //bridge.LastName = (string)response["LastName"];
                    //bridge.MobileNumber = (string)response["MobileNumber"];
                    //bridge.PinCode = (string)response["PinCode"];
                }
            }
            catch (Exception ex)
            {
                bridge.ContGuid = ex.StackTrace.ToUpper();
                bridge.MobileNumber = "NOT FOUND";
                bridge.IfValidated = false;
                bridge.LastName = "NOT FOUND";
                bridge.FirstName = "NOT FOUND";
                bridge.PinCode = "NOT FOUND";
                bridge.status = false;
                bridge.ERROR_CODE = 3;
            }
            return (bridge);
        }
        public string GetPermanentAddressPinCode(IOrganizationService service, Guid ContId)
        {
            string Pin = "";
            QueryExpression Query = new QueryExpression(hil_address.EntityLogicalName);
            Query.ColumnSet = new ColumnSet("hil_pincode");
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, ContId);
            Query.Criteria.AddCondition("hil_addresstype", ConditionOperator.Equal, 1);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                hil_address iAddress = Found.Entities[0].ToEntity<hil_address>();
                Pin = iAddress.hil_PinCode.Name;
            }
            return Pin;
        }
    }
}
