using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class SFA_ValidateCustomer
    {
        [DataMember]
        public string CustomerMobileNo { get; set; }
        [DataMember]
        public string CustomerName { get; set; }
        [DataMember]
        public string CustomerEmailId { get; set; }
        [DataMember]
        public Guid CustomerGuid { get; set; }
        [DataMember]
        public bool ResultStatus { get; set; }
        [DataMember]
        public string ResultMessage { get; set; }
        [DataMember]
        public List<SFA_AddressBook> AddressBook { get; set; }

        public SFA_ValidateCustomer ValidateCustomer(SFA_ValidateCustomer customerData) {
            Guid customerGuid = Guid.Empty;
            SFA_ValidateCustomer objValidateCustomer;
            SFA_AddressBook objAddress;
            List<SFA_AddressBook> lstAddressBook = new List<SFA_AddressBook>();
            EntityCollection entcoll;
            QueryExpression Query;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (service != null)
                {
                    if (customerData.CustomerMobileNo.Trim().Length == 0 || customerData.CustomerMobileNo == null)
                    {
                        objValidateCustomer = new SFA_ValidateCustomer { ResultStatus = false, ResultMessage = "Customer Mobile No. is required.", AddressBook = new List<SFA_AddressBook>() };
                        return objValidateCustomer;
                    }

                    Query = new QueryExpression("contact");
                    Query.ColumnSet = new ColumnSet("fullname", "emailaddress1");
                    Query.Criteria = new FilterExpression(LogicalOperator.And);
                    Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, customerData.CustomerMobileNo);
                    entcoll = service.RetrieveMultiple(Query);

                    if (entcoll.Entities.Count == 0)
                    {
                        objValidateCustomer = new SFA_ValidateCustomer { ResultStatus = false, ResultMessage = "Customer Mobile No. does not exist.", AddressBook = new List<SFA_AddressBook>() };
                        return objValidateCustomer;
                    }
                    else
                    {
                        objValidateCustomer = new SFA_ValidateCustomer();
                        objValidateCustomer.CustomerMobileNo = customerData.CustomerMobileNo;

                        objValidateCustomer.CustomerName = entcoll.Entities[0].GetAttributeValue<string>("fullname");

                        objValidateCustomer.CustomerEmailId = entcoll.Entities[0].GetAttributeValue<string>("emailaddress1");

                        objValidateCustomer.CustomerGuid = entcoll.Entities[0].Id;
                        objValidateCustomer.ResultStatus = true;
                        objValidateCustomer.ResultMessage = "SUCCESS";
                        objValidateCustomer.AddressBook = new List<SFA_AddressBook>();

                        string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_address'>
                            <attribute name='hil_addressid' />
                            <attribute name='hil_pincode' />
                            <attribute name='hil_area' />
                            <attribute name='hil_fulladdress' />
                            <order attribute='hil_fulladdress' descending='false' />
                            <filter type='and'>
                                <condition attribute='hil_customer' operator='eq' value='{" + entcoll.Entities[0].Id + @"}' />
                            </filter>
                            <link-entity name='hil_area' from='hil_areaid' to='hil_area' visible='false' link-type='outer' alias='area'>
                                <attribute name='hil_areacode' />
                                <attribute name='hil_name' />
                            </link-entity>
                        </entity>
                        </fetch>";

                        entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                        if (entcoll.Entities.Count > 0)
                        {
                            foreach (Entity ent in entcoll.Entities)
                            {
                                objAddress = new SFA_AddressBook()
                                {
                                    AddressGuid = ent.GetAttributeValue<Guid>("hil_addressid"),
                                    Address = ent.GetAttributeValue<string>("hil_fulladdress"),

                                    PINCode = ent.GetAttributeValue<EntityReference>("hil_pincode").Name,

                                    AreaCode = ent.GetAttributeValue<AliasedValue>("area.hil_areacode").Value.ToString(),

                                    AreaName = ent.GetAttributeValue<AliasedValue>("area.hil_name").Value.ToString()
                                };
                                objValidateCustomer.AddressBook.Add(objAddress);
                            }
                        }
                        return objValidateCustomer;
                    }
                }
                else {
                    objValidateCustomer = new SFA_ValidateCustomer { ResultStatus = false, ResultMessage = "D365 Service Unavailable", AddressBook = new List<SFA_AddressBook>() };
                    return objValidateCustomer;
                }
            }
            catch (Exception ex)
            {
                objValidateCustomer = new SFA_ValidateCustomer { ResultStatus = false, ResultMessage = "D365 Internal Server Error : " + ex.Message, AddressBook = new List<SFA_AddressBook>() };
                return objValidateCustomer;
            }
        }
    }

    [DataContract]
    public class SFA_AddressBook
    {
        [DataMember]
        public string Address { get; set; }
        [DataMember]
        public string PINCode { get; set; }
        [DataMember]
        public string AreaCode { get; set; }
        [DataMember]
        public string AreaName { get; set; }
        [DataMember]
        public Guid AddressGuid { get; set; }
    }
}
