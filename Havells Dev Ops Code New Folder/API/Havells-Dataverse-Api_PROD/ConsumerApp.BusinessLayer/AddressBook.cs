using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class AddressBook
    {
        [DataMember]
        public string CustomerMobleNo { get; set; }
        [DataMember]
        public string FullAddress { get; set; }
        [DataMember]
        public string PINCode { get; set; }
        [DataMember]
        public Guid Guid { get; set; }
        [DataMember]
        public bool ResultStatus { get; set; }
        [DataMember]
        public string ResultMessage { get; set; }

        public List<AddressBook> GetAddresses(AddressBook AddressBook)
        {
            Guid customerGuid = Guid.Empty;
            AddressBook objAddressBook;
            List<AddressBook> lstAddressBook = new List<AddressBook>();
            EntityCollection entcoll;
            QueryExpression Query;
            try
            {
                IOrganizationService service = ConnectToCRM.GetOrgService();
                if (AddressBook.CustomerMobleNo.Trim().Length == 0)
                {
                    objAddressBook = new AddressBook { ResultStatus = false, ResultMessage = "Customer Mobile No. is required." };
                    lstAddressBook.Add(objAddressBook);
                    return lstAddressBook;
                }

                Query = new QueryExpression("contact");
                Query.ColumnSet = new ColumnSet("fullname");
                Query.Criteria = new FilterExpression(LogicalOperator.And);
                Query.Criteria.AddCondition("mobilephone", ConditionOperator.Equal, AddressBook.CustomerMobleNo);
                entcoll = service.RetrieveMultiple(Query);

                if (entcoll.Entities.Count == 0)
                {
                    objAddressBook = new AddressBook { ResultStatus = false, ResultMessage = "Customer Mobile No. does not exist." };
                    lstAddressBook.Add(objAddressBook);
                    return lstAddressBook;
                }
                else
                {
                    customerGuid = entcoll.Entities[0].Id;

                    string fetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='hil_address'>
                            <attribute name='hil_addressid' />
                            <attribute name='hil_pincode' />
                            <attribute name='hil_fulladdress' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                                <condition attribute='hil_customer' operator='eq' value='{" + customerGuid.ToString() + @"}' />
                            </filter>
                        </entity>
                        </fetch>";

                    entcoll = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    if (entcoll.Entities.Count == 0)
                    {
                        objAddressBook = new AddressBook { ResultStatus = false, ResultMessage = "No Addrerss Book found for given Mobile No."};
                        lstAddressBook.Add(objAddressBook);
                    }
                    else {
                        foreach (Entity ent in entcoll.Entities) {
                            objAddressBook = new AddressBook();
                            objAddressBook.Guid = ent.GetAttributeValue<Guid>("hil_addressid");
                            objAddressBook.FullAddress = ent.GetAttributeValue<string>("hil_fulladdress");
                            objAddressBook.PINCode = ent.GetAttributeValue<EntityReference>("hil_pincode").Name;
                            objAddressBook.CustomerMobleNo = AddressBook.CustomerMobleNo;
                            objAddressBook.ResultStatus = true;
                            objAddressBook.ResultMessage = "Success";
                            lstAddressBook.Add(objAddressBook);
                        }
                    }
                    return lstAddressBook;
                }
            }
            catch (Exception ex)
            {
                objAddressBook = new AddressBook { ResultStatus = false, ResultMessage = ex.Message };
                lstAddressBook.Add(objAddressBook);
                return lstAddressBook;
            }
        }
    }
}
