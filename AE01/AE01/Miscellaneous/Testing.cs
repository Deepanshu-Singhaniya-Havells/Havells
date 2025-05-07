using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.Runtime.Caching;
using System.Text.RegularExpressions;

namespace AE01.Miscellaneous
{
    internal class Testing
    {
        private IOrganizationService service;


        
        public void ChangeConsumerName()
        {


            //if (!entity.Contains("mobilephone") || !entity.Contains("firstname") || !entity.Contains("hil_consumersource"))
            //{
            //    throw new InvalidPluginExecutionException("Consumer Name, Mobile Number and Source of Creation is required.");
            //}
            //if (_Cnt.MobilePhone != null)
            //{
            //    if (_Cnt.MobilePhone.Length != 10)
            //    {
            //        throw new InvalidPluginExecutionException("MOBILE NUMBER MUST BE 10 DIGIT");
            //    }
            //}
            //else
            //{
            //    throw new InvalidPluginExecutionException("MOBILE NUMBER CAN'T BE NULL");
            //}

            Entity entity = new Entity();
            entity.Attributes.Add("firstname", "Azad123");
            entity.Attributes.Add("middlename", "$");
            entity.Attributes.Add("lastname", "Kumar."); 

            string namePattern = @"^[a-zA-Z_. ]*$";
            Regex regex = new Regex(namePattern);
            string nameStr = string.Empty;

            if (entity.Contains("firstname"))
            {
                nameStr = entity.GetAttributeValue<string>("firstname");
                if (!regex.IsMatch(nameStr))
                {
                    throw new InvalidPluginExecutionException("First Name contains invalid characters. Only Letters A-z,a-z, and dot(.) are allowed.");
                }
            }
            if (entity.Contains("middlename"))
            {
                nameStr = entity.GetAttributeValue<string>("middlename");
                if (!regex.IsMatch(nameStr))
                {
                    throw new InvalidPluginExecutionException("Middle Name contains invalid characters. Only Letters A-z,a-z, and dot(.) are allowed.");
                }
            }
            if (entity.Contains("lastname"))
            {
                nameStr = entity.GetAttributeValue<string>("lastname");
                if (!regex.IsMatch(nameStr))
                {
                    throw new InvalidPluginExecutionException("Last Name contains invalid characters. Only Letters A-z,a-z, and dot(.) are allowed.");
                }
            }


            Entity consumer = service.Retrieve("contact", new Guid("def54344-ebcd-ed11-a7c6-6045bdac5348"), new ColumnSet(true));
            consumer["firstname"] = "Deepanshu^";
            service.Update(consumer); 





        }

        public void ValidateAttachmentCount()
        {

            string workDone = "2927FA6C-FA0F-E911-A94E-000D3AF060A1";
            Guid jobId = new Guid("5d4fda8d-eb78-ef11-ac21-6045bdad7f34");
            string fetchPermittedCount = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='msdyn_workorder'>
                                            <attribute name='msdyn_customerasset' />
                                            <attribute name='msdyn_workorderid' />
                                            <order attribute='msdyn_name' descending='false' />
                                            <filter type='and'>
                                            <condition attribute='hil_sourceofjob' operator='eq' value='2' />
                                            <condition attribute='msdyn_substatus' operator='eq' value='{workDone}' />
                                            <condition attribute='msdyn_workorderid' operator='eq' value='{jobId}' />
                                            </filter>
                                            <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='msdyn_customerasset' link-type='inner' alias='am'>
                                            <filter type='and'>
                                            <condition attribute='createdon' operator='on-or-after' value='2024-09-26' />
                                            </filter>
                                            <link-entity name='product' from='productid' to='hil_productcategory' link-type='inner' alias='an'>
                                            <attribute name='hil_assetmediacount' />
                                            <filter type='and'>
                                            <condition attribute='hil_assetmediacount' operator='gt' value='0' />
                                            </filter>
                                            </link-entity>
                                            </link-entity>
                                            </entity>
                                            </fetch>";

            EntityCollection permittedCountColl = service.RetrieveMultiple(new FetchExpression(fetchPermittedCount));
            if (permittedCountColl.Entities.Count > 0)
            {
                int permittedCount = (int)permittedCountColl.Entities[0].GetAttributeValue<AliasedValue>("an.hil_assetmediacount").Value;

                EntityReference _entCustomerAsser = permittedCountColl.Entities[0].GetAttributeValue<EntityReference>("msdyn_customerasset");

                string fetchactualCount  = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='annotation'>
                                            <attribute name='subject' />
                                            <attribute name='annotationid' />
                                            <order attribute='subject' descending='false' />
                                            <filter type='and'>
                                            <filter type='or'>
                                            <condition attribute='notetext' operator='like' value='%https://d365storagesa.blob.core.windows.net%'/>
                                            <condition attribute='isdocument' operator='eq' value='1' />
                                            </filter>
                                            </filter>
                                            <link-entity name='msdyn_customerasset' from='msdyn_customerassetid' to='objectid' link-type='inner' alias='as'>
                                            <filter type='and'>
                                            <condition attribute='createdon' operator='on-or-after' value='2024-09-26' />
                                            <condition attribute='msdyn_customerassetid' operator='eq' value='{_entCustomerAsser.Id}' />
                                            </filter>
                                            </link-entity>
                                            </entity>
                                            </fetch>";
                
                int actualCount = service.RetrieveMultiple(new FetchExpression(fetchactualCount)).Entities.Count;

                if(actualCount < permittedCount)
                {
                    Console.WriteLine("Attachment count is less than permitted count");
                }

            }

        }
       
        public List<IoTAddressBookResult> GetIoTAddressBook(IoTAddressBook address)
        {
            List<IoTAddressBookResult> addressList = new List<IoTAddressBookResult>();
            IoTAddressBookResult objAddress;
            QueryExpression query;
            EntityCollection collection;
            EntityReference erPincode = null;
            try
            {
                if (address.CustomerGuid == Guid.Empty && address.MobileNumber == null)
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Customer GUID/Mobile Number is required." };
                    addressList.Add(objAddress);
                    return addressList;
                }

                if (service != null)
                {
                    if (address.Pincode != null)
                    {
                        if (address.Pincode.Trim().Length == 6)
                        {
                            query = new QueryExpression("hil_pincode");
                            query.ColumnSet = new ColumnSet(false);
                            query.Criteria = new FilterExpression(LogicalOperator.And);
                            query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, address.Pincode);
                            query.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                            EntityCollection entcoll = service.RetrieveMultiple(query);
                            if (entcoll.Entities.Count > 0)
                            {
                                erPincode = entcoll.Entities[0].ToEntityReference();
                            }
                            else
                            {
                                objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Invalid Pincode." };
                                addressList.Add(objAddress);
                                return addressList;
                            }
                        }
                        else
                        {
                            objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "Invalid Pincode." };
                            addressList.Add(objAddress);
                            return addressList;
                        }
                    }

                    query = new QueryExpression("hil_address");
                    query.ColumnSet = new ColumnSet("hil_customer", "hil_addressid", "hil_street3", "hil_street2", "hil_street1", "hil_fulladdress", "hil_pincode", "hil_area", "hil_businessgeo", "hil_addresstype", "hil_city", "hil_state");
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    if (address.CustomerGuid != Guid.Empty)
                    {
                        query.Criteria.AddCondition("hil_customer", ConditionOperator.Equal, address.CustomerGuid);
                    }
                    if (erPincode != null)
                    {
                        query.Criteria.AddCondition("hil_pincode", ConditionOperator.Equal, erPincode.Id);
                    }
                    if (address.MobileNumber != null)
                    {
                        LinkEntity lnk = new LinkEntity("hil_address", "contact", "hil_customer", "contactid", JoinOperator.Inner);
                        lnk.LinkCriteria.AddCondition("mobilephone", ConditionOperator.Equal, address.MobileNumber);
                        query.LinkEntities.Add(lnk);
                    }
                    collection = service.RetrieveMultiple(query);

                    if (collection.Entities.Count == 0)
                    {
                        objAddress = new IoTAddressBookResult { StatusCode = "204", StatusDescription = "No Address Book found." };
                        addressList.Add(objAddress);
                        return addressList;
                    }

                    //query = new QueryExpression()
                    //{
                    //    EntityName = hil_address.EntityLogicalName,
                    //    ColumnSet = new ColumnSet("hil_addressid", "hil_street3", "hil_street2", "hil_street1", "hil_fulladdress", "hil_pincode", "hil_area", "hil_businessgeo", "hil_addresstype")
                    //};
                    //FilterExpression filterExpression = new FilterExpression(LogicalOperator.And);
                    //filterExpression.Conditions.Add(new ConditionExpression("hil_customer", ConditionOperator.Equal, address.CustomerGuid));
                    //query.Criteria.AddFilter(filterExpression);

                    //collection = service.RetrieveMultiple(query);

                    if (collection.Entities != null && collection.Entities.Count > 0)
                    {
                        foreach (Entity item in collection.Entities)
                        {
                            objAddress = new IoTAddressBookResult();
                            if (item.Attributes.Contains("hil_street1"))
                            {
                                objAddress.AddressLine1 = item.GetAttributeValue<string>("hil_street1");
                            }
                            if (item.Attributes.Contains("hil_street2"))
                            {
                                objAddress.AddressLine2 = item.GetAttributeValue<string>("hil_street2");
                            }
                            if (item.Attributes.Contains("hil_street3"))
                            {
                                objAddress.AddressLine3 = item.GetAttributeValue<string>("hil_street3");
                            }
                            if (item.Attributes.Contains("hil_fulladdress"))
                            {
                                objAddress.FullAddress = item.GetAttributeValue<string>("hil_fulladdress");
                            }
                            if (item.Attributes.Contains("hil_addressid"))
                            {
                                objAddress.AddressGuid = item.Id;
                            }
                            if (item.Attributes.Contains("hil_businessgeo"))
                            {
                                objAddress.BizGeoGuid = item.GetAttributeValue<EntityReference>("hil_businessgeo").Id;
                            }
                            if (item.Attributes.Contains("hil_businessgeo"))
                            {
                                objAddress.BizGeoName = item.GetAttributeValue<EntityReference>("hil_businessgeo").Name;
                            }
                            if (item.Attributes.Contains("hil_pincode"))
                            {
                                objAddress.PINCodeGuid = item.GetAttributeValue<EntityReference>("hil_pincode").Id;
                            }
                            if (item.Attributes.Contains("hil_pincode"))
                            {
                                objAddress.PINCode = item.GetAttributeValue<EntityReference>("hil_pincode").Name;
                            }
                            if (item.Attributes.Contains("hil_area"))
                            {
                                objAddress.AreaGuid = item.GetAttributeValue<EntityReference>("hil_area").Id;
                            }
                            if (item.Attributes.Contains("hil_area"))
                            {
                                objAddress.Area = item.GetAttributeValue<EntityReference>("hil_area").Name;
                            }
                            if (item.Attributes.Contains("hil_addresstype"))
                            {
                                OptionSetValue osv = item.GetAttributeValue<OptionSetValue>("hil_addresstype");
                                objAddress.AddressType = osv.Value == 1 ? "Permanent" : "Alternate";
                            }
                            if (item.Attributes.Contains("hil_addresstype"))
                            {
                                objAddress.AddressTypeEnum = item.GetAttributeValue<OptionSetValue>("hil_addresstype").Value.ToString();
                            }
                            if (item.Attributes.Contains("hil_city"))
                            {
                                objAddress.CityName = item.GetAttributeValue<EntityReference>("hil_city").Name;
                            }
                            if (item.Attributes.Contains("hil_state"))
                            {
                                objAddress.StateName = item.GetAttributeValue<EntityReference>("hil_state").Name;
                            }
                            objAddress.CustomerGuid = address.CustomerGuid;
                            objAddress.MobileNumber = address.MobileNumber;
                            objAddress.StatusCode = "200";
                            objAddress.StatusDescription = "OK";
                            addressList.Add(objAddress);
                        }
                    }
                    return addressList;
                }
                else
                {
                    objAddress = new IoTAddressBookResult { StatusCode = "503", StatusDescription = "D365 Service Unavailable" };
                    addressList.Add(objAddress);
                }
            }
            catch (Exception ex)
            {
                objAddress = new IoTAddressBookResult { StatusCode = "500", StatusDescription = "D365 Internal Server Error : " + ex.Message.ToUpper() };
                addressList.Add(objAddress);
            }
            return addressList;
        }

        public List<IoTAddressBookResult> GetIoTAddressBookAsync(IoTAddressBook address)
        {
            Console.WriteLine("");

            // Check cache first
            var cacheKey = $"IoTAddressBook_{address.CustomerGuid}";
            var cachedResult = MemoryCache.Default.Get(cacheKey) as List<IoTAddressBookResult>;
            if (cachedResult != null)
            {
                return cachedResult;
            }

            // Asynchronous database query
            var query = new QueryExpression("address")
            {
                ColumnSet = new ColumnSet("CustomerGuid", "AddressGuid", "MobileNumber", "CustomerName", "EmailAddress", "AddressLine1", "AddressLine2", "AddressLine3", "AddressPhone", "BizGeoGuid", "BizGeoName", "PINCodeGuid", "PINCode", "Area", "AreaGuid", "FullAddress", "AddressType", "AddressTypeEnum", "StatusCode", "StatusDescription", "CityName", "StateName")
            };
            query.Criteria.AddCondition("CustomerGuid", ConditionOperator.Equal, address.CustomerGuid);

            var result = service.RetrieveMultiple(query);

            var addressBookResults = result.Entities.Select(e => new IoTAddressBookResult
            {
                CustomerGuid = e.GetAttributeValue<Guid>("CustomerGuid"),
                AddressGuid = e.GetAttributeValue<Guid>("AddressGuid"),
                MobileNumber = e.GetAttributeValue<string>("MobileNumber"),
                CustomerName = e.GetAttributeValue<string>("CustomerName"),
                EmailAddress = e.GetAttributeValue<string>("EmailAddress"),
                AddressLine1 = e.GetAttributeValue<string>("AddressLine1"),
                AddressLine2 = e.GetAttributeValue<string>("AddressLine2"),
                AddressLine3 = e.GetAttributeValue<string>("AddressLine3"),
                AddressPhone = e.GetAttributeValue<string>("AddressPhone"),
                BizGeoGuid = e.GetAttributeValue<Guid>("BizGeoGuid"),
                BizGeoName = e.GetAttributeValue<string>("BizGeoName"),
                PINCodeGuid = e.GetAttributeValue<Guid>("PINCodeGuid"),
                PINCode = e.GetAttributeValue<string>("PINCode"),
                Area = e.GetAttributeValue<string>("Area"),
                AreaGuid = e.GetAttributeValue<Guid>("AreaGuid"),
                FullAddress = e.GetAttributeValue<string>("FullAddress"),
                AddressType = e.GetAttributeValue<string>("AddressType"),
                AddressTypeEnum = e.GetAttributeValue<string>("AddressTypeEnum"),
                StatusCode = e.GetAttributeValue<string>("StatusCode"),
                StatusDescription = e.GetAttributeValue<string>("StatusDescription"),
                CityName = e.GetAttributeValue<string>("CityName"),
                StateName = e.GetAttributeValue<string>("StateName")
            }).ToList();

            // Store result in cache
            MemoryCache.Default.Add(cacheKey, addressBookResults, DateTimeOffset.UtcNow.AddMinutes(10));

            return addressBookResults;
        }

    }

    public class IoTAddressBookResult
    {
        public Guid CustomerGuid { get; set; }

        public Guid AddressGuid { get; set; }

        public string MobileNumber { get; set; }

        public string CustomerName { get; set; }

        public string EmailAddress { get; set; }

        public string AddressLine1 { get; set; }

        public string AddressLine2 { get; set; }

        public string AddressLine3 { get; set; }

        public string AddressPhone { get; set; }

        public Guid BizGeoGuid { get; set; }

        public string BizGeoName { get; set; }

        public Guid PINCodeGuid { get; set; }

        public string PINCode { get; set; }

        public string Area { get; set; }

        public Guid AreaGuid { get; set; }

        public string FullAddress { get; set; }

        public string AddressType { get; set; }

        public string AddressTypeEnum { get; set; }

        public string StatusCode { get; set; }

        public string StatusDescription { get; set; }
        public string CityName { get; set; }

        public string StateName { get; set; }
    }

    public class IoTAddressBook
    {
        public Guid CustomerGuid { get; set; }
        public string Pincode { get; set; }
        public string MobileNumber { get; set; }

    }
}
