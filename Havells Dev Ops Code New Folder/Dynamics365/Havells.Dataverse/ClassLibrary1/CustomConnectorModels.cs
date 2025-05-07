using System;

namespace Havells.Dataverse.CustomConnector.Models
{
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
    public class IoTAddressBookResultV1
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
        public int AddressTypeEnum { get; set; } = 3;
        public bool IsDefault { get; set; }
        public int StatusCode { get; set; }
        public string StatusDescription { get; set; }
    }
    public class IoTAddressBookV1
    {
        public Guid CustomerGuid { get; set; }
        public string Pincode { get; set; }
        public string MobileNumber { get; set; }
        public bool IsDefault { get; set; }

    }
}
