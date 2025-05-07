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
using Microsoft.Xrm.Sdk.Deployment;
using System.ServiceModel.Description;
using System.Configuration;
using System.Threading.Tasks;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class ViewAddress
    {
        [DataMember]
        public string ADDRESS_ID { get; set; }
        public AddressFields GetThisAddress(ViewAddress iViewAdd)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            AddressFields iAddFields = new AddressFields();
            hil_address ThisAddress = (hil_address)service.Retrieve(hil_address.EntityLogicalName, new Guid(iViewAdd.ADDRESS_ID), new ColumnSet("hil_street1", "hil_street2", "hil_street3", "hil_city", "hil_district", "hil_state", "hil_pincode", "hil_customer", "statecode", "hil_addresstype"));
            if(ThisAddress.hil_Street1 != null)
            {
                iAddFields.STREET1 = ThisAddress.hil_Street1;
            }
            if(ThisAddress.hil_Street2 != null)
            {
                iAddFields.STREET2 = ThisAddress.hil_Street2;
            }
            if(ThisAddress.hil_Street3 != null)
            {
                iAddFields.STREET3 = ThisAddress.hil_Street3;
            }
            if(ThisAddress.hil_CIty != null)
            {
                iAddFields.CITY_ID = ThisAddress.hil_CIty.Id.ToString();
                iAddFields.CITY_NAME = ThisAddress.hil_CIty.Name.ToUpper();
            }
            if (ThisAddress.hil_State != null)
            {
                iAddFields.STATE_ID = ThisAddress.hil_State.Id.ToString();
                iAddFields.STATE_NAME = ThisAddress.hil_State.Name.ToUpper();
            }
            if (ThisAddress.hil_District != null)
            {
                iAddFields.DISTRICT_ID = ThisAddress.hil_District.Id.ToString();
                iAddFields.DISTRICT_NAME = ThisAddress.hil_District.Name.ToUpper();
            }
            if (ThisAddress.hil_PinCode != null)
            {
                iAddFields.PINCODE_ID = ThisAddress.hil_PinCode.Id.ToString();
                iAddFields.PINCODE_NAME = ThisAddress.hil_PinCode.Name.ToUpper();
            }
            if (ThisAddress.hil_AddressType != null)
            {
                iAddFields.ADD_TYPE = ThisAddress.hil_AddressType.Value.ToString();
            }
            return iAddFields;
        }
    }
    public class AddressFields
    {
        [DataMember(IsRequired = false)]
        public string STREET1 { get; set; }
        [DataMember(IsRequired = false)]
        public string STREET2 { get; set; }
        [DataMember(IsRequired = false)]
        public string STREET3 { get; set; }
        [DataMember(IsRequired = false)]
        public string CITY_NAME { get; set; }
        [DataMember(IsRequired = false)]
        public string CITY_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string STATE_NAME { get; set; }
        [DataMember(IsRequired = false)]
        public string STATE_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string DISTRICT_NAME { get; set; }
        [DataMember(IsRequired = false)]
        public string DISTRICT_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string PINCODE_NAME { get; set; }
        [DataMember(IsRequired = false)]
        public string PINCODE_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string ADD_TYPE { get; set; }
    }
}
