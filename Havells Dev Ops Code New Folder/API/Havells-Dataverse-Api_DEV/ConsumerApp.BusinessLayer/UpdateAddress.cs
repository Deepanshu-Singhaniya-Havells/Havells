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
    public class UpdateAddress
    {
        [DataMember]
        public string ADDRESS_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string STREET1 { get; set; }
        [DataMember(IsRequired = false)]
        public string STREET2 { get; set; }
        [DataMember(IsRequired = false)]
        public string STREET3 { get; set; }
        [DataMember(IsRequired = false)]
        public string CITY_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string STATE_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string DISTRICT_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string PINCODE_ID { get; set; }
        public Int32 ADDRESS_TYPE { get; set; } // 1- Permanent, 2 - Alternate
        public OutputUpdateAddress UpdateThisAddress(UpdateAddress iUpdate)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            OutputUpdateAddress iOut = new OutputUpdateAddress();
            try
            {
                hil_address iUpdateAddress = new hil_address();
                iUpdateAddress.Id = new Guid(iUpdate.ADDRESS_ID);
                if (iUpdate.STREET1 != null)
                    iUpdateAddress.hil_Street1 = iUpdate.STREET1;
                if (iUpdate.STREET2 != null)
                    iUpdateAddress.hil_Street2 = iUpdate.STREET2;
                if (iUpdate.STREET3 != null)
                    iUpdateAddress.hil_Street3 = iUpdate.STREET3;
                if (iUpdate.PINCODE_ID != null)
                {
                    iUpdateAddress.hil_PinCode = new EntityReference(hil_pincode.EntityLogicalName, new Guid(iUpdate.PINCODE_ID));
                    hil_businessmapping iDetails = ConsumerApp.BusinessLayer.AddAddress.GetBusinessGeo(service, new Guid(iUpdate.PINCODE_ID));
                    if (iDetails.Id != Guid.Empty)
                    {
                        iUpdateAddress.hil_SalesOffice = iDetails.hil_salesoffice;
                        iUpdateAddress.hil_Area = iDetails.hil_area;
                        iUpdateAddress.hil_Branch = iDetails.hil_branch;
                        iUpdateAddress.hil_Region = iDetails.hil_region;
                    }
                }
                if (iUpdate.STATE_ID != null)
                    iUpdateAddress.hil_State = new EntityReference(hil_state.EntityLogicalName, new Guid(iUpdate.STATE_ID));
                if (iUpdate.CITY_ID != null)
                    iUpdateAddress.hil_CIty = new EntityReference(hil_city.EntityLogicalName, new Guid(iUpdate.CITY_ID));
                if (iUpdate.DISTRICT_ID != null)
                    iUpdateAddress.hil_District = new EntityReference(hil_district.EntityLogicalName, new Guid(iUpdate.DISTRICT_ID));
                if (iUpdate.ADDRESS_TYPE > 0)
                {
                    iUpdateAddress.hil_AddressType = new OptionSetValue(iUpdate.ADDRESS_TYPE);
                    if (iUpdate.ADDRESS_TYPE == 1)
                    {
                        iUpdateAddress.hil_name = "Permanent";
                    }
                    else if (iUpdate.ADDRESS_TYPE == 2)
                    {
                        iUpdateAddress.hil_name = "Alternate";
                    }
                }
                else
                {
                    iUpdateAddress.hil_AddressType = new OptionSetValue(2);
                    iUpdateAddress.hil_name = "Alternate";
                }
                service.Update(iUpdateAddress);
                iOut.STATUS = "SUCCESS";
                iOut.MESSAGE = "SUCCESS";
            }
            catch(Exception ex)
            {
                iOut.STATUS = "FAILURE";
                iOut.MESSAGE = ex.Message.ToUpper();
            }
            return iOut;
        }
    }
    public class OutputUpdateAddress
    {
        [DataMember(IsRequired = false)]
        public string STATUS { get; set; }
        [DataMember(IsRequired = false)]
        public string MESSAGE { get; set; }
    }
}
