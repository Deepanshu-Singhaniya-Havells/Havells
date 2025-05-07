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
    public class AddressDelete
    {
        [DataMember]
        public string ADDRESS_ID { get; set; }
        public OutputAddressDelete DeleteThisAddress(AddressDelete IDelete)
        {
            OutputAddressDelete clsOut = new OutputAddressDelete();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            try
            {
                service.Delete(hil_address.EntityLogicalName, new Guid(IDelete.ADDRESS_ID));
                clsOut.STATUS = true;
                clsOut.DESC = "SUCCESS";
            }
            catch
            {
                clsOut.STATUS = false;
                clsOut.DESC = "FAILURE";
            }
            return clsOut;
        }
    }
    public class OutputAddressDelete
    {
        [DataMember(IsRequired = false)]
        public bool STATUS { get; set; }
        [DataMember(IsRequired = false)]
        public string DESC { get; set; }
    }
}
