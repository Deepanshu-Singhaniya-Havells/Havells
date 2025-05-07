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

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class ConsumerAddress
    {
        [DataMember]
        public string CONT_ID { get; set; }
        [DataMember]
        public string ADD_TYPE { get; set; }
        [DataMember]
        public string STREET1 { get; set; }
        [DataMember(IsRequired = false)]
        public string STREET2 { get; set; }
        [DataMember(IsRequired = false)]
        public string LANDMARK { get; set; }
        [DataMember]
        public string PIN { get; set; }
        [DataMember(IsRequired = false)]
        public string DISTRICT_NAME { get; set; }
        [DataMember(IsRequired = false)]
        public string ADD_ID { get; set; }//Output
        [DataMember(IsRequired = false)]
        public bool STATUS { get; set; }//Output
        [DataMember(IsRequired = false)]
        public string STATUS_DESC { get; set; }//Output
        public ConsumerAddress AddAddress(ConsumerAddress Addr)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();//get org service obj for connection
            hil_address iAdd = new hil_address();
            if (Addr.ADD_TYPE == "1")
            {
                iAdd.hil_AddressType = new OptionSetValue(1);
                iAdd.hil_name = "Permanent";
            }
            else if (Addr.ADD_TYPE == "2")
            {
                iAdd.hil_AddressType = new OptionSetValue(2);
                iAdd.hil_name = "Alternate";
            }
            iAdd.hil_Street1 = Addr.STREET1;
            if (Addr.STREET2 != null)
                iAdd.hil_Street2 = Addr.STREET2;
            if (Addr.LANDMARK != null)
                iAdd.hil_Street3 = Addr.LANDMARK;
            iAdd.hil_Customer = new EntityReference(Contact.EntityLogicalName, new Guid(Addr.CONT_ID));
            QueryExpression Query = new QueryExpression(hil_pincode.EntityLogicalName);
            Query.ColumnSet = new ColumnSet(true);
            Query.Criteria = new FilterExpression(LogicalOperator.And);
            Query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Addr.PIN);
            EntityCollection Found = service.RetrieveMultiple(Query);
            if (Found.Entities.Count > 0)
            {
                hil_pincode iPin = Found.Entities[0].ToEntity<hil_pincode>();
                iAdd.hil_PinCode = new EntityReference(hil_pincode.EntityLogicalName, iPin.Id);
                if (iPin.hil_City != null)
                    iAdd.hil_CIty = iPin.hil_City;
                if (iPin.hil_State != null)
                    iAdd.hil_State = iPin.hil_State;
                if (iPin.hil_area != null)
                    iAdd.hil_Area = iPin.hil_area;
                Guid AddressId = new Guid(); 
                //= service.Create(iAdd);
                Addr.ADD_ID = AddressId.ToString();
                Addr.STATUS = true;
                Addr.STATUS_DESC = "SUCCESS";
            }
            else
            {
                Addr.ADD_ID = "";
                Addr.STATUS = false;
                Addr.STATUS_DESC = "FAILURE : PINCODE NOT FOUND";
            }
            return Addr;
        }
    }
}
