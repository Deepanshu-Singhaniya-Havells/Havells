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
    public class AddAddress
    {
        [DataMember]
        public string CUSTOMER_ID { get; set; }
        [DataMember]
        public string STREET1 { get; set; }
        [DataMember(IsRequired = false)]
        public string STREET2 { get; set; }
        [DataMember(IsRequired = false)]
        public string STREET3 { get; set; }
        [DataMember]
        public string CITY_ID { get; set; }
        [DataMember]
        public string STATE_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string DISTRICT_ID { get; set; }
        [DataMember]
        public string PINCODE_ID { get; set; }
        [DataMember(IsRequired = false)]
        public Int32 ADDRESS_TYPE { get; set; } // 1- Permanent, 2 - Alternate
        public OutputAddAddress AddCustomerAddress(AddAddress iAdd)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            OutputAddAddress iStatus = new OutputAddAddress();
            try
            {
                hil_address eCreate = new hil_address();
                if(iAdd.CUSTOMER_ID != null)
                    eCreate.hil_Customer = new EntityReference(Contact.EntityLogicalName, new Guid(iAdd.CUSTOMER_ID));
                if (iAdd.STREET1 != null)
                    eCreate.hil_Street1 = iAdd.STREET1;
                if (iAdd.STREET2 != null)
                    eCreate.hil_Street2 = iAdd.STREET2;
                if (iAdd.STREET3 != null)
                    eCreate.hil_Street3 = iAdd.STREET3;
                if (iAdd.PINCODE_ID != null)
                {
                    eCreate.hil_PinCode = new EntityReference(hil_pincode.EntityLogicalName, new Guid(iAdd.PINCODE_ID));
                    hil_businessmapping iDetails = GetBusinessGeo(service, new Guid(iAdd.PINCODE_ID));
                    if(iDetails.Id != Guid.Empty)
                    {
                        eCreate.hil_SalesOffice = iDetails.hil_salesoffice;
                        eCreate.hil_Area = iDetails.hil_area;
                        eCreate.hil_Branch = iDetails.hil_branch;
                        eCreate.hil_Region = iDetails.hil_region;
                    }
                }
                if (iAdd.STATE_ID != null)
                    eCreate.hil_State = new EntityReference(hil_state.EntityLogicalName, new Guid(iAdd.STATE_ID));
                if (iAdd.CITY_ID != null)
                    eCreate.hil_CIty = new EntityReference(hil_city.EntityLogicalName, new Guid(iAdd.CITY_ID));
                if (iAdd.DISTRICT_ID != null)
                    eCreate.hil_District = new EntityReference(hil_district.EntityLogicalName, new Guid(iAdd.DISTRICT_ID));
                if (iAdd.ADDRESS_TYPE > 0)
                {
                    eCreate.hil_AddressType = new OptionSetValue(iAdd.ADDRESS_TYPE);
                    if (iAdd.ADDRESS_TYPE == 1)
                    {
                        eCreate.hil_name = "Permanent";
                    }
                    else if (iAdd.ADDRESS_TYPE == 2)
                    {
                        eCreate.hil_name = "Alternate";
                    }
                }
                Guid AddressGuid = service.Create(eCreate);
                if (AddressGuid != Guid.Empty)
                {
                    iStatus.STATUS = "SUCCESS";
                    iStatus.ADDRESS_ID = AddressGuid.ToString();
                }
                else
                {
                    iStatus.STATUS = "FAILURE";
                    iStatus.ADDRESS_ID = AddressGuid.ToString();
                }
            }
            catch(Exception ex)
            {
                iStatus.STATUS = "FAILURE";
                iStatus.ADDRESS_ID = ex.Message.ToUpper();
            }
            return iStatus;
        }
        public static hil_businessmapping GetBusinessGeo(IOrganizationService service, Guid PIN_ID)
        {
            hil_businessmapping iDetails = new hil_businessmapping();
            QueryExpression Qry = new QueryExpression(hil_businessmapping.EntityLogicalName);
            Qry.ColumnSet = new ColumnSet("hil_salesoffice", "hil_area", "hil_branch", "hil_region", "hil_pincode");
            Qry.Criteria = new FilterExpression(LogicalOperator.And);
            Qry.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            Qry.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, PIN_ID));
            Qry.Distinct = true;
            EntityCollection Colec = service.RetrieveMultiple(Qry);
            if (Colec.Entities.Count > 0)
            {
                iDetails = Colec.Entities[0].ToEntity<hil_businessmapping>();
            }
            else
            {
                iDetails.Id = Guid.Empty;
            }
            return iDetails;
        }
    }
    public class OutputAddAddress
    {
        [DataMember(IsRequired = false)]
        public string STATUS { get; set; }
        [DataMember(IsRequired = false)]
        public string ADDRESS_ID { get; set; }
    }
}
