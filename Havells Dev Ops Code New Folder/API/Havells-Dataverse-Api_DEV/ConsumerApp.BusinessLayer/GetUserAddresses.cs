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
    public class GetUserAddresses
    {
        [DataMember]
        public string CONT_ID { get; set; }
        public List<OutputAddressClass> GetCustomerInformation(GetUserAddresses iConsumer)
        {
            List<OutputAddressClass> obj = new List<OutputAddressClass>();
            try
            {
                string ADD_CONCAT = string.Empty;
                string Address_Guid = string.Empty;
                
                IOrganizationService service = ConnectToCRM.GetOrgService();
                QueryExpression Qry = new QueryExpression(hil_address.EntityLogicalName);
                Qry.ColumnSet = new ColumnSet("hil_street1", "hil_street2", "hil_street3", "hil_city", "hil_district", "hil_state", "hil_pincode", "hil_customer", "statecode", "hil_addresstype");
                Qry.Criteria = new FilterExpression(LogicalOperator.And);
                Qry.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                Qry.Criteria.AddCondition(new ConditionExpression("hil_customer", ConditionOperator.Equal, new Guid(iConsumer.CONT_ID)));
                Qry.Criteria.AddCondition(new ConditionExpression("hil_addresstype", ConditionOperator.NotNull));
                EntityCollection Colec = service.RetrieveMultiple(Qry);
                foreach (hil_address et in Colec.Entities)
                {
                    try
                    {
                        Address_Guid = et.Id.ToString();
                        if (et.hil_Street1 != null)
                        {
                            ADD_CONCAT = ADD_CONCAT + et.hil_Street1.ToUpper() + ", ";
                        }
                        if (et.hil_Street2 != null)
                        {
                            ADD_CONCAT = ADD_CONCAT + et.hil_Street2.ToUpper() + ", ";
                        }
                        if (et.hil_Street3 != null)
                        {
                            ADD_CONCAT = ADD_CONCAT + et.hil_Street3.ToUpper() + ", ";
                        }
                        if (et.hil_CIty != null)
                        {
                            ADD_CONCAT = ADD_CONCAT + et.hil_CIty.Name.ToUpper() + ", ";
                        }
                        if (et.hil_District != null)
                        {
                            ADD_CONCAT = ADD_CONCAT + et.hil_District.Name.ToUpper() + ", ";
                        }
                        if (et.hil_State != null)
                        {
                            ADD_CONCAT = ADD_CONCAT + et.hil_State.Name.ToUpper() + " - ";
                        }
                        if (et.hil_PinCode != null)
                        {
                            ADD_CONCAT = ADD_CONCAT + et.hil_PinCode.Name;
                        }
                        obj.Add(
                        new OutputAddressClass
                        {
                            FULL_ADDRESS = ADD_CONCAT,
                            ADD_TYPE = et.hil_AddressType.Value.ToString(),
                            ADDRESS_ID = et.Id.ToString()
                        });
                        ADD_CONCAT = null;
                    }
                    catch(Exception ex)
                    {
                        continue;
                    }
                }
            }
            catch(Exception ex)
            {
                obj.Add(
                new OutputAddressClass
                {
                    FULL_ADDRESS = ex.Message.ToUpper(),
                    ADD_TYPE = ex.StackTrace,
                    ADDRESS_ID = ""
                });
            }
            
            return (obj);
        }
    }
    public class OutputAddressClass
    {
        [DataMember(IsRequired = false)]
        public string ADD_TYPE { get; set; }
        [DataMember(IsRequired = false)]
        public string FULL_ADDRESS { get; set; }
        [DataMember(IsRequired = false)]
        public string ADDRESS_ID { get; set; }
    }
}