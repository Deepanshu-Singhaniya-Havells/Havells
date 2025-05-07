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
    public class PinCodeMaster
    {
        [DataMember]
        public string PinCode { get; set; }
        [DataMember(IsRequired = false)]
        public string PinCodesName { get; set; }
        [DataMember(IsRequired = false)]
        public string PinCodeGuId { get; set; }
        [DataMember(IsRequired = false)]
        public string StateName { get; set; }
        [DataMember(IsRequired = false)]
        public string StateGuId { get; set; }
        [DataMember(IsRequired = false)]
        public string CityName { get; set; }
        [DataMember(IsRequired = false)]
        public string CityGuId { get; set; }
        [DataMember(IsRequired = false)]
        public bool IsValid { get; set; }
        [DataMember(IsRequired = false)]
        public string DISTRICT_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string DISTRICT_NAME { get; set; }
        public List<PinCodeMaster> GetAllActivePinCodes(PinCodeMaster Pin)
        {
            List<PinCodeMaster> obj = new List<PinCodeMaster>();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            QueryExpression Qry = new QueryExpression(hil_pincode.EntityLogicalName);
            //Qry.EntityName = hil_pincode.EntityLogicalName;
            ColumnSet Col = new ColumnSet("hil_name", "hil_pincodeid", "hil_city", "hil_state");
            Qry.ColumnSet = Col;
            Qry.Criteria = new FilterExpression(LogicalOperator.And);
            Qry.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
            Qry.Criteria.AddCondition(new ConditionExpression("hil_name", ConditionOperator.Equal, Pin.PinCode));
            EntityCollection Colec = service.RetrieveMultiple(Qry);
            if(Colec.Entities.Count >= 1)
            {
                hil_pincode iPin = Colec.Entities[0].ToEntity<hil_pincode>();
                EntityReference State = new EntityReference();
                EntityReference City = new EntityReference();
                QueryExpression Qry1 = new QueryExpression(hil_businessmapping.EntityLogicalName);
                ColumnSet Col1 = new ColumnSet("hil_name", "hil_pincode", "hil_city", "hil_state", "hil_district");
                Qry1.ColumnSet = Col1;
                Qry1.Criteria = new FilterExpression(LogicalOperator.And);
                Qry1.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                Qry1.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, iPin.Id));
                EntityCollection Colec1 = service.RetrieveMultiple(Qry1);
                if (Colec1.Entities.Count > 0)
                {
                    hil_businessmapping iBus = Colec1.Entities[0].ToEntity<hil_businessmapping>();
                    obj.Add(
                    new PinCodeMaster
                    {
                        StateGuId = Convert.ToString(iBus.hil_state.Id),
                        StateName = iBus.hil_state.Name,
                        CityGuId = Convert.ToString(iBus.hil_city.Id),
                        CityName = iBus.hil_city.Name,
                        PinCodeGuId = Convert.ToString(iBus.hil_pincode.Id),
                        PinCodesName = Convert.ToString(iBus.hil_pincode.Name),
                        DISTRICT_ID = Convert.ToString(iBus.hil_district.Id),
                        DISTRICT_NAME = Convert.ToString(iBus.hil_district.Name),
                        IsValid = true
                    });
                }
                    //if (et.Attributes.Contains("hil_state"))
                    //{
                    //    State = (EntityReference)et["hil_state"];
                    //}
                    //if(et.Attributes.Contains("hil_city"))
                    //{
                    //    City = (EntityReference)et["hil_city"];
                    //}
                    //obj.Add(
                    //new PinCodeMaster
                    //{
                    //    StateGuId = Convert.ToString(State.Id),
                    //    StateName = State.Name,
                    //    CityGuId = Convert.ToString(City.Id),
                    //    CityName = City.Name,
                    //    PinCodeGuId = Convert.ToString(et["hil_pincodeid"]),
                    //    PinCodesName = Convert.ToString(et["hil_name"]),
                    //    IsValid = true
                    //});
            }
            else
            {
                obj.Add(
                    new PinCodeMaster
                    {
                        StateGuId = "",
                        StateName = "",
                        CityGuId = "",
                        CityName = "",
                        PinCodeGuId = "",
                        PinCodesName = "",
                        IsValid = false
                    });
            }
            return (obj);
        }
    }
}