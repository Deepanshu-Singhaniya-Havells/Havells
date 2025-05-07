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
    public class GeoLocator
    {
        [DataMember]
        public string PinCode { get; set; }
        [DataMember]
        public string State { get; set; }
        [DataMember]
        public string City { get; set; }
        [DataMember(IsRequired = false)]
        public string AccountName { get; set; }
        [DataMember(IsRequired = false)]
        public string Address { get; set; }
        [DataMember(IsRequired = false)]
        public string ACCOUNTNAME { get; set; }
        [DataMember(IsRequired = false)]
        public string MOB { get; set; }
        [DataMember(IsRequired = false)]
        public string EMAIL { get; set; }
        public List<GeoLocator> LocateAccounts(GeoLocator Geo)
        {
            List<GeoLocator> obj = new List<GeoLocator>();
            IOrganizationService service = ConnectToCRM.GetOrgService();
            QueryExpression Qry = new QueryExpression();
            Qry.EntityName = "account";
            ColumnSet Col = new ColumnSet("hil_pincode", "hil_city", "hil_state", "address1_line1", "address1_line2", "address1_line3",
                "name");
            Qry.ColumnSet = Col;

            Qry.PageInfo = new PagingInfo();
            Qry.PageInfo.Count = 50;
            Qry.PageInfo.PageNumber = 1;

            int[] customertypecode = new int[6];
            customertypecode[0] = 4;
            customertypecode[1] = 3;
            customertypecode[2] = 6;
            customertypecode[3] = 5;
            customertypecode[4] = 2;
            customertypecode[5] = 1;

            Qry.Criteria = new FilterExpression(LogicalOperator.And);
            if (Geo.PinCode != null && Geo.PinCode != string.Empty && new Guid(Geo.PinCode) != Guid.Empty)
            {
                Qry.Criteria.AddCondition(new ConditionExpression("hil_pincode", ConditionOperator.Equal, new Guid(Geo.PinCode)));
            }
            if (Geo.City != null && Geo.City != string.Empty && new Guid(Geo.City) != Guid.Empty)
            {
                Qry.Criteria.AddCondition(new ConditionExpression("hil_city", ConditionOperator.Equal, new Guid(Geo.City)));
            }
            if (Geo.State != null && Geo.State != string.Empty && new Guid(Geo.State) != Guid.Empty)
            {
                Qry.Criteria.AddCondition(new ConditionExpression("hil_state", ConditionOperator.Equal, new Guid(Geo.State)));
            }

            Qry.Criteria.AddCondition(new ConditionExpression("customertypecode", ConditionOperator.In, customertypecode));

            EntityCollection Colec = service.RetrieveMultiple(Qry);
            foreach (Entity et in Colec.Entities)
            {
                string Address = string.Empty;
                Account Acc = et.ToEntity<Account>();
                EntityReference Sta = new EntityReference("hil_state");
                EntityReference Cty = new EntityReference("hil_city");
                EntityReference Pncd = new EntityReference("hil_pincode");
                Entity City = new Entity("hil_city");
                Entity PinCode = new Entity("hil_pincode");
                Entity State = new Entity("hil_state");
                string CityName = string.Empty;
                string StateName = string.Empty;
                string PinCodeName = string.Empty;
                if (et.Attributes.Contains("address1_line1"))
                {
                    Address = (string)et["address1_line1"];
                }
                if (et.Attributes.Contains("address1_line2"))
                {
                    Address = Address + (string)et["address1_line2"];
                }
                if (et.Attributes.Contains("address1_line3"))
                {
                    Address = Address + (string)et["address1_line3"];
                }
                if (et.Attributes.Contains("hil_city"))
                {
                    Cty = (EntityReference)et["hil_city"];
                    City = service.Retrieve("hil_city", Cty.Id, new ColumnSet("hil_name"));
                    if (City.Attributes.Contains("hil_name"))
                    {
                        CityName = (string)City["hil_name"];
                        Address = Address.ToUpper() + ", " + CityName.ToUpper() + ", ";
                    }
                }
                if (et.Attributes.Contains("hil_state"))
                {
                    Sta = (EntityReference)et["hil_state"];
                    State = service.Retrieve("hil_state", Sta.Id, new ColumnSet("hil_name"));
                    if (State.Attributes.Contains("hil_name"))
                    {
                        StateName = (string)State["hil_name"];
                        Address = Address.ToUpper() + StateName.ToUpper() + " - ";
                    }
                }
                if (et.Attributes.Contains("hil_pincode"))
                {
                    Pncd = (EntityReference)et["hil_pincode"];
                    PinCode = service.Retrieve("hil_pincode", Pncd.Id, new ColumnSet("hil_name"));
                    if (PinCode.Attributes.Contains("hil_name"))
                    {
                        PinCodeName = (string)PinCode["hil_name"];
                        Address = Address.ToUpper() + PinCodeName.ToUpper();
                    }
                }

                obj.Add(
                new GeoLocator
                {
                    AccountName = Convert.ToString(et["name"]),
                    Address = Address,
                    PinCode = PinCodeName,
                    City = CityName,
                    State = StateName,
                    MOB = Acc.Telephone1,
                    EMAIL = Acc.EMailAddress1,
                    ACCOUNTNAME = Acc.Name.ToUpper()
                });
            }
            return (obj);
        }
    }
}