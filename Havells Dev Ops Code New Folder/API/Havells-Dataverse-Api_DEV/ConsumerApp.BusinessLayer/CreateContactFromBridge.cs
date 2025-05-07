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
    public class CreateContactFromBridge
    {
        [DataMember]
        public string REC_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string CONT_ID { get; set; }//Output
        [DataMember(IsRequired = false)]
        public bool STATUS { get; set; }//Output
        [DataMember(IsRequired = false)]
        public string STATUS_DESC { get; set; }//Output
        public CreateContactFromBridge iCreateContactFromBridge(CreateContactFromBridge iBrdg)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();//get org service obj for connection
            Guid BrdId = new Guid(iBrdg.REC_ID);
            if (BrdId != Guid.Empty)
            {
                hil_consumerappbridge iBrd = (hil_consumerappbridge)service.Retrieve(hil_consumerappbridge.EntityLogicalName, BrdId, new ColumnSet(true));
                Contact iCont = new Contact();
                iCont.EMailAddress1 = iBrd.hil_EmailId;
                iCont.MobilePhone = iBrd.hil_MobileNumber;
                iCont.LastName = iBrd.hil_LastName;
                iCont.hil_Password = iBrd.hil_Password;
                if (iBrd.Attributes.Contains("hil_salutationcode"))
                {
                    int iSal = (int)iBrd["hil_salutationcode"];
                    iCont.hil_Salutation = new OptionSetValue(iSal);
                }
                if (iBrd.hil_FirstName != null)
                    iCont.FirstName = iBrd.hil_FirstName;
                Guid iContId = service.Create(iCont);
                iBrdg.CONT_ID = iContId.ToString();
                iBrdg.STATUS = true;
                iBrdg.STATUS_DESC = "SUCCESS";
            }
            else
            {
                iBrdg.CONT_ID = "";
                iBrdg.STATUS = false;
                iBrdg.STATUS_DESC = "FAILURE - INCORRECT BRIDGE ID";
            }
            return iBrdg;
        }
    }
}
