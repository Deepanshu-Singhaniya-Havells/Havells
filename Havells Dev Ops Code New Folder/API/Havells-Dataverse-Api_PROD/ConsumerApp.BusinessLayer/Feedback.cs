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
    public class Feedback
    {
        [DataMember]
        public int TYPE { get; set; }
        [DataMember]
        public string MESSAGE { get; set; }
        [DataMember]
        public string CONT_ID { get; set; }
        [DataMember(IsRequired = false)]
        public string STATUS { get; set; }
        [DataMember]
        public string ServiceAttachment { get; set; }
        [DataMember]
        public int FileType { get; set; }
        [DataMember(IsRequired = false)]
        public string EXCEPTION_DESC { get; set; }
        public Feedback PostFeedback(Feedback Feed)
        {
            IOrganizationService service = ConnectToCRM.GetOrgService();
            try
            {
                Entity iFeedback = new Entity("hil_feedback");
                iFeedback["hil_type"] = new OptionSetValue(Feed.TYPE);
                iFeedback["hil_message"] = Feed.MESSAGE;
                iFeedback["hil_consumer"] = new EntityReference(Contact.EntityLogicalName, new Guid(Feed.CONT_ID));
                Guid enJobId = service.Create(iFeedback);
                if (enJobId != Guid.Empty)
                {
                    new ProductRegistration().AttachNotes(service, Feed.ServiceAttachment, enJobId, Feed.FileType, "hil_feedback");
                }
                Feed.STATUS = "SUCCESS";
                Feed.EXCEPTION_DESC = "";
            }
            catch (Exception ex)
            {
                Feed.STATUS = "FAILURE";
                Feed.EXCEPTION_DESC = ex.Message.ToUpper();
            }
            return Feed;
        }
    }
}