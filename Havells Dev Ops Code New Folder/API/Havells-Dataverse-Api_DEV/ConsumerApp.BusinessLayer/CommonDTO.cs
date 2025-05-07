using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Runtime.Serialization;

namespace ConsumerApp.BusinessLayer
{
    [DataContract]
    public class HashTableDTO
    {
        [DataMember]
        public string Label { get; set; }

        [DataMember]
        public int? Value { get; set; }

        [DataMember]
        public string Extension { get; set; }

        [DataMember]
        public string StatusCode { get; set; }

        [DataMember]
        public string StatusDescription { get; set; }
    }

    [DataContract]
    public class HashTableGuidDTO
    {
        [DataMember]
        public string Label { get; set; }

        [DataMember]
        public Guid Value { get; set; }

        [DataMember]
        public string StatusCode { get; set; }

        [DataMember]
        public string StatusDescription { get; set; }
    }

    [DataContract]
    public class ProductHierarchyDTO
    {
        [DataMember]
        public string ProductCategory { get; set; }

        [DataMember]
        public Guid ProductCategoryGuid { get; set; }

        [DataMember]
        public string ProductSubCategory { get; set; }

        [DataMember]
        public Guid ProductSubCategoryGuid { get; set; }

        [DataMember]
        public bool IsSerialized { get; set; }

        [DataMember]
        public bool IsVerificationrequired { get; set; }

        [DataMember]
        public string StatusCode { get; set; }

        [DataMember]
        public string StatusDescription { get; set; }
    }

    [DataContract]
    public class ProductDTO
    {
        [DataMember]
        public string ProductCode { get; set; }

        [DataMember]
        public string Product { get; set; }

        [DataMember]
        public Guid ProductGuid { get; set; }

        [DataMember]
        public string StatusCode { get; set; }

        [DataMember]
        public string StatusDescription { get; set; }
    }

    public class SMSService {
        public static void sendSMS(IOrganizationService service, String _ConsumerMobile, String _templateName)
        {
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_smstemplates");
                qsCType.ColumnSet = new ColumnSet("hil_templatebody");
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, _templateName);
                Entity _smsTemp = service.RetrieveMultiple(qsCType)[0];
                if (_smsTemp != null)
                {
                    Entity _SMS = new Entity("hil_smsconfiguration");
                    _SMS["hil_smstemplate"] = _smsTemp.ToEntityReference();
                    _SMS["subject"] = "Havells Consumer Connect App Migration";
                    _SMS["hil_message"] = _smsTemp.GetAttributeValue<string>("hil_templatebody");
                    _SMS["hil_mobilenumber"] = _ConsumerMobile;
                    _SMS["hil_direction"] = new OptionSetValue(2);
                    service.Create(_SMS);
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
    }
}
