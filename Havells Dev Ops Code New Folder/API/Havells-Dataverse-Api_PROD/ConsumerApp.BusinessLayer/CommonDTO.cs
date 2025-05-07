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
}
