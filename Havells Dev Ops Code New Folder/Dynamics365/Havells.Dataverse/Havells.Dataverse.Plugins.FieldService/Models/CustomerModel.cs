using System;
using System.Collections.Generic;

namespace Havells.Dataverse.Plugins.FieldService.Models
{
    public class WOSchemesResult
    {
        public List<WOSchemes> lstWOSchemes { get; set; }
        public CustomerResult result { get; set; }
    }
    public class WOSchemes
    {
        public Guid SchemeId { get; set; }
        public string SchemeName { get; set; }
    }
    public class CustomerResult
    {
        public bool ResultStatus { get; set; }
        public string ResultMessage { get; set; }
    }
    public class CustomerResultWithMessageType
    {
        public string ResultMessageType { get; set; }
        public bool ResultStatus { get; set; }
        public string ResultMessage { get; set; }
    }
    public class ServiceResponseData
    {
        public List<InvoiceInfo> InvoiceInfo { get; set; }
        public CustomerResult result { get; set; }
    }
    public class InvoiceInfo
    {
        public Guid InvoiceId { get; set; }
        public string PlanName { get; set; }
        public string CreatedOn { get; set; }
    }
}
