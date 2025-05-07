using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;

namespace HomeAdvisory
{
    public static class common{
        public static Integration IntegrationConfiguration(IOrganizationService service, string Param)
        {
            Integration output = new Integration();
            try
            {
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, Param);
                Entity integrationConfiguration = service.RetrieveMultiple(qsCType)[0];
                output.uri = integrationConfiguration.GetAttributeValue<string>("hil_url");
                output.Auth = integrationConfiguration.GetAttributeValue<string>("hil_username") + ":" + integrationConfiguration.GetAttributeValue<string>("hil_password");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
            return output;
        }
    }
    public class Integration
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
    public class LookUpRes
    {
        public int LookupID { get; set; }
        public string Name { get; set; }
        public string Code { get; set; }
        public int Sequence { get; set; }
        public int MasterTypeID { get; set; }
        public string LookupType { get; set; }
        public String CreatedDate { get; set; }
        public int CreatedBy { get; set; }
        public String ModifiedDate { get; set; }
        public int? ModifiedBy { get; set; }
        public bool IsActive { get; set; }
    }
    public class Root
    {
        public object Result { get; set; }
        public List<LookUpRes> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }
    public class EnquiryLineReq
    {
        public string ParentEnquiryId { get; set; }
        public string EnquiryId { get; set; }
        public string EnquiryLineGuid { get; set; }
        public string EnquiryTypeCode { get; set; }
        public string ProducTypeCode { get; set; }
        public string PinCode { get; set; }
        public string AdvisorCode { get; set; }
        public string CustomerRemarks { get; set; }
        public string AppointmentType { get; set; }
        public string AppointmentDate { get; set; }
        public string AppointmentEndDate { get; set; }
        public string AppointmentId { get; set; }
        public string AppointmentStatus { get; set; }
        public string VideoCallURL { get; set; }
        public string EnquiryStatus { get; set; }
    }
    public class EnquiryReq
    {
        public string EnquiryId { get; set; }
        public string CustomerName { get; set; }
        public string MobileNumber { get; set; }
        public string EmailId { get; set; }
        public string CustomerTypeCode { get; set; }
        public string PropertyTypeCode { get; set; }
        public string ConstructionTypeCode { get; set; }
        public string StateCode { get; set; }
        public string CityCode { get; set; }
        public string PINCode { get; set; }
        public string Area { get; set; }
        public string Rooftop { get; set; }
        public string TDS { get; set; }
        public string AssetType { get; set; }
        public string SourceOfCreation { get; set; }
        public string EnquiryGuid { get; set; }
        public List<EnquiryLineReq> EnquiryLines { get; set; }
    }
    public class EnqueryRequestReq
    {
        public EnquiryReq Data { get; set; }
    }
    public class EnqueryResponseReq
    {
        public int ResponseStatusCode { get; set; }
        public string Message { get; set; }
        public object ServiceToken { get; set; }
        public List<object> Results { get; set; }
        public bool Result { get; set; }
        public bool IsDeleteInsert { get; set; }
        public object EarlierChannelType { get; set; }
        public object MaxLastSyncDate { get; set; }
        public bool IsSuccess { get; set; }
    }
    public class EnquiryLine
    {
        public string ParentEnquiryId { get; set; }
        public string EnquiryId { get; set; }
        public string EnquiryLineGuid { get; set; }
        public string EnquiryTypeCode { get; set; }
        public string ProducTypeCode { get; set; }
        public string PinCode { get; set; }
        public string AdvisorCode { get; set; }
        public string CustomerRemarks { get; set; }
        public string AppointmentType { get; set; }
        public string AppointmentDate { get; set; }
        public string AppointmentEndDate { get; set; }
        public string AppointmentId { get; set; }
        public string AppointmentStatus { get; set; }
        public string VideoCallURL { get; set; }
        public string LineCreatedDate { get; set; }
        public string EnquiryStatus { get; set; }
    }
    public class EnquirySendSFA
    {
        public string EnquiryId { get; set; }
        public string CustomerName { get; set; }
        public string EnquiryCreatedDate { get; set; }
        public string MobileNumber { get; set; }
        public string EmailId { get; set; }
        public string CustomerTypeCode { get; set; }
        public string PropertyTypeCode { get; set; }
        public string ConstructionTypeCode { get; set; }
        public string StateCode { get; set; }
        public string CityCode { get; set; }
        public string PINCode { get; set; }
        public string Area { get; set; }
        public string Rooftop { get; set; }
        public string TDS { get; set; }
        public string AssetType { get; set; }
        public string SourceOfCreation { get; set; }
        public string EnquiryGuid { get; set; }
        public List<EnquiryLine> EnquiryLines { get; set; }
    }
    public class EnqueryRequest
    {
        public EnquirySendSFA Data { get; set; }
    }
    public class EnqueryResponse
    {
        public int ResponseStatusCode { get; set; }
        public string Message { get; set; }
        public object ServiceToken { get; set; }
        public List<object> Results { get; set; }
        public bool Result { get; set; }
        public bool IsDeleteInsert { get; set; }
        public object EarlierChannelType { get; set; }
        public object MaxLastSyncDate { get; set; }
        public bool IsSuccess { get; set; }
    }
    public class Attachments
    {
        public string EnquiryId { get; set; }
        public string FileName { get; set; }
        public string Subject { get; set; }
        public string BlobURL { get; set; }
        public string MIMEType { get; set; }
        public string DocType { get; set; }
        public string Size { get; set; }
        public string DocGuid { get; set; }
        public bool IsDelete { get; set; }
    }
    public class AttachmentsList
    {
        public List<Attachments> Data { get; set; }
    }
    public class AttachmentResponse
    {
        public int ResponseStatusCode { get; set; }
        public string Message { get; set; }
        public object ServiceToken { get; set; }
        public List<object> Results { get; set; }
        public bool Result { get; set; }
        public bool IsDeleteInsert { get; set; }
        public object EarlierChannelType { get; set; }
        public object MaxLastSyncDate { get; set; }
        public bool IsSuccess { get; set; }
    }
    public class EnquryType
    {
        public string Result { get; set; }
        public string Success { get; set; }
        public string Message { get; set; }
        public string ErrorCode { get; set; }
        public List<EnquryTypeDate> Results { get; set; }
    }
    public class EnquryTypeDate
    {
        public string EnquiryTypeCode { get; set; }
        public string EnquiryTypeDesc { get; set; }
        public string CreatedDate { get; set; }
        public string CreatedBy { get; set; }
        public string ModifiedDate { get; set; }
        public string ModifiedBy { get; set; }
        public bool IsActive { get; set; }
    }
    public class ProductType
    {
        public string Result { get; set; }
        public string Success { get; set; }
        public string Message { get; set; }
        public string ErrorCode { get; set; }
        public List<ProducttypeResult> Results { get; set; }
    }
    public class ProducttypeResult
    {
        public string AdvisoryProductTypeCode { get; set; }
        public string AdvisoryProductType { get; set; }
        public string CreatedDate { get; set; }
        public string CreatedBy { get; set; }
        public string ModifiedDate { get; set; }
        public string ModifiedBy { get; set; }
        public bool IsActive { get; set; }
    }
    public class GetEnquiryProductMapping
    {
        public string Result { get; set; }
        public string Success { get; set; }
        public string Message { get; set; }
        public string ErrorCode { get; set; }
        public List<EnquiryTypeprodmapping> Results { get; set; }
    }
    public class EnquiryTypeprodmapping
    {
        public string EnquiryTypeCode { get; set; }
        public string AdvisoryProductTypeCode { get; set; }
        public string CreatedDate { get; set; }
        public string CreatedBy { get; set; }
        public string ModifiedDate { get; set; }
        public string ModifiedBy { get; set; }
        public bool IsActive { get; set; }
    }
    public class GetPropertyType
    {
        public string Result { get; set; }
        public string Success { get; set; }
        public string Message { get; set; }
        public string ErrorCode { get; set; }
        public List<PropertyType> Results { get; set; }
    }
    public class PropertyType
    {
        public string LookupID { get; set; }
        public string Name { get; set; }
        public string Code { get; set; }
        public string Sequence { get; set; }
        public string MasterTypeID { get; set; }
        public string LookupType { get; set; }
        public string CreatedDate { get; set; }
        public string CreatedBy { get; set; }
        public string ModifiedDate { get; set; }
        public string ModifiedBy { get; set; }
        public bool IsActive { get; set; }
    }
    public class WorkflowData
    {
        public string Result { get; set; }
        public string Success { get; set; }
        public string Message { get; set; }
        public string ErrorCode { get; set; }
        public List<AdvisordivisionSetup> Results { get; set; }
    }
    public class AdvisordivisionSetup
    {
        public string EnquiryTypeCode { get; set; }
        public string AdvisoryProductTypeCode { get; set; }
        public string UserCode { get; set; }
        public string SaleOfficeCode { get; set; }
        public string CreatedDate { get; set; }
        public string CreatedBy { get; set; }
        public string ModifiedBy { get; set; }
        public string ModifiedDate { get; set; }
        public string IsActive { get; set; }
        public string WorkFlowDataId { get; set; }
        public string Workflow { get; set; }
        public string Isactive { get; set; }
        public string UserName { get; set; }
        public string MobileNumber { get; set; }
        public string EmailId { get; set; }

    }
}
