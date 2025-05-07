using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.JobSheet
{
    public class Models
    {
        private static String[] units = { "Zero", "One", "Two", "Three",
    "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven",
    "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
    "Seventeen", "Eighteen", "Nineteen" };
        private static String[] tens = { "", "", "Twenty", "Thirty", "Forty",
    "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" };
        public static String ConvertAmount(double amount)
        {
            try
            {
                Int64 amount_int = (Int64)amount;
                Int64 amount_dec = (Int64)Math.Round((amount - (double)(amount_int)) * 100);
                if (amount_dec == 0)
                {
                    return Convert(amount_int) + " Only.";
                }
                else
                {
                    return Convert(amount_int) + " Point " + Convert(amount_dec) + " Only.";
                }
            }
            catch (Exception e)
            {
                // TODO: handle exception  
            }
            return "";
        }
        public static String Convert(Int64 i)
        {
            if (i < 20)
            {
                return units[i];
            }
            if (i < 100)
            {
                return tens[i / 10] + ((i % 10 > 0) ? " " + Convert(i % 10) : "");
            }
            if (i < 1000)
            {
                return units[i / 100] + " Hundred"
                        + ((i % 100 > 0) ? " And " + Convert(i % 100) : "");
            }
            if (i < 100000)
            {
                return Convert(i / 1000) + " Thousand "
                + ((i % 1000 > 0) ? " " + Convert(i % 1000) : "");
            }
            if (i < 10000000)
            {
                return Convert(i / 100000) + " Lakh "
                        + ((i % 100000 > 0) ? " " + Convert(i % 100000) : "");
            }
            if (i < 1000000000)
            {
                return Convert(i / 10000000) + " Crore "
                        + ((i % 10000000 > 0) ? " " + Convert(i % 10000000) : "");
            }
            return Convert(i / 1000000000) + " Arab "
                    + ((i % 1000000000 > 0) ? " " + Convert(i % 1000000000) : "");
        }
    }
    public class jobModel
    {
        public string msdyN_WORKORDERID { get; set; }
        public string hiL_PRODUCTCATEGORYNAME { get; set; }
        public string hiL_PINCODENAME { get; set; }
        public string hiL_PRODUCTCATSUBCATMAPPINGNAME { get; set; }
        public string hiL_MOBILENUMBER { get; set; }
        public string jobnumber { get; set; }
        public string hiL_FULLADDRESS { get; set; }
        public string hiL_CUST { get; set; }
        public string omercomplaintdescription { get; set; }
        public string hiL_CUSTOMERREFNAME { get; set; }
        public string hiL_CALLINGNUMBER { get; set; }
        public string hiL_ALTERNATE { get; set; }
        public string hiL_WARRANTYSTATUS { get; set; }
        public string hiL_PURCHASEDFROM { get; set; }
        public string hiL_COUNTRYCLASSIFICATION { get; set; }
        public string hiL_ACTUALCHARGES { get; set; }
        public string hiL_QUANTITY { get; set; }
        public string owneridname { get; set; }
        public string hiL_SENDTCR { get; set; }
        public string msdyN_WORKORDERINCIDENTID { get; set; }
        public string hiL_OBSERVATIONNAME { get; set; }
        public string hiL_MODELNAME { get; set; }
        public string msdyN_CUSTOMERASSETNAME { get; set; }
        public string msdyN_INCIDENTTYPENAME { get; set; }
        public string hiL_INVOICENO { get; set; }
        public string hiL_INVOICEDATE { get; set; }
        public string prddesc { get; set; }
        public string warrantystatus { get; set; }
        public string partname { get; set; }
        public string partprice { get; set; }
        public string servicename { get; set; }
        public string serviceamount { get; set; }
    }
    public class Job
    {
        public string workOrderId { get; set; }
        public string productCategory { get; set; }
        public string pincode { get; set; }
        public string productSubCategory { get; set; }
        public string mobileNumber { get; set; }
        public string jobNumber { get; set; }
        public string fullAddress { get; set; }
        public string complaintDescription { get; set; }
        public string customerName { get; set; }
        public string callingNumber { get; set; }
        public string alternateNumber { get; set; }
        public string jonWarranty { get; set; }
        
        public string countryClassification { get; set; }
        public string actualCharges { get; set; }
        public string quantity { get; set; }
        public string owner { get; set; }
        public string VisitDate { get; set; }
        public string sendTRC { get; set; }
    }
    public class IncidentData
    {
        public string incidentId { get; set; }
        public string TechnicianRemarks { get; set; }
        public string observationName { get; set; }
        public string modelName { get; set; }
        public string assetNo { get; set; }
        public string incidentType { get; set; }
        public string invoiceNo { get; set; }
        public string invoiceDate { get; set; }
        public string productDesription { get; set; }
        public string purchasedFrom { get; set; }
    }
    public class SparePartAndProductData
    {
        public string recordID { get; set; }
        public string amount { get; set; }
        public string warranty { get; set; }
        public string product { get; set; }
        public string recordType { get; set; }
        public string index { get; set; }
    }
    public class WorkOrder
    {
        public List<Job> job { get; set; }
        public List<IncidentData> incidentData { get; set; }
        public List<SparePartAndProductData> sparePartAndProductData { get; set; }
    }
    public class JobResponseAPI
    {
        public List<WorkOrder> workOrders { get; set; }
        public bool success { get; set; }
        public string error { get; set; }
    }
    public class IntegrationConfiguration
    {
        public string url { get; set; }
        public string userName { get; set; }
        public string password { get; set; }
    }
}
