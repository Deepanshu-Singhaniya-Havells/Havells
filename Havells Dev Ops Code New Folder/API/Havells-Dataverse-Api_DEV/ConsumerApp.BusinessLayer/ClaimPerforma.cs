using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Pechkin;
using Pechkin.Synchronized;

namespace ConsumerApp.BusinessLayer
{
    public static class ConvertToDataTable
    {
        public static DataTable PropertiesToDataTable<T>(this IEnumerable<T> source)
        {
            DataTable dt = new DataTable();
            var props = TypeDescriptor.GetProperties(typeof(T));
            foreach (PropertyDescriptor prop in props)
            {
                DataColumn dc = dt.Columns.Add(prop.Name, prop.PropertyType);
                dc.Caption = prop.DisplayName;
                dc.ReadOnly = prop.IsReadOnly;
            }
            foreach (T item in source)
            {
                DataRow dr = dt.NewRow();
                foreach (PropertyDescriptor prop in props)
                {
                    dr[prop.Name] = prop.GetValue(item);
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }

    [DataContract]
    public class ClaimPerforma
    {
        public AnnextureResponse GenerateAnnexureUrl(string _PerformaInvoice)
        {
            string AnnextureUrl = string.Empty;

            IOrganizationService _service = ConnectToCRM.GetOrgService();
            if (_service != null)
            {
                try
                {
                    String fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='hil_claimheader'>
                            <attribute name='hil_claimheaderid' />
                            <attribute name='hil_name' />
                            <attribute name='hil_franchiseeinvoiceno' />
                            <attribute name='hil_franchisee' />
                            <attribute name='hil_fiscalmonth' />
                            <attribute name='hil_sapponumber'/>
                            <attribute name='hil_syncdoneon'/>
                            <attribute name='modifiedon'/>
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                              <condition attribute='hil_performastatus' operator='eq' value='4' />
                              <condition attribute='hil_name' operator='eq' value='" + _PerformaInvoice + @"' />
                            </filter>
                            <link-entity name='account' from='accountid' to='hil_franchisee' visible='false' link-type='outer' alias='ab'>
                              <attribute name='hil_vendorcode' />
                              <attribute name='hil_state' />
                              <attribute name='hil_salesoffice' />
                              <attribute name='hil_district' />
                              <attribute name='hil_city' />
                              <attribute name='hil_category' />
                              <attribute name='hil_branch' />
                              <attribute name='hil_area' />
                              <attribute name='accountnumber' />
                            </link-entity>
                            <link-entity name='hil_claimperiod' from='hil_claimperiodid' to='hil_fiscalmonth' visible='false' link-type='outer' alias='ac'>
                                  <attribute name='hil_todate' />
                                  <attribute name='hil_fromdate' />
                            </link-entity>
                          </entity>
                        </fetch>";

                    EntityCollection ClaimInvoiceCollection = _service.RetrieveMultiple(new FetchExpression(fetchXML));

                    if (ClaimInvoiceCollection.Entities.Count > 0)
                    {
                        EntityReference _performaInvoiceRef = ClaimInvoiceCollection[0].ToEntityReference();

                        //check if attachmemnt is exist or not
                        string DocUrl = findAttachment(_service, _performaInvoiceRef);
                        if (string.IsNullOrWhiteSpace(DocUrl))
                        {
                            string ClaimfetchXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                          <entity name='hil_claimline'>
                            <attribute name='hil_claimlineid' />
                            <attribute name='hil_name' />
                            <attribute name='createdon' />
                            <attribute name='hil_productcategory' />
                            <attribute name='hil_claimcategory' />
                            <attribute name='hil_claimamount_base' />
                            <attribute name='hil_claimamount' />
                            <attribute name='hil_callsubtype' />
                            <attribute name='hil_activitycode' />
                            <order attribute='hil_name' descending='false' />
                            <filter type='and'>
                                <condition attribute='hil_claimheader' operator='eq' uiname='CL-00032852' uitype='hil_claimheader' value='{" + _performaInvoiceRef.Id + @"}' />
                                <condition attribute='statecode' operator='eq' value='0' />
                            </filter>
                            <link-entity name='hil_claimcategory' from='hil_claimcategoryid' to='hil_claimcategory' visible='false' link-type='outer' alias='ac'>
                              <attribute name='hil_claimtype' />
                            </link-entity>
                        </entity>
                        </fetch>";

                            EntityCollection ClaimLinesInvoiceCollection = _service.RetrieveMultiple(new FetchExpression(ClaimfetchXml));

                            DataTable dt = new DataTable();
                            dt.Columns.Add("ActivityCode", typeof(string));
                            dt.Columns.Add("ProductCategory", typeof(string));
                            dt.Columns.Add("CallSubType", typeof(string));
                            dt.Columns.Add("ClaimType", typeof(string));
                            dt.Columns.Add("ClaimCategory", typeof(string));
                            dt.Columns.Add("ClaimAmount", typeof(decimal));

                            DataRow dr;
                            foreach (Entity row in ClaimLinesInvoiceCollection.Entities)
                            {
                                dr = dt.NewRow();
                                dr["ActivityCode"] = row.Contains("hil_activitycode") ? row.GetAttributeValue<string>("hil_activitycode").ToString() : "";
                                dr["ProductCategory"] = row.Contains("hil_productcategory") ? row.GetAttributeValue<EntityReference>("hil_productcategory").Name.ToString() : "";
                                dr["CallSubType"] = row.Contains("hil_callsubtype") ? row.GetAttributeValue<EntityReference>("hil_callsubtype").Name.ToString() : "";
                                dr["ClaimType"] = row.Contains("ac.hil_claimtype") ? (row.GetAttributeValue<AliasedValue>("ac.hil_claimtype").Value as EntityReference).Name.ToString() : "";
                                dr["ClaimCategory"] = row.Contains("hil_claimcategory") ? row.GetAttributeValue<EntityReference>("hil_claimcategory").Name.ToString() : "";
                                dr["ClaimAmount"] = row.Contains("hil_claimamount") ? Convert.ToDecimal(row.GetAttributeValue<Money>("hil_claimamount").Value.ToString("0.00")) : 0;
                                dt.Rows.Add(dr);
                            }

                            DataTable Newdt = dt.AsEnumerable()
                              .GroupBy(r => new
                              {
                                  ActivityCode = r["ActivityCode"],
                                  ProductCategory = r["ProductCategory"],
                                  CallSubType = r["CallSubType"],
                                  ClaimType = r["ClaimType"],
                                  ClaimCategory = r["ClaimCategory"]
                              })
                              .Select(x => new ClaimLine
                              {
                                  ActivityCode = x.Key.ActivityCode,
                                  ProductCategory = x.Key.ProductCategory,
                                  CallSubType = x.Key.CallSubType,
                                  ClaimType = x.Key.ClaimType,
                                  ClaimCategory = x.Key.ClaimCategory,
                                  ClaimAmount = x.Sum(z => z.Field<decimal>("ClaimAmount"))
                              }).PropertiesToDataTable<ClaimLine>();

                            string fetchCategoryXml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='hil_claimcategory'>
                                <attribute name='hil_name' />
                                <attribute name='modifiedon' />
                                <attribute name='hil_claimtype' />
                                <attribute name='hil_displayindex' />
                                <attribute name='hil_claimcategoryid' />
                                <order attribute='hil_claimtype' descending='false' />
                                <order attribute='hil_displayindex' descending='false' />
                                <filter type='and'>
                                  <condition attribute='statecode' operator='eq' value='0' />
                                </filter>
                                <link-entity name='hil_claimtype' from='hil_claimtypeid' to='hil_claimtype' visible='false' link-type='outer' alias='ab'>
                                  <attribute name='hil_displayindex' />
                                </link-entity>
                              </entity>
                            </fetch>";

                            EntityCollection ClaimCategoryCollection = _service.RetrieveMultiple(new FetchExpression(fetchCategoryXml));

                            DataTable Categorydt = new DataTable();
                            Categorydt.Columns.Add("ClaimType", typeof(string));
                            Categorydt.Columns.Add("TypeIndex", typeof(int));
                            Categorydt.Columns.Add("ClaimCategory", typeof(string));
                            Categorydt.Columns.Add("CategoryIndex", typeof(int));

                            DataRow cdr;
                            foreach (Entity row in ClaimCategoryCollection.Entities)
                            {
                                cdr = Categorydt.NewRow();
                                cdr["ClaimType"] = row.Contains("hil_claimtype") ? row.GetAttributeValue<EntityReference>("hil_claimtype").Name.ToString() : "";
                                cdr["ClaimCategory"] = row.Contains("hil_name") ? row.GetAttributeValue<string>("hil_name") : "";
                                cdr["CategoryIndex"] = row.Contains("hil_displayindex") ? row.GetAttributeValue<int>("hil_displayindex") : 0;
                                cdr["TypeIndex"] = row.Contains("ab.hil_displayindex") ? (int)row.GetAttributeValue<AliasedValue>("ab.hil_displayindex").Value : 0;
                                Categorydt.Rows.Add(cdr);
                            }

                            AnnextureUrl = documentHtml(_service, ClaimInvoiceCollection, Newdt, Categorydt, _performaInvoiceRef);
                        }

                    }

                }
                catch (Exception ex)
                {
                    return new AnnextureResponse() { URL = AnnextureUrl, ResultStatus = false, ResultMessage = "D365 Service Unavailable. " + ex.Message };

                }
                return new AnnextureResponse() { URL = AnnextureUrl, ResultStatus = true, ResultMessage = "D365 Service Unavailable." };
            }
            else
            {
                return new AnnextureResponse() { URL = AnnextureUrl, ResultStatus = false, ResultMessage = "D365 Service Unavailable." };
            }
        }

        #region supportingMethod
        static string documentHtml(IOrganizationService _service, EntityCollection ClaimInvoiceCollection, DataTable dt, DataTable Catdt, EntityReference performaInvoice)
        {
            string AnnextureURL = "";
            string accountnumber = "";
            string hil_name = "";
            string hil_vendorcode = "";
            string hil_category = "";
            string hil_salesoffice = "";
            string hil_fiscalmonth = "";
            DateTime? hil_syncdoneon = null;
            DateTime? hil_fromdate = null;
            DateTime? hil_todate = null;
            string hil_franchisee = "";
            string hil_district = "";
            string hil_area = "";
            string AreaLocation = "";
            DateTime? modifiedon = null;
            string hil_sapponumber = "";
            try
            {
                if (_service != null)
                {

                    if (ClaimInvoiceCollection.Entities.Count > 0)
                    {
                        if (ClaimInvoiceCollection.Entities[0].Contains("hil_name"))
                        {
                            hil_name = ClaimInvoiceCollection.Entities[0].GetAttributeValue<string>("hil_name");
                        }
                        if (ClaimInvoiceCollection.Entities[0].Contains("hil_franchisee"))
                        {
                            hil_franchisee = ClaimInvoiceCollection.Entities[0].GetAttributeValue<EntityReference>("hil_franchisee").Name;
                        }
                        if (ClaimInvoiceCollection.Entities[0].Contains("ab.accountnumber"))
                        {
                            accountnumber = ClaimInvoiceCollection.Entities[0].GetAttributeValue<AliasedValue>("ab.accountnumber").Value.ToString();
                        }
                        if (ClaimInvoiceCollection.Entities[0].Contains("ab.hil_vendorcode"))
                        {
                            hil_vendorcode = ClaimInvoiceCollection.Entities[0].GetAttributeValue<AliasedValue>("ab.hil_vendorcode").Value.ToString();
                        }
                        if (ClaimInvoiceCollection.Entities[0].Contains("ab.hil_category"))
                        {
                            hil_category = ClaimInvoiceCollection.Entities[0].FormattedValues["ab.hil_category"].ToString();
                        }
                        if (ClaimInvoiceCollection.Entities[0].Contains("ab.hil_district"))
                        {
                            hil_district = ((EntityReference)ClaimInvoiceCollection.Entities[0].GetAttributeValue<AliasedValue>("ab.hil_district").Value).Name.ToString();
                        }
                        if (ClaimInvoiceCollection.Entities[0].Contains("ab.hil_area"))
                        {
                            hil_area = ((EntityReference)ClaimInvoiceCollection.Entities[0].GetAttributeValue<AliasedValue>("ab.hil_area").Value).Name.ToString();
                        }
                        if (ClaimInvoiceCollection.Entities[0].Contains("ab.hil_salesoffice"))
                        {
                            hil_salesoffice = ((EntityReference)ClaimInvoiceCollection.Entities[0].GetAttributeValue<AliasedValue>("ab.hil_salesoffice").Value).Name.ToString();
                        }
                        if (ClaimInvoiceCollection.Entities[0].Contains("hil_fiscalmonth"))
                        {
                            hil_fiscalmonth = ClaimInvoiceCollection[0].GetAttributeValue<EntityReference>("hil_fiscalmonth").Name;
                        }
                        if (ClaimInvoiceCollection.Entities[0].Contains("hil_syncdoneon"))
                        {
                            hil_syncdoneon = ClaimInvoiceCollection.Entities[0].GetAttributeValue<DateTime>("hil_syncdoneon").AddMinutes(330);
                        }
                        if (ClaimInvoiceCollection.Entities[0].Contains("ac.hil_fromdate"))
                        {
                            hil_fromdate = Convert.ToDateTime(ClaimInvoiceCollection.Entities[0].GetAttributeValue<AliasedValue>("ac.hil_fromdate").Value).AddMinutes(330);
                        }
                        if (ClaimInvoiceCollection.Entities[0].Contains("ac.hil_todate"))
                        {
                            hil_todate = Convert.ToDateTime(ClaimInvoiceCollection.Entities[0].GetAttributeValue<AliasedValue>("ac.hil_todate").Value).AddMinutes(330);
                        }
                        if (ClaimInvoiceCollection.Entities[0].Contains("hil_sapponumber"))
                        {
                            hil_sapponumber = ClaimInvoiceCollection.Entities[0].GetAttributeValue<string>("hil_sapponumber");
                        }
                        if (ClaimInvoiceCollection.Entities[0].Contains("modifiedon"))
                        {
                            modifiedon = ClaimInvoiceCollection.Entities[0].GetAttributeValue<DateTime>("modifiedon").AddMinutes(330);
                        }

                        AreaLocation = hil_district + "-" + hil_area;

                        StringBuilder _errorMsg = new StringBuilder();

                        Console.WriteLine("************************************************************************");
                        Console.WriteLine("Started");
                        #region HTML Code
                        String header = @"<!DOCTYPE html>
                        <html>
                        <head>
                            <title>Invoice</title>
                            <style>
                        body{background:#ededed;font-family:'Calibri';}
		                        table{background: #fff;  padding: 15px; width: 100%;border-collapse: collapse; }
		                        table tr td{padding:2px; font-size:11px;border:1px solid #ddd;text-align:center;}
		                        h1{font-size:18px;font-weight:600;text-align:center;margin: 5px 0;}
                                .firstTable{width:100%;}
		                        .firstTable td{text-align:left;}
		                        .firstTable tr td:first-child{font-weight:bold;}
		                        .box{padding:15px;max-width:100%;margin:0 auto; width: 100%;}
                        </style>
                        </head>

                        <body>
                        <div class='box'>
                            <table border='0' style='padding:0;border-collapse: collapse;'>
                            <tr><td style='border:0;'>
                           <table style='padding:0;border-collapse: collapse;'>
                           <tr style='background:#f5f5f5;'>
		                        <td colspan='3'><h1>Annexure to Invoice - Franchisee Claim Summary for Branch Commercial</h1></td>			
                           </tr>
  
                           <tr>
	                           <td style='width:33.33%;padding:0 5% 0 0;'>
		                           <table class='firstTable'>
		                           <tr>
				                        <td>Performa Number:</td>
				                        <td>" + hil_name + @"</td>
		                           </tr>
			                         <tr>
				                        <td>Franchisee's Customer code:</td>
				                        <td>" + accountnumber + @"</td>
		                           </tr>
			                         <tr>
				                        <td>Franchisee's Vendor  code:</td>
				                        <td>" + hil_vendorcode + @"</td>
		                           </tr>
			                         <tr>
				                        <td>Franchisee Name</td>
				                        <td>" + hil_franchisee + @"</td>
		                           </tr>
			                         <tr>
				                        <td>Category For Incentive (A/B/C)</td>
				                        <td>" + hil_category + @"</td>
		                           </tr>
			                        </table>
		                        </td>
	                        <td style='width:33.33%;padding:0 5% 0 0;'>
                           <table class='firstTable' >  
                            <tr>
                                <td>Area/ Location</td>
                                <td>" + AreaLocation + @"</td>
       
                            </tr>
                            <tr>
                                <td>Branch</td>
                                <td>" + hil_salesoffice + @"</td>
        
                            </tr>
                            <tr>
                                <td>Billing Period</td>
                                <td>" + (hil_fromdate == null ? "" : String.Format("{0:dd/MM/yyyy}", hil_fromdate)) + (hil_todate == null ? "" : " to " + String.Format("{0:dd/MM/yyyy}", hil_todate)) + @"</td>        
                            </tr>
                            <tr>
                                <td>Approval Date</td>
                                <td>" + (modifiedon == null ? "" : String.Format("{0:dd/MM/yyyy}", modifiedon)) + @"</td>
                            </tr>
	                         <tr>
		                        <td>PO Number & Date</td>
		                        <td>" + hil_sapponumber + " & " + (hil_syncdoneon == null ? "" : String.Format("{0:dd/MM/yyyy}", hil_syncdoneon)) + @"</td>
                           </tr>
                        </table>
                        </td>
	                        <td style='width:33.33%;padding:0 5% 0 0;'>
		                        <table class='firstTable'>	
			                        <tr>
			                            <td style='border:none;text-align:right;'><img src='data:image/png;base64, iVBORw0KGgoAAAANSUhEUgAAANoAAABTCAYAAAD5ohnLAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAyZpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDcuMS1jMDAwIDc5LjljY2M0ZGU5MywgMjAyMi8wMy8xNC0xNDowNzoyMiAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bWxuczp4bXBNTT0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL21tLyIgeG1sbnM6c3RSZWY9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZVJlZiMiIHhtcDpDcmVhdG9yVG9vbD0iQWRvYmUgUGhvdG9zaG9wIDIzLjMgKFdpbmRvd3MpIiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOkUwQkRDQUQzRjZDMDExRUNBOTkyODZCNzhEMTBCODgzIiB4bXBNTTpEb2N1bWVudElEPSJ4bXAuZGlkOkUwQkRDQUQ0RjZDMDExRUNBOTkyODZCNzhEMTBCODgzIj4gPHhtcE1NOkRlcml2ZWRGcm9tIHN0UmVmOmluc3RhbmNlSUQ9InhtcC5paWQ6RTBCRENBRDFGNkMwMTFFQ0E5OTI4NkI3OEQxMEI4ODMiIHN0UmVmOmRvY3VtZW50SUQ9InhtcC5kaWQ6RTBCRENBRDJGNkMwMTFFQ0E5OTI4NkI3OEQxMEI4ODMiLz4gPC9yZGY6RGVzY3JpcHRpb24+IDwvcmRmOlJERj4gPC94OnhtcG1ldGE+IDw/eHBhY2tldCBlbmQ9InIiPz73h6SKAAAttUlEQVR42uxdB3hUZdY+09MLKXSB0HtR6UWKoNJEsbsi61pB3FVWd1UEVBARBXXX/deGXVEpigUUlSa9917SSCAhvUz/z3tmbpjM3ElCCBHde3zuE7kzc+93v/u9p5/z6dxuN2mkkUYXl/TaFGikkQY0jTTSgKaRRhppQNNIIw1oGmmkAU0jjTTSgKaRRhrQNNJIIw1oGmmkAU0jjf6gZLyA39ZK7lZRUTFlZGSSy+0iJV3MaDRRTHQU1akTq71BjWqbdH9IiVZqtZLZYiGn08WHWw78u7i4WHvlGv1PSLQLIofDQS4Xg8ZsqhhoJaUMLifZ7fayc9bSUpFwOG8wGIL+Fp87+LCYzdqb1uh/D2gFhYWUm5tHNpudGjduSGaTOtigKkKiqYFUp9PJ32BAc7lclJqWTtA2oyIjKCYmmvR6zSTV6LehWl15paVWOnUqUw6ogiaTmU6lZ5DT5VT9vt1hF6mEw19SwUS02qxB75WZeVoeD4DLYVCnpZ2igoJC0sqCNPpDAy0/v0CcGnaWQgBKfn4+H3mkY4mUzmADIPzJZrWXSTB/oFksISIR1ej06SyWhDYqKiqkQpaepaxqOvn6Z3NyKTs7R3vrGv2xJRoWOxa+ovpBuhQw4GCrARy+qmFhURFln82R7/lLNAEh22yQUACwzWYrO5+XlyeqqZVVTvwOvweICwoKyGg08hgc2lvXqNZJdwGq1Hn9EIv+6PETpNfpqaSkpLyhyABw83+x0TEe8BV6VDzYVFZrcPXQDCcHjwJgCgkNYSlnplyWWnqWkrgHzitkYjsQ/76scSMK5e9qpFF1MXNJSzQ4LSIjI2XBqzk33CzVSljFs7J0AsAANF9JpUb4XOw7nQfIiLm5+Hf+IBNQWizifdRAptFvQbXqdYwIC6MMVvcAOn91EOpdEauL562Oep0lvpLPH2TCURi8keGh2hvX6I8PtLDQUF7wrOaFhFQIKgAFYIRKCTUQDhNIKreoiQwa2HcMLhfiZF7nSkUqMFRMGwMxLDFOe+Ma/fGBZjAaWIUzk9PhCgqwEAajgQFmhRcy1MIgYZUPnzkYWHY76fkaLv4cSmWp3UY6q51CAUJINVYZXSqA86irbga4JtE0+gMArZgXOiRRsEyM3Lw8studktnhDzBIOxuDyxAVSWGs5iVkZZNh715yHzlClJFJxvwCciM7xMD2W0QkuRLjiZo0JWfrVlRUN5FKDEbonxSuN/A4istJOLj3AbasrCyKj4tTVS010uh3ATQEozMyTpObF3sogyac7SH8NXtBl5eXLy582EpIp5LF7naLxDKFWMjN6mSiUU9h23eS/qdfyL51KznOZFEIA0en14mzxy2OTp1IJwP/1orMkNhYim3bmmIGXkXFvXpSLl/HYtCRy2orc6ZAtcR9CwqKxUsZH6+BTaPapRpz76dLhodLPH4Aj9PBYOLFHRJikXzGM2eyJROkuLiobJGHhYWTXeemWJYyUTt3keGTz8i+aTNZGFx2L6xMsNfKQOa5KVylUD7tbu93yBNX07N0091yM+UP6Es5xSzFGIiKLYjnrFOnDv91UaOGDTSgaVRtzPxmQEMmfWpquthHSvIvJAikGVQ2SBRIlzJJxhQZEUEUHkaxLK3CP/6UbJ8sIIvLTcUuJ5n5O3B4FPFvLS2ak7tRQ2KUkI6B6bax2pmTQ7pTmWQ/fITMBQVk4O+WMshD+J5wehiGDqHS+++lnJhY1NlQIX9HL6APoTqxMRTLh0Ya/a6ABrd8SmqaeASRfRHMyeF7H8TTWK+kBAaN6cWXybhhIxUjQM3fM/Jz2JKakmnYUDL27UW6Ro1IHxYWeHMGruv0aXKxBLQt/5Fo63Yy6vRkZYkVwX+dzZqS9YnH6EzTZhiYgNzFIG7G57Vsfo1+d0BDkPnEiWSysIpYVFhY4ThwrwiWZDq23xJtdjJOmUaWvfsplwEQwepiaUw0mf50O5lHjSBddHTVR8Igsv+8kuzz3yfLsRNUKNfTk4NtMcf0ZyiTwRalR0ZKETVm4Gqk0e9SdUSCcFb2WZZcesllDEaQJPrQEEoMCyXzlOkUsmUb5TodFGkwUkmHdhT6j7+ToVXLas+AKyuLrP/6Dxm+W0Yl4oHUk71uXbLNmkE5iYkUZdCJnaaRRrUNtBpJwYqKiqLEhASRWPj/YOqjgYEWxfaRZf6HFLJ5q0iySIOBbH16U/i8ORcEMnmY+HgKnTaFnHfeTiGw8RhslowMMs59lSJsVoqAyqqRRr8FOmsyqRgu/tNnzkitmb9kg8roDg+nRvv2kX3yE+TiX4fpdVTapTOFvTKbdHCO1BDhmUrY9rMsWsxqpIsiWdI6Jj5I5nF3SvjBVeTxfPp7Ht3e7BND5HmMhR/EVVIsoYqqejLdit3KEldnsZyfB5TH7y4tIRfPcU0UsmI+CBk4oVow/2JKtBoNWMOVX69uIqWln5Lk3ZIST2AaQWz4IhN5XTjYjhLvIEIBrMaFPvWPGgWZIj1DJ02g4iNHKXLXbgkDkHchWZcsJcOiJRI+IL8FzlyHjDw251/+TOb+fat0L2dyMtlnzCJdSQk5zwdoSC0zm8gVzs9ery4ZmzUjI6vPutatJO0s6P1YPXY8/wLpWFW31gDQTGA8HduTafKjpDdWbTnknM2hPXv20KGDBymDNQbUFWKKw5mRxrFd3KRJE2rfoQMlJSVVa0xpaam0a+cuOnzosCQZIGSEdLzYOrFUl02Blq1aUdu2bSk29vybM8ErvY+ZPcafkpxCuTk5EjqKjoqmevXrU9t27ahT504iGGqSajwFCy59tCaw67AIPEBD4FofFUGh6zeSnSewCNkjcOHf9xcyNLns4rAdtgMtf51I1vsnkN1kpIir+os0cCz9hkIYgO4gVd0mvZGKv19WdaAVF5Nt526KYozZzlM70HvZI8tWcvDLLrGYydS5M+lvGUumfkHuz1qDc89+Cisukue5YKChpIjfT1VYxOnTp2nhF1/Sih9/pFOnTgVIYuXf8EQDdFd270433jSWLr/88ioCLI0++ehjWvnLL5TDAEBJldtHcfL1Xjds2JAGDx5MN9x0IyUkJFbp+j8uX04LFy6k/fv2e2oifdQyJYECl09qnkQjRo2k0ddfLyGhSxJoqKAu5MWn13m4LZKInbDd+N8uXsAhRubi/EzWtm0oYvg1F1XGGzu0p1JesKaiQtInJpDzwEFy8ZHvsJMjCCgsfN65ZRu5TmWQvn69ygEN6cPSG4Wqtgtok4DZMrIGYNq0hRxbtzGD+AtZ7r5L5Yu8+FhzKCnIF4fPBdvXWGxVCHds2byZXp4zh6VAqoRJsOBdFdwfiQKrV62iDevX013jxtG48XdXqCKvXbOWXnl5Dp3OPC3XRdmUgxxBNZa01FT68MMPadXqVTT58cepW7duFY5l3ty59N033whbk+sHeVe49rFjx+j1V1+lTRs30WN/n0wNGjS44Hmu8Xq0Um8eoxK4Rg9GF0uUMOZWDub8NvJwDcv1o8Q+udhkYelgGD3Kozb89AuFs23jqAAQUGlD8vLIsW59rSr+WLIAao7TQU50CHvjv2T7edUlYZTs3r2bpjz1NKWcTGZJYK9woZ4zJV0iNbAe3nrzTfp8wefBQbZ2LT07bRplZmR6u6O5ykkyNRvc6b1+Mo/pGR7boUOHVL8LtXPGc8/T999+J5X8lVV6yLWlKsRJGzdskOfOOnPm0gIaHryosIjMRnNZvRnSriysRuh27yFTcYmIa2tsDBl79ayVRWLq2oXMQwaRmycc4HFVUhiOT5HyZWcOWxOqmTAb5pJmv8MQhLvjbDG8pQAeq2m8omrkfsEOfJeczqDXKmHtZN7Lr0gygkPle5AAcMrgUJNYCijnv/MOJSefDPg8JSWZZs96UaSOf42iUi7le6itubzcXHr9tdfKtSRU6IP33hfJKsXFfgCrbOz4zcEDB+hfr79eofS+qKqjklaFBjn4i8A18g0lA8NncegMegqB4X/4CNs/rPbw87jatyVd3cRa5coOBrruyJEqqVslUI127CInc0tDs6Y1IFYt5EapjtvlsTPQAYzVxFCejyLcSwVsGKedX7KDNQHjZY3Pz+3lc79K54WvZggPC/r5zz//Qvv371ddaFigYWFhlJCYKLWBGWy32b0SyR9sSIODenj7HU3Kffbu2+9QTna2KsiQvjd02FDq0qWLeLR//XUtq6IbAxzekG7btmyl9cxI+w/oX3YeIFmw4DPVsSvAatiooZg5sA/xPf/vYlw/LP+BBg8ZQv369699oKWmpknbAbx2JUMfRj3erTJY6b3IAzVZreROSxfbDN/RNW9e60m9rjW/UhijPNefq/lIsjL1kb8TyZzcsWbdBQFN731Gy8MPkaFzJx6EUzydblZL3JlnyLZwMVm2bpXaO3+wYdkZoYafPk3kBzR3ELCJnceL3zT+LjL06e25X2UeUB6LMTraY2uq2k5rVN+Vge/TvWdPuvf++8TLCOZ68OABenXuPDp69GjAgsWY4UX0pQMM4J9W/BQgKXE/JDdMfuIJGjFyRNn5628YQy/MnEnL2dZXGi8B7BZmLPiL+/sCbcniJWSz2gLGItfn30x4eCINHTpUvLzbt22juSy5MzMzy30fUhCSFE6Uvv36VXvdVhtoxajxYhWxtLQkaAMdDMpqZ57JgHSfzRE7BB2GDQ3q1yrI3PkFZGXQuFU4vI6Cdxmyscphuf1miTNVy/OpvKgWLcjYtk3g5Pe8kkom/Y3C9uyT4HoAcBBsrKRvihrYwBzU7ne+hFjovr17BVTlQMYLL5rB+dSUpyk+Pr7s/BVXXknj77mH7ZqngjglyrdxX8EgU7P3cL8BAweWA5nH3jfSQxMmUI8ePcUbCLDgiIgIpzA2TyJ9EhLOnj1L69evDyqJR40eTTfdfHPZuf4DBrCaXELPP/ecxBbdfhJ5146dIvUaVTOFz3gh3Nput1VoWEpQ2MAH9GOA0e3h1KYajlFUqh4xt9KnppA/O7Ag+bhzR3JlnyUzJLQPEFENoDtwkBx8wHt5QUAPwoj0aFbUuxcZ9u5XBakD3FolkFwZT62pFrF4f8OHDxf7yTc4DtulZevW5UCmUHSMp5OZfyI5KManagJOkh3bt6vaTTgz4KqrVMcUFxcn6mRlBLXxDGsDatcHOK8dfl3Ab/r27ydhg+TkZE8g3wdoECb79+2rfaBVeZG5AxdGbauNtlVrJO8xx0dF0YnTgI8R15Hr4EEyL0onq48GgyB3DEtj26/rLhhoVEFg2Xk2V9UjBWeJk6WGXuXFBlMd3aIxuMl59Djp6h8AIoKDSFJzQsnQMon/oT4+xMLue/CB83KGLfvu+7Jemv6soVOnjmVnYM+l8IJWkziQVghKi3rJgNm5Y4cExgF4CVzH1qHWDPSu3bqWk2K+dIzVV0X18xcQ9VmjUgMMnjeJzZqUlBTVa+7ds5euHjq0doHm2aDCrGpA+uq3EgTkydFZzDLfBj7lClJOc1FsMza0UUpT6jfh8LaV8pgiruhGTn6xjoVfBaiRCAM4Vq8lGncXv/3qhSKEqejU1Vkb2z/OH1eIMe8PGLRzoMu7kj4hvupMzWtf2t9+l3TzPyBdBbItHP01WyRRxP/9i3RhYdWe3xMnTlBqSiplZZ0R7x4cEv6L22QyUvMWzcs5E06zLYTmuWrzFc4az7EjR+jdt9+ilb+sLKuU1/kAF20GW7ZsSTeMvVHUQP90tFOnMlQZOs7FxycEBWi9evWCTq6a1/SiAy0Ui5MlhBR38uH2xjUwKQrw8DcE3YGRtY/CzRPJUjFtO5VZa0BzbtxMxtNnPM1+fNVGpIE1SxKJwYoOlYSHkakQQWcf9ZH/X3/kqHgsjVdefv4g91aAl/7fW2RbtETKecQZgh1ysDcAc3S4IKxQtXzVL7aBipDZcsdtQW2/CgUoP6veVbFtBzvIXUFz2qrSgs8+o2++WipBbFJRF2HPdezYkf42+TGx6xTKysr2fF/FQ4mskOlTp4qzLVjMDoCB0+WlF2eLU+XRyZPLnHIgdKxW4zM4FV1BCVaw1Cv4FpAZU+tAwy4wABV011IrXPxWctjdAfq5UW8gBwLTDeuTftsOWSZuiHV3YK7hRVEbf14lmfwlAQ/OXLFzJ0+shlUJXauWFMIGr81ZXqLBqrCuXlMtoCnv2bh3H5n2lZcvyJaxeiWQ7yxAkjnrxJLpb5PI2K5t0OtWNHO4dmX+Rsn/rCCnsuqmgcc2cFbgQodjA9KnnDOtpDioCWHndeVQAa3/fRXP4zdLv2GARNLESQ+XfQYnXbCgd0Xd0PSG4Gp+cVH19+SrdsAanAo5jDFs/NarmyCttpOSmlF8XB3RdZVJRE5hKaLxLVqQnf8fC8vBRqXz9JmLDjJncgq52OC2+nNZXhlWfgn6Lh6bAa5tI6tpatLCjmyC9RtZ1cuv9jgQE8vnOSjwORCUdvqMS+dVZ519+pDhv2+QedhQ+j2QAgY10ChpWv967TWaM3tOuc7TMD0quy7UQawzxY2vFrBW7vH5ggW0ZcsWn3PuCtZudZe9u/aBFuwBwiPCZUIVnRkB7FI2Yl0dO5CN1U04Ti3ZZ3nxbrj43saNmyiE7UF/tdGk15E9NoZM7c85OQzdulGx1wnhS1AfdWyHOFjaXXR7Eszh5Elyf7eMnBWk/VSmB0AtRts9SMdgh5k/P9/Qgeq9WFsxGoxlzFet9AgmxZLFi+iTjz8+N99GQyX+Iz0l1q1L995/P81+eQ5NnT6NevbspVoapPgJFi9cVAZQ2WshGAOuIBPG6Qj+2YUkGNe41xG2G+YaOY7KTp0hDjOVNGpI4V06kXHzVtEaS7/+hszXXoOy64vm7nSsXE0WFZcAUo/cbduWSxo2tG9HuiZNyHTiJKuMznJqWAQvIOsvq8nUv1/1FqOOubHPW3eKpHSRy49XYpGYU1PJ9N6HVLx2LYXMnkX6hg3OS3XEMnQiUyM6quIUMrTgS0oKGqiuKt122+00aNBgcddv27aNFi1cKC0tfNU+qfNjCbPoy4V03fDhlMjjiwgLFxtWLQyAc3Xr1aPZc14qp3IOZBV0ytNTJIge0FLe6RLv5BlmULh+VFSkqmkiWTclJUGfpzBIB22MqU5c3KUDNLhfQ0PDhItJ6AwBWz5fwBMdweqQjdUwxIfMu/eSdfkPZBk14uKojXBi7NqlGqTGazV27aywQ88Chau7c0cynTypYs+wdNy8hdxZWaSLjz/vsbhYtWZ92nMvxBJZDTXk5DIAKSAlDPE7HHWOHqeSDz6i0H8+fl7qiRnc/P57yAgmVoF7H+/FApBVkNitOCJ8pZTyb0WywFVe35uA0LNXT6pfvz7NnTNH0vH8nQlnzmTRwQMHBQh12MSgIDYYrj1s2LAAuw4blYy9aSz9uvZXHkN5JwnsMdSuwTUv168TF/T6eXm5Ac+lENLBdKTuraxXr/6lAzRMMDxJvhwHHEQanvbuTUaWHGGHDpMD33nrXXJ07ULGxjXfMMe+dh1FWO2Ur8LVi2BEL/1WYmRlCbWsSrkyM8muksQrrez4M9uWrWS5ZljVVWlv8NXy14fJ2P1Kz4uH/VBcQs7de8j27nsUevyEav4l7Forjy8kN490MdHnpToCPDrkOppM1VY/AbLX5s2jnTt3ltsBCOehrTz+xOMSc/Knfv360vx336EsBlVgKpOecnI9G0GigBP2PbyDvoBRJBzCAarud9ZCsD1XcbEjQIPBb/O9tnSjyxp7JKafZYVzp0+fkdzLSL+2G3DCpLBGEcyJ0qJli0sDaEj8zEQLOJe7nHgG6MJYfcrmB2gwfhzZnniSVTKiMF681lkvkf7FGaSvyWwRuIV//bXMva5mC4WcTCZTckr5HEeXW5wkuoDvu8nE4y+FKnoeQFMWgD46ho9zYNHHxpKRVUJD/bpkffgxMhcH1rKJV/BsDjmZQxv9gFaZ11FnvPDXikV7kufo2JFj5HA6Aj47fPiwKtBkR1cVO0fnHbjF7JGgdRkwjRo3LgOGvxMkmApXWlIaMB7voDw8xuvib9mihVT8l5ZYy2k1uP6p9HR+tpPUoWPHcpc4wedOsj2ulhsJQvX1b+4MQR7bqYxMuaRaJywpg8AOnd2vIMONYyiSB1+AchC22azPziRXXn6N4QxpU449+yrM1IeUgvev0OdADC3YAkbA27ZpCznT0s7fXLSrbwFs5BfnvKyRqHuq6i1Ut/MM7uN3zpQ0crGtCfW5SsfhI+iCG7C4kNmOwLCyNZZyYLF+++23UrHhT6j7QhwswDWPdDy9jhLiE7yAsFDXbt2CuvhXLF+umkO7jhmow+5QdQAijqakhTVj+zMpqYW3nXygcwaOE39azPYlBERANgm/HwSymzdv/ttJNAwKwMpgcQz7DP0jVFURXsTYDSYvJ59C7hlP7mPHKWbHLsrlFxe5chXZ+BroW2FIanrhQFu9lsJ5keYFUQGqkysIN384Mwr7uo1kuOmGmuEIrJJB0gVLpxKbzl71rYCleJR/Y3v7HdJ/+HGV6unQJwXlS6YZ0wNqBNu1ay/ODX/1C2DbsmkzPf3PJ2nMDWNkEcIZghYECz5b4G1ypOLgYNupSbNzZTKDBg+izz/7TEpsfHcBwvW3b99BM55/nu655x7xPpYyANatW0cff/yJamEoSl0aNGwohwK6AQOvov379wXOE1//+++/l65o11x7rXQB+PHHFfT1V18HLanp07fvBfURuWCgwZWffipTHqwwv7DCVVzKHCqCF9cZl5HqT3mSSv/5NMUwN5Xejlu2UemESWRi1dI44lrV7sRBFxjUlZ9+IX3TJqRvniQFnvogWQEIXofo9UHhhheITHqHXyAZrxY2FwpC3WNG1Yh6Ji8RLvFKVLjzYRQiCVmFN5RWzXWvl3CHHsZ1wGc9evYQBwDyEtUyODbwPEPChIeHidmgOE/UAs2QCr15scb5eO6Qr4g6r+XLlsk79GfgK374UQANDySAhnQv5bOAOWGw9O7TuxwY4FBZ+MUX4on09Rm4vSr9F59/TosXLZL8R/EtqIxdau5Cw2j06NEX5iSsCS8jjFN3BYaDYuDiL1RIdA3OCAulujOepdKZL1LUth3Sc9+QlUXul16h0qXfkn7Y1WTqcSXpmzUNuqhdrGvbt24n3bIfyHrkCIV//D45du8l3dHjohoGxEGQrd+8GVmlVssV6JXCXtpoerP8RzLnF3jsJB8qcbPqtH27qGWGFs1rCGjnr71XZqNJjmYVg6uyYQhsJ33gFQGKO+68g16Z87KqGx5qpeedFgcFgDiFmJmgce1td9wesC7+ct9fxOECu6kcGLyLPjc3t5zDRO0ecn0e641jx5Y7n1g3Uerl0MrAf/zK9ZWq7GBpXpCUt915OzW/AEdIjQAND5kQH0cpqelk4oVq93ImJWiIAw8BfVvZ5CIfHh/+TiZzirjnp5PhvQ/JtGgxGXiiYVdZDh4kHR9FzCl1TZuSsUkT0sXXkZZx6LmvO5tLzvQ0crB9Yc7OoVCTkZyDrpIGqtb3P5KNCXMpsMATi8kxdgyF3DS2YnvzTDaFrF5DeX6xGjhLoqw2cqxdV2NAI0vwIKjb5aqxkpcqQF717JgbbqCDBw9JYxu1BPIKy6RQvW00SK3Yk1OeVm1yU79+A5ryzDNSw4YaMv+eHpW1EJD9+CwWmjxZvYnO8BEj6DibKZ9+8olqQyF3BSEGHEOGXE1/uuuuCw971cQrQspVZESYtLeAMYkHRwAR+Y9mJOtiB09eoCh7x84zABt6UMCOOs1AjH7gXoq6ohu5PvqUDDt2SNwNzgdjYZHkCer5OLczmucvvqMXKemxZQyDBrI0Kib7ip8plM+F+jkYEKQu5nOhHTpUPindupBhDatErH/6v2ZspFHK9zDfMpZ0qBVj8KGwFfcz+Kd6KVy0gr4fhphoqYoO9VdZhDEYqHT/ATL17X2u1AZeOWZYofqa8WMpqqMrSLYEFtvkv0+WXXg+Y3sKjoiKVESlMa3SKq5Dxw704IQJ1Llz56Bj6NylM82Z+wq9Nu9VqXQuYzJBvMZlcTx+z40aNqKHJk6k/lcNCHp9FIsiY+mjDz4Uhu/psqXOxKBG6r0ZLqgMeODBB0VruySApqgZJ1M8MQjES6Kxcycy0L0BUWT4Z2eflWB2iTehFE4UfJ7PUrC4cyeKat+OwrdsIdeKX8iNosCcXFEb3TrPpLp9uRC4E790PeyDKy4nS4/u4m00sHQrrlc3ACCoNEALBUNSs8onpfsVVNSYjWq0YPBb0PCzGZiBuFLTyMDqBMp/9E0uIxurT05/D5diF1bQk8PQri3Z2NZx+rioRS3zYIrsa38l88jrSK9wa34+atSIrNlZVW7YWqF96/YwE1NY8ERb2N8PPPQQ9ezVi75asoS2bt0mmez+uYcK+NAWvn379jRoyGC6auBAYcSVUatWreill+dIa4MVK1bQ3j17eH1gSy69f7REpqlFixZs8/WR3otBS1sU8PB6HP/nP9PlV1wh3sbNmzbxWswOqBxHWCo8Mpy6XX45XX/9GOrVu1fN6Qs12RK8WFzEOulSrGbEn83JoZycPAkM2n2MX9lal6UdchJDoqMpgp/fkn6KLMeOkfvQYXKlnUI435PpgPIONniRPqVLSqJSBo6xdSuKYMngRk4lS1SdmrsckifEUuV4nZNtA0gi1bWMMhTUM6EJDlSd/HzhkOquauaQaDEeLHjM0smZlx/ghi6zxZAqFR3Dq92kuMzIhW2Gg42tWp5jkjHqqridVWZmBqtjJ6SjMNQ9SDmAEXZSfX4vjRs3lq6/hmqmd0F9TE1NpeSTJymd3z1sNKfTIUnsdZmJNmrUmJo1a0YRkdXzAsIePMprK42ZZW7OWak8QOkMrtu0WVPpgXLeOnZtAq0qlJaGTklOVh0D42bgMBYkbvILQpmEkQEL9SuE72R0u2RRA2hOPqyo4kXzH1ZJ46MiKLaW2yNo9D9Lv33v/apwqlJradCxgrMUewOnohejdRnbKUWMaST3urybFcoBlcZkJntxEUUkxGmvX6NLmmoVaMVoYupwVlimUGZTMcgcfuqlv6fIEBUtdqCpgpw+jTS6FEhfq0ArLpGaHjWgwYsUHh4hvRw8ffoMAeDy9XTh+/BqRlTB0NZIo/8ZoMGlWpBfyLZ/YEoR1EQAC7vQIBTgSSxFY9bgJRzKJvTh4WHaW9RIUx0VQjY2XP92v6peT1vpcAaZQTxWIKnKLrVSTm4uSzW3xOD8nTZS2Ws0VFltdPA1DejG5SMl4aLGvxMSEsrOIWcPmxog7ccTbCfxeMFRA2+mb4dc2ZDB66Ax+41DiTUpnjdFiuPfSuBXqUiGWzk9PY3i6sRRKKvCSj8M5IfCW6oE/yujst/BA4rfeOM/yv2UMfhez/d5FEJ33zNnTsscgAlKCIC/p4wXz63Em6o6Jqc3E0fpoa92TumPj8PofVdKNj/+H2sH1SEoscH4fcelmBtKQNp3vvD7c3NuFG8t7o3UMmSswJuJ+Bp+o8yDMo81ZZbUGtBcTo/7G7loiJ8pLx2qIlzjdes2KCet5AF5QtLTM+ThfdVN/BtgjKriVrlr1qyh+W+/Q3369qF77r1XrvXuu+/STz/8KJM5fOQIGnf33bR3717Z0AHu5HgG39PPTBFVd/ozU+nOP91JV/boQa/OnSv7cY0cPUp2QDmbfVZidIjrIN0HDTgRjH922nTJz0NGBIoh0S1q06bNNPOFmXTk8GH6zxv/oenPPSvPOfP5GXTi+HFps/bU00+L23r2C7PEdY7F1K1bV/rbY49V+pwo8Zjz0ktUWFAopSTt23eghyZOkNqrWS+8wBpFgTwv9i17+JFJlH4qnZ6f9qxs6PfPp5+Sd4MskDmzZ8sGfcjenzJ1Kp1KP0X/+fe/6Smej9iYWHp2+jSpdh56TeXbbqFv/7xX5gp4kbKFhqyIU82bN4+KigolTHH1sKGyh9rcV16R94r3g3QqxOPQJ3LGrBekoHPWzJk8JzmSATJl6jM899n0xr/e4LE/SVu3bKEFn36GiLOsj4kPP0zHjx6ldevW00z+PZqfvj//fXpm+lQZFxKWU5NTpCbtGX5GJB8v+34ZPTfjeRnr0089SbfceqvEAX9XqmNMTDRPUD1RCcGB8FIRuwAXUs77E7yKmDT/yDxUSiyYyMiq2WfYXOEQLyBs2ADpcZwX9QfvzqdhvFCGDhtG77z5lsRtsMkeqnQnP/G41Fuh9B7dmDdu2uhpj8a0d+8+aXOG3UzXr99Abdu1larfVSt/kU30QEePHKXNGzdJsxi8YBC2F9q5bZtkoMOzirJ7cNGlXy+lbVu3CiCwK8p8ZgBY5Js3b5aMidtuv03aVftTgUr5TD4DadvWbdKae+TIUfTtN9/Q1199JcncW7ZspS5du9CtfL0+3k0Od+/YJVnyq1etlnbXoPfmz5dnvf/BB2VcSxYvFqAiY0PRRtCyIJ3BV5UxITC8d/du6jegP916223UhZkGpOUOvkavXr3oFj6HQHJ6Wro0KEXK0628wDt06EDH+D1twd4EPP6PP/xI9k5Dv/zt27fRkkWLKT8vn7Zs3iTdqY4cPiJNVm8ae5P8vnGjRvIOMe5NGzYKc9/E7xEay5IlS/jZd8ozopvxO2+/LQ1bN23cIMFsBMwP7Nt/QWUxv6nXMYxFdFjDUJYY+aIWAmQNG9QPKp7ReBMA1OkM5WqTcN5o0FcpNQYTm5qcLNul4qWn84IKDwsT1SovP4/+NO4uUR9kp1I+MCZsfP8serBL7ZJTslkUNQkJ1Pge+AL+f+CgwdS9R3f6YsHnZcWuyGqoj0wO/g6ANmjwYFF1oa5gUV89dIineQwKKA8elEU1cNAguQdABgkZwnPVt39/6s7Sx18lRZObpV99LYvjr489SvHejHil3fU1115DLVq2pM8+/VQKJcGsQvh8Pwas7+6bu/fspvYd2ksj06O8UJFNn8aSoy9L/iFXD5FYZyzPjdOr5lpl16ASCrWEkNHnnaGC4/3579GKH36gjp0606S/TWJp5ClWlRxT1goG8fM182blYD+0iPAIkYqt23j2CMDea9Burr5mmJTTgFavXi1ZJXj3aNONzBT85lGW7nVkXE5JdMAaQcoXJPBI7154Hs3HIO8QTVgHXDWANaAoKUo9fvQYdejUkQYPGSwdZfL4+QEqMF4AGpn8o8ZcT40vu+z3CTSFoqOjxC0P6WKpoDmPslOIw6diV+ft+V8nJqFK94IenpqawpxwknDrXbt20XXDrxPOiFJ9cMJnpk2VPhN33HknHTp0kO6+a5yoaqOvH+3pw+72r7p2e2wg1vfRSg06PrLMb7zpRvkU98ACjuLnhCT1/MIt2QzpLDkP7j8gC0jp6V7W/9AruaDCQp1+ceYLsqCGDB3Kaswt8hlaXb/53zfLtklq3aZ1WdIrUt9g00EVxYChdqLHPFRYXG/WjJkyJqQtobvvHh5n/wFXicq5c+cOum7E8HJJvUgoBq1auZLsfB4qNJgbVFqTTxcrMBO0L8DCTucxdeRFPJoXquLoAgN68h//EOZx7333esBjs0qjHfw/pHZkVKRUdjwy8WFR1x948AHRepSNQiTU4zeuNQxEZayWEAvPzTEad+ef5Jp4p2DISTznqWk858zQxP4lT8a+YosNGjKk7Dlw3Qfuu19+f8ONN/5+42j+XsOqELpqIZNEialhUeOFhFaQm+dLhxlIsKOOHzsqEu0wA8nFHP/6MWOofr36Yhc89MCD9O833qAGDRvQc6y7v/3WW/TSrFlyT9h1SqKsh0PrzoUb+C8KDePi6rA6ekxaY6NYMhUNYthgR04naqjsNk+tU3xiAnPrurLdLBiI229jAqiUvtnlUBmhmjZt2rTcvEE6ATwAUohP9r9cjrk7+mWggjkleaUwikTs8czjRyEkbElIEQ8DSmMVMJ0yMjOl3N+zmcW5sApsVbNX8mJZjhw1kgERLQzK4ZMoDQmPlnNu7xjMPswTYzTygh/KkgpSHszmJM8PpPcQXuSXNWkikhTb2QKUwxnsqJLGebw7Jf/I7dNwFyqjOC588iCRBoYN32+65WYBCiQqmFhC3URJDdu0ceO5OfchPDNsMmxE36FDR2rO84N3V0ka1u8HaFUl9H0oOpvHc6wnZdYhEasK1H3MbbEjC4x8LBp4GqHaoZU0+gXOeulFGj9uPO1gm2kvq1IGo4mmTp8uttrSr7+W4keic95DeAKdTm83KwbFyFGjBIyoqcIWtF26dhV7ByoNjP3cnFw6cfKELEQshn79+9EPy5eL9wsLRQpmCz39Md5juxEAGMuLBfeD+tber9qgCYMO0hbFklDFfHdFURbRrbfdKoCHfQJpU29QfZHA2Fxd6Xuxku1VeHYzMk4JwNEiDvaKyWzytPdmggQEaK9kmw/zDfUWnsh3mBH5Ltg2bdvSXx99VHrvt2HGgO+dc4I5WfqZpM0cvIViw7JUhscYSccAPugAS3nZ5eW668raEUBzASOChxPzpzAh7JGGfEo4dRSmAOdPQt0EGjFyZNm9IbmgDUFTWL5suZguBmYkKN1BJ2PQW6wdwNnz4pzZknwMxh5ZRSfbHwpoWLDSUx0TKln7lTff9J1oqBfo/f7y3Ffo7Tffoi+//EI8WthwDnYFmsTg+nhx6OG+lA3lVq1a8gLMpHZt21Cd2Fhx2nz33XfSxnrf3n3SyBP2AVQiZJiL2sYvEA4SOBDAIac+O13c67eMvVnO4fvZDN7OXbqwhGomXi5IHzSI+ffr/6aFX35Jy5Z9T/369ROGgGtj3+eW27aLWjvUpynQtdddy/bEsIBmorDtwMXhdVSAUAwQo0U2Xw8VxeilAa4NRw3U9zks0QHGiQ9NkA5QAA1aEsBmgeSdcMUVssAxl3AohPMBYPpuY6uUlEBdVMvoB6A/fP8DkRTI6pf6RJ6rIp/eMi5v+h3uoQANGgXCQmgz15LfCRxOCxe2EYfShIkT5DqYJ6eXMaQkp8gOomAKKJsBBHNYzYWzBe8XdijG16Z1G9FaMOc//bSCevfpUzYO2KClFfR9rC4Zpk2bVt3fTqsNoJXtY+xtCa3WDTdo7I4XXDqrRzB6scCgUkDtgEHdsXNH2TnyyKHDdN8D91Pfvn2pDds7UKO+XfqNAGDiI49IS7Qk/u2GDetpDXPsgYMG0t1/Hi8qJLxs3a/sLpW8kErwrIIxYEsgOEBgY0A1SYhPpIYNGsq10PuwTnyccGiU3oMJYHFBerZt244e+esj4jzA76zoKpaRIapXt27dAuZFLVaIhdeDGQFUody8XIkRoj8iFjEcI5DosIsRQmjfvp14KMHBYWNijgBiqLvYTnbg4EF09/jxsn0yAAzJjeeDwwb1Zf7OArUuwtj4vqi4SDyDuDfG1axZkmgD3bv3KNtwAk4rPBPmJ8zbxgJzAMdV3779qH3HDqKS/8CSfODAQfIOXCyt4LhAOQtUXgATMdDsrGxqwcCM4mvX4fvhGePi4sWLjWfo2LmTOHrg8YUj6uFJk8ruicZCzZu3YI0gaMer6dVaxxeQva+RRhpdanE0jTTSgKaRRhppQNNIIw1oGmmkkQY0jTTSgKaRRhrQNNJIIw1oGmmkAU0jjTTypf8XYACnz1OtSWpGJwAAAABJRU5ErkJggg==' style='width:70%;'/></td>		
			                        </tr>			                        
		                        </table>
	                        </td>
                        </tr>
                        </table>
                        </td></tr><tr><td style='border:0;'>";

                        string body = @"<table style='margin-top:15px;'>
                          <tr>
                        <td style='border:0;' colspan='4'></td>";

                        DataTable distinctClaimType = Catdt.AsEnumerable()
                             .GroupBy(r => new
                             {
                                 ClaimType = r["ClaimType"],
                                 TypeIndex = r["TypeIndex"]
                             })
                             .Select(x => new TypewiseCat
                             {
                                 ClaimType = (string)x.Key.ClaimType,
                                 TypeIndex = (int)x.Key.TypeIndex,
                                 TypewiseCatCount = x.Count()
                             }).OrderBy(x => x.TypeIndex).PropertiesToDataTable<TypewiseCat>();

                        for (int i = 0; i < distinctClaimType.Rows.Count; i++)
                        {
                            body = body + @" <td colspan = '" + distinctClaimType.Rows[i]["TypewiseCatCount"] + @"' style = 'background:#0070c0;color:#fff;text-align:center;font-weight:bold;'>" + distinctClaimType.Rows[i]["ClaimType"] + @"</td>";
                        }
                        body = body + @"</tr><tr>
                        <td style='background:#ffff00;font-weight:bold;'>Activity Code</td>
                        <td style='background:#ffff00;font-weight:bold;'>Product Category</td>
                        <td style='background:#ffff00;font-weight:bold;'>Call Sub Type</td>
                        <td style='background:#ffff00;font-weight:bold;'>Total Claimable Calls</td>";

                        DataView ClaimTypeview = new DataView(Catdt);
                        ClaimTypeview.Sort = "TypeIndex ASC";
                        Catdt = ClaimTypeview.ToTable();
                        for (int i = 0; i < Catdt.Rows.Count; i++)
                        {
                            body = body + @" <td style='font-weight:bold;'>" + Catdt.Rows[i]["ClaimCategory"] + @"</td>";
                        }
                        body = body + @"<td style='background:#ffff00;font-weight:bold;'>Total Amount</td></tr>";

                        string[] strCol = new string[] { "ActivityCode", "ProductCategory", "CallSubType" };
                        DataView distinctview = new DataView(dt);
                        distinctview.Sort = "ActivityCode ASC";
                        DataTable distinctValues = distinctview.ToTable(true, strCol);

                        for (int j = 0; j < distinctValues.Rows.Count; j++)
                        {
                            string tdvalues = setvalue(dt, distinctValues.Rows[j], Catdt);
                            body = body + @"<tr>
                            <td>" + distinctValues.Rows[j]["ActivityCode"] + @"</td>
                            <td>" + distinctValues.Rows[j]["ProductCategory"] + @"</td>
                            <td>" + distinctValues.Rows[j]["CallSubType"] + @"</td>";
                            body = body + tdvalues;
                        }

                        string strbottom = setBottomvalue(dt, Catdt);
                        body = body + strbottom;
                        string BlankRow = "<tr><td colspan='" + (Catdt.Rows.Count + 5) + @"' style='border: none;'>&nbsp;&nbsp;</td></tr>";
                        body = body + BlankRow;
                        string footer = @"</table></td></tr></table>
                                </ div >
                                </ body >
                                </ html > ";
                        #endregion

                        string html = header + body + footer;

                        string _fileName = "ClaimInvoice.pdf";
                        Console.WriteLine("Invoice against Claim : " + ClaimInvoiceCollection.Entities[0].GetAttributeValue<string>("msdyn_name"));

                        Console.WriteLine("Invoice is created.");
                        if (html != null)
                        {
                            AnnextureURL = HtmltoPdf(_fileName, html);

                            // create an attachment 
                            createAttachment(_service, performaInvoice, AnnextureURL);

                        }


                    }

                }
                return AnnextureURL;
            }
            catch  {
                //_errorMsg.AppendLine("ClaimInvoice: " + ClaimInvoiceCollection.Entities[0].GetAttributeValue<string>("msdyn_name") + "," + ex.Message);

                return AnnextureURL;
            }
        }

        public static string setvalue(DataTable dt, DataRow dr, DataTable Catdt)
        {
            bool IsGenTd;
            string body = "";
            int TotalClaimableCalls = 0;
            string result = "";
            string[] TDvalues = new string[Catdt.Rows.Count];
            int k = 0;

            for (int j = 0; j < Catdt.Rows.Count; j++)
            {
                TotalClaimableCalls = 0;
                IsGenTd = true;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dr["ActivityCode"].ToString() == dt.Rows[i]["ActivityCode"].ToString() && dr["ProductCategory"].ToString() == dt.Rows[i]["ProductCategory"].ToString() && dr["CallSubType"].ToString() == dt.Rows[i]["CallSubType"].ToString())
                    {
                        TotalClaimableCalls += 1;

                        if (dt.Rows[i]["ClaimType"].ToString() == Catdt.Rows[j]["ClaimType"].ToString() && dt.Rows[i]["ClaimCategory"].ToString() == Catdt.Rows[j]["ClaimCategory"].ToString())
                        {
                            body = body + "<td>" + dt.Rows[i]["ClaimAmount"].ToString() + "</td>";
                            TDvalues[k] = dt.Rows[i]["ClaimAmount"].ToString();
                            k++;
                            IsGenTd = false;
                        }
                    }
                }
                if (IsGenTd)
                {
                    body = body + "<td></td>";
                }
            }

            var a = Array.ConvertAll(TDvalues.Take(TDvalues.Length - 1).ToArray(), s => decimal.TryParse(s, out var i) ? i : 0);
            result = "<td>" + (TotalClaimableCalls == 0 ? "" : TotalClaimableCalls.ToString()) + "</td>" + body + "<td style='font-size:13px; font-weight:bold;width: 80px;'>&#8377; " + (a.Sum() == 0 ? "" : a.Sum().ToString("0.00").ToString()) + "</td></tr>";
            return result;
        }

        public static string setBottomvalue(DataTable dt, DataTable Catdt)
        {
            bool IsBottomGenTd;
            int OverAllTotal = 0;
            string BottomTotal = "<tr><td colspan='3' style='border: none;'></td><td style='font-size:13px; font-weight:bold;'>" + dt.Rows.Count.ToString() + "</td>";

            DataTable BottomTotaldt = dt.AsEnumerable()
              .GroupBy(r => new
              {
                  ClaimType = r["ClaimType"],
                  ClaimCategory = r["ClaimCategory"]
              })
              .Select(x => new ClaimLine
              {
                  ClaimType = x.Key.ClaimType,
                  ClaimCategory = x.Key.ClaimCategory,
                  ClaimAmount = x.Sum(z => z.Field<decimal>("ClaimAmount"))
              }).PropertiesToDataTable<ClaimLine>();

            for (int j = 0; j < Catdt.Rows.Count; j++)
            {
                IsBottomGenTd = true;
                for (int l = 0; l < BottomTotaldt.Rows.Count; l++)
                {
                    if (BottomTotaldt.Rows[l]["ClaimType"].ToString() == Catdt.Rows[j]["ClaimType"].ToString() && BottomTotaldt.Rows[l]["ClaimCategory"].ToString() == Catdt.Rows[j]["ClaimCategory"].ToString())
                    {
                        BottomTotal = BottomTotal + "<td style='font-size:13px; font-weight:bold;'>" + BottomTotaldt.Rows[l]["ClaimAmount"].ToString() + "</td>";
                        OverAllTotal = OverAllTotal + Convert.ToInt32(BottomTotaldt.Rows[l]["ClaimAmount"]);
                        IsBottomGenTd = false;
                    }
                }
                if (IsBottomGenTd)
                {
                    BottomTotal = BottomTotal + "<td></td>";
                }
            }
            BottomTotal = BottomTotal + "<td style='font-size:13px; font-weight:bold;width: 80px;'>&#8377; " + OverAllTotal.ToString("0.00") + "</td></tr>";

            return BottomTotal;
        }

        public static string HtmltoPdf(string _fileName, string htmlBody)
        {
            string blobUrl = null;
            try
            {
                GlobalConfig gc = new GlobalConfig();
                gc.SetMargins(new System.Drawing.Printing.Margins(0, 0, 15, 15))
                    .SetDocumentTitle("Annexure to Invoice - Franchisee Claim Summary for Branch Commercial")
                    .SetPaperSize(PaperKind.A4Rotated);


                IPechkin pechkin = new SynchronizedPechkin(gc);
                ObjectConfig configuration = new ObjectConfig();
                configuration
                .SetPrintBackground(true)
                .SetAllowLocalContent(true);
                byte[] pdfContent = pechkin.Convert(configuration, htmlBody);
                Console.WriteLine("Document is Created and Uploading in Blob ");
                blobUrl = Upload(_fileName, pdfContent, "warrantydocs");
            }
            catch (Exception ex)
            {
                Console.WriteLine("error");
            }

            return blobUrl;
        }
        static string Upload(string fileName, byte[] fileContent, string containerName)
        {
            Console.WriteLine("Upload to Blob Started");
            string _blobURI = string.Empty;
            try
            {
                //byte[] fileContent = Convert.FromBase64String(noteBody);
                string ConnectionSting = "DefaultEndpointsProtocol=https;AccountName=d365storagesa;AccountKey=6zw5Lx4X+zHrh+CAKLdgaWSqVZ1zC7AKugPkNEGevep6qhh1xRm5Q3DBWKGJ+DsWfCZk59BbLSJvja81DD4++w==;EndpointSuffix=core.windows.net";
                // create object of storage account
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(ConnectionSting);

                // create client of storage account
                CloudBlobClient client = storageAccount.CreateCloudBlobClient();

                // create the reference of your storage account
                CloudBlobContainer container = client.GetContainerReference(ToURLSlug(containerName));

                // check if the container exists or not in your account
                var isCreated = container.CreateIfNotExists();

                // set the permission to blob type
                container.SetPermissionsAsync(new BlobContainerPermissions
                { PublicAccess = BlobContainerPublicAccessType.Blob });

                Console.WriteLine("*1");
                // create the memory steam which will be uploaded
                using (MemoryStream memoryStream = new MemoryStream(fileContent))
                {
                    // set the memory stream position to starting
                    memoryStream.Position = 0;

                    Console.WriteLine("*2");

                    // create the object of blob which will be created
                    // Test-log.txt is the name of the blob, pass your desired name
                    CloudBlockBlob blob = container.GetBlockBlobReference(fileName);

                    Console.WriteLine("*3");

                    // get the mime type of the file
                    string mimeType = "application/unknown";
                    string ext = (fileName.Contains(".")) ?
                                System.IO.Path.GetExtension(fileName).ToLower() : "." + fileName;
                    Microsoft.Win32.RegistryKey regKey =
                                Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext);
                    if (regKey != null && regKey.GetValue("Content Type") != null)
                        mimeType = regKey.GetValue("Content Type").ToString();

                    Console.WriteLine("*4");

                    // set the memory stream position to zero again
                    // this is one of the important stage, If you miss this step, 
                    // your blob will be created but size will be 0 byte
                    memoryStream.ToArray();
                    memoryStream.Seek(0, SeekOrigin.Begin);

                    Console.WriteLine("*5");

                    // set the mime type
                    blob.Properties.ContentType = mimeType;

                    Console.WriteLine("*6");

                    // upload the stream to blob
                    blob.UploadFromStream(memoryStream);
                    _blobURI = blob.Uri.AbsoluteUri;
                }

                Console.WriteLine("*7");
            }
            catch (Exception ex)
            {
                throw new Exception("JobSheet.JobSheet.Upload: " + ex.Message);
            }
            Console.WriteLine("file uploaded");
            return _blobURI;
        }
        static string ToURLSlug(string s)
        {
            return Regex.Replace(s, @"[^a-z0-9]+", "-", RegexOptions.IgnoreCase)
                .Trim(new char[] { '-' })
                .ToLower();
        }
        public static void createAttachment(IOrganizationService service, EntityReference regardingEntity, string URL)
        {
            Console.WriteLine("Attachment Creation started");
            Entity _attachment = new Entity("hil_attachment");
            _attachment["subject"] = "Performa Invoice " + regardingEntity.Name;
            _attachment["hil_docurl"] = URL;
            _attachment["hil_documenttype"] = new EntityReference("hil_attachmentdocumenttype", new Guid("5b078e41-ce00-ed11-82e6-6045bdac5e78")); // Performa Invoice
            _attachment["regardingobjectid"] = regardingEntity;
            service.Create(_attachment);
            Console.WriteLine("Attachment created");
        }
        public static string findAttachment(IOrganizationService service, EntityReference PerformaInvoice)
        {
            string DocUrl = null;
            try
            {
                string fetcheqline = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='hil_attachment'>
                                            <attribute name='activityid' />
                                            <attribute name='subject' />
                                            <attribute name='hil_docurl' />
                                            <order attribute='subject' descending='false' />
                                            <filter type='and'>
                                              <condition attribute='hil_documenttype' operator='eq' value='{5B078E41-CE00-ED11-82E6-6045BDAC5E78}' />
                                              <condition attribute='regardingobjectid' operator='eq' value='{" + PerformaInvoice.Id + @"}' />
                                            </filter>
                                          </entity>
                                        </fetch>";

                EntityCollection _DocColl = service.RetrieveMultiple(new FetchExpression(fetcheqline));
                if (_DocColl.Entities.Count > 0)
                {
                    DocUrl = (string)_DocColl[0]["hil_docurl"];
                }

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return DocUrl;
        }
        #endregion suppportingMethod
    }



    [DataContract]
    public class AnnextureResponse
    {
        [DataMember]
        public string URL { get; set; }
        [DataMember]
        public bool ResultStatus { get; set; }
        [DataMember]
        public string ResultMessage { get; set; }
    }
    [DataContract]
    public class ClaimLine
    {
        [DataMember]
        public object ActivityCode { get; set; }
        [DataMember]
        public object ProductCategory { get; set; }
        [DataMember]
        public object CallSubType { get; set; }
        [DataMember]
        public object ClaimType { get; set; }
        [DataMember]
        public object ClaimCategory { get; set; }
        [DataMember]
        public decimal ClaimAmount { get; set; }
    }
    [DataContract]
    public class TypewiseCat
    {
        [DataMember]
        public string ClaimType { get; set; }
        [DataMember]
        public int TypeIndex { get; set; }
        [DataMember]
        public int TypewiseCatCount { get; set; }
    }
}
