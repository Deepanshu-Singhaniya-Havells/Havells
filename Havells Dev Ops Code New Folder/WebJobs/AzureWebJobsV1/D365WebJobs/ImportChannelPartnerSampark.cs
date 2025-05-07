using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using excel= Microsoft.Office.Interop.Excel;

namespace D365WebJobs
{
    class ImportChannelPartnerSampark
    {
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

        #region Global Varialble declaration
        static IOrganizationService _service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        #endregion
        static void Main(string[] args)
        {
            _service = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));
            if (((Microsoft.Xrm.Tooling.Connector.CrmServiceClient)_service).IsReady)
            {
                //ChannelPartnerDataUpdatesFromSampark();
                //ClosePendingWorkDoneJobs();
                //MG_HSN_ProductUpdateExcel();
                //MG_HSN_ProductUpdateTempTable();
                //SparePartFamilyUpdate();
                //UpdateSparePartFamily();
                DeleteDuplicateJobs();
            }
        }
        static void DeleteDuplicateJobs()
        {
            string filePath = @"C:\Kuldeep khare\PMSJobs.xlsx";
            string conn = string.Empty;
            Application excelApp = new Application();
            if (excelApp != null)
            {
                Workbook excelWorkbook = excelApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet excelWorksheet = (Worksheet)excelWorkbook.Sheets[1];
                Range excelRange = excelWorksheet.UsedRange;
                Range range;

                string JobId = string.Empty;
                string _fetchXML = string.Empty;

                for (int i = 2; i <= excelRange.Rows.Count; i++)
                {
                    try
                    {
                        range = (excelWorksheet.Cells[i, 1] as Range);
                        JobId = range.Value.ToString().Replace("'","");
                        _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='msdyn_workorder'>
                        <attribute name='msdyn_workorderid' />
                        <filter type='and'>
                            <condition attribute='msdyn_name' operator='eq' value='" + JobId + @"' />
                        </filter>
                        </entity>
                        </fetch>";

                        EntityCollection entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (entcoll.Entities.Count > 0)
                        {
                            _service.Delete("msdyn_workorder", entcoll.Entities[0].Id);
                        }
                        Console.WriteLine("Processed.." + JobId + " :: " + i.ToString());
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
        }
        static void UpdateSparePartFamily()
        {
            string filePath = @"C:\Kuldeep khare\Purifier.xlsx";
            string conn = string.Empty;
            Application excelApp = new Application();
            if (excelApp != null)
            {
                Workbook excelWorkbook = excelApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet excelWorksheet = (Worksheet)excelWorkbook.Sheets[1];
                Range excelRange = excelWorksheet.UsedRange;
                Range range;

                string Material = string.Empty;
                string MaterialDesc = string.Empty;
                string SparePart = string.Empty;
                for (int i = 2; i <= excelRange.Rows.Count; i++)
                {
                    try
                    {
                        #region reading Values from Excel file and declaration of local variables 
                        range = (excelWorksheet.Cells[i, 1] as Range);
                        Material = range.Value.ToString();
                        range = (excelWorksheet.Cells[i, 2] as Range);
                        MaterialDesc = range.Value.ToString();
                        range = (excelWorksheet.Cells[i, 3] as Range);
                        SparePart = range.Value.ToString();
                        #endregion

                        string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='product'>
                        <attribute name='name' />
                        <attribute name='hil_sparepartfamily' />
                        <filter type='and'>
                            <condition attribute='name' operator='eq' value='" + Material + @"' />
                        </filter>
                        </entity>
                        </fetch>";

                        EntityCollection productcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (productcoll.Entities.Count > 0)
                        {
                            string _fetchXMLSparePart = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                            <entity name='product'>
                            <attribute name='name' />
                            <filter type='and'>
                                <condition attribute='productstructure' operator='eq' value='2' />
                                <condition attribute='hil_division' operator='eq' value='{72981D83-16FA-E811-A94C-000D3AF0694E}' />
                                <filter type='or'>
                                <condition attribute='name' operator='eq' value='" + SparePart + @"' />
                                </filter>
                            </filter>
                            </entity>
                            </fetch>";

                            EntityCollection spareColl = _service.RetrieveMultiple(new FetchExpression(_fetchXMLSparePart));
                            if (spareColl.Entities.Count > 0)
                            {
                                Entity ent = productcoll.Entities[0];
                                ent["description"] = MaterialDesc;
                                ent["hil_sparepartfamily"] = spareColl.Entities[0].ToEntityReference();
                                _service.Update(ent);
                            }
                        }
                        Console.WriteLine("Processed.." + Material + " :: " + i.ToString());
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    Console.WriteLine("Record Updated.. " + i.ToString() + "|" + Material);
                }
            }
        }
        static void SparePartFamilyUpdate()
        {
            string filePath = @"C:\Kuldeep khare\Purifier.xlsx";
            string conn = string.Empty;
            Application excelApp = new Application();
            if (excelApp != null)
            {
                Workbook excelWorkbook = excelApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet excelWorksheet = (Worksheet)excelWorkbook.Sheets[2];
                Range excelRange = excelWorksheet.UsedRange;
                Range range;

                string Material = string.Empty;
                for (int i = 1; i <= excelRange.Rows.Count; i++)
                {
                    try
                    {
                        #region reading Values from Excel file and declaration of local variables 
                        range = (excelWorksheet.Cells[i, 1] as Range);
                        Material = range.Value.ToString();
                        #endregion

                        string _fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                        <entity name='product'>
                        <attribute name='name' />
                        <filter type='and'>
                            <condition attribute='productstructure' operator='eq' value='2' />
                            <condition attribute='hil_division' operator='eq' value='{72981D83-16FA-E811-A94C-000D3AF0694E}' />
                            <filter type='or'>
                            <condition attribute='name' operator='eq' value='"+ Material + @"' />
                            <condition attribute='name' operator='eq' value='" + Material + @"-WP' />
                            </filter>
                        </filter>
                        </entity>
                        </fetch>";

                        EntityCollection productcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
                        if (productcoll.Entities.Count == 0) {
                            #region Calculate Next Serial Number
                            string _fetchXMLNum = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                              <entity name='product'>
                                <attribute name='productid' />
                                <filter type='and'>
                                  <condition attribute='productstructure' operator='eq' value='2' />
                                  <condition attribute='hil_division' operator='eq' value='{72981D83-16FA-E811-A94C-000D3AF0694E}' />
                                </filter>
                              </entity>
                            </fetch>";
                            EntityCollection Numcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXMLNum));
                            int _nextNum = Numcoll.Entities.Count + 1;
                            #endregion

                            Entity ent = new Entity("product");
                            ent["name"] = Material+"-WP";
                            ent["productnumber"] = "PF-WP" + _nextNum.ToString().PadLeft(2, '0');
                            ent["producttypecode"] = new OptionSetValue(1);
                            ent["description"] = "WP-Spare Part Family";
                            ent["productstructure"] = new OptionSetValue(2);
                            ent["msdyn_fieldserviceproducttype"] = new OptionSetValue(690970001);
                            ent["msdyn_purchasename"] = Material + "-WP";
                            ent["hil_division"] = new EntityReference("product",new Guid("72981D83-16FA-E811-A94C-000D3AF0694E"));
                            _service.Create(ent);
                        }
                        else
                        {
                            string _name = productcoll.Entities[0].GetAttributeValue<string>("name");
                            if (!_name.Contains("-WP")) {
                                Entity ent = productcoll.Entities[0];
                                ent["name"] = _name + "-WP";
                                _service.Update(ent);
                            }
                        }
                        Console.WriteLine("Processed.." + Material + " :: "+ i.ToString());
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    Console.WriteLine("Record Updated.. " + i.ToString() + "|" + Material);
                }
            }
        }
        static void MG_HSN_ProductUpdateTempTable()
        {
            QueryExpression query = null;
            EntityCollection entcoll = null;
            string Material = string.Empty;
            string HSN_Code = string.Empty;
            string HSN_Name = string.Empty;
            string HSN_Tax_Rate = string.Empty;
            string MG_Code = string.Empty;
            string MG_Name = string.Empty;
            string Div_Code = string.Empty;
            string Div_Name = string.Empty;
            Guid _hsnId = Guid.Empty;
            Guid _mgId = Guid.Empty;
            Guid _divId = Guid.Empty;
            Entity entTemp = null;
            string fetchquery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='hil_producttaxdata'>
                <attribute name='hil_producttaxdataid' />
                <attribute name='hil_name' />
                <attribute name='createdon' />
                <attribute name='hil_divisioncode' />
                <attribute name='hil_divisionname' />
                <attribute name='hil_hsncode' />
                <attribute name='hil_hsnname' />
                <attribute name='hil_hsntaxrate' />
                <attribute name='hil_mgcode' />
                <attribute name='hil_mgname' />
                <order attribute='hil_name' descending='false' />
                <filter type='and'>
                  <condition attribute='statecode' operator='eq' value='0' />
                </filter>
              </entity>
            </fetch>";

            while (true)
            {
                EntityCollection jobs = _service.RetrieveMultiple(new FetchExpression(fetchquery));
                if (jobs.Entities.Count == 0) { break; }
                int i = 1;
                foreach (Entity  ent in jobs.Entities)
                {
                    try
                    {
                        #region reading Values from TempData and declaration of local variables 
                        if (ent.Contains("hil_name"))
                            Material = ent.GetAttributeValue<string>("hil_name");
                        if (ent.Contains("hil_hsncode"))
                            HSN_Code = ent.GetAttributeValue<string>("hil_hsncode");
                        if (ent.Contains("hil_hsnname"))
                            HSN_Name = ent.GetAttributeValue<string>("hil_hsnname");
                        if (ent.Contains("hil_hsntaxrate"))
                            HSN_Tax_Rate = ent.GetAttributeValue<string>("hil_hsntaxrate");
                        if (ent.Contains("hil_mgcode"))
                            MG_Code = ent.GetAttributeValue<string>("hil_mgcode");
                        if (ent.Contains("hil_mgname"))
                            MG_Name = ent.GetAttributeValue<string>("hil_mgname");
                        if (ent.Contains("hil_divisioncode"))
                            Div_Code = ent.GetAttributeValue<string>("hil_divisioncode");
                        if (ent.Contains("hil_divisionname"))
                            Div_Name = ent.GetAttributeValue<string>("hil_mgname");

                        _hsnId = Guid.Empty;
                        _mgId = Guid.Empty;
                        _divId = Guid.Empty;
                        #endregion

                        #region Product Division Validation
                        query = new QueryExpression("product");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_sapcode", ConditionOperator.Equal, Div_Code);
                        entcoll = _service.RetrieveMultiple(query);
                        if (entcoll.Entities.Count > 0)
                        {
                            _divId = entcoll.Entities[0].Id;
                        }
                        else
                        {
                            Console.WriteLine("Product Division Doesn't exist: " + Material + "|" + Div_Code);
                            continue;
                        }
                        #endregion

                        #region HSN Code Validation
                        query = new QueryExpression("hil_hsncode");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, HSN_Code);

                        entcoll = _service.RetrieveMultiple(query);
                        if (entcoll.Entities.Count == 0)
                        {
                            Console.Write("HSN Code doesnt exitst in system: " + HSN_Code);
                            Entity newhsn = new Entity("hil_hsncode");
                            newhsn["hil_name"] = HSN_Code;
                            newhsn["hil_taxtypetext"] = HSN_Name;
                            newhsn["hil_taxpercentage"] = Convert.ToDecimal(HSN_Tax_Rate);
                            _hsnId = _service.Create(newhsn);
                        }
                        else
                        {
                            _hsnId = entcoll.Entities[0].Id;
                        }
                        #endregion

                        #region Meterial Group Validation
                        query = new QueryExpression("product");
                        query.ColumnSet = new ColumnSet("hil_productcode");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_productcode", ConditionOperator.Equal, MG_Code);
                        query.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 3); // Material Group

                        entcoll = _service.RetrieveMultiple(query);

                        if (entcoll.Entities.Count == 0)
                        {
                            Console.Write("Material Group doesn't exitst in system: " + MG_Code);
                            Entity newhsn = new Entity("product");
                            newhsn["hil_productcode"] = MG_Code;
                            newhsn["productnumber"] = MG_Code;
                            newhsn["name"] = MG_Name;
                            newhsn["hil_hierarchylevel"] = new OptionSetValue(3);// Material Group
                            newhsn["producttypecode"] = new OptionSetValue(1); // Finished Good
                            newhsn["description"] = MG_Name;
                            newhsn["hil_division"] = new EntityReference("product", _divId);
                            newhsn["defaultuomscheduleid"] = new EntityReference("uomschedule", new Guid("AF39A94C-F79F-4E6D-9A9E-20F2948FE185"));
                            newhsn["defaultuomid"] = new EntityReference("uom", new Guid("0359D51B-D7CF-43B1-87F6-FC13A2C1DEC8"));
                            _mgId = _service.Create(newhsn);
                        }
                        else
                        {
                            _mgId = entcoll.Entities[0].Id;
                        }
                        #endregion


                        query = new QueryExpression("product");
                        query.ColumnSet = new ColumnSet("hil_productcode");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_productcode", ConditionOperator.Equal, Material);

                        EntityCollection productcoll = _service.RetrieveMultiple(query);
                        if (productcoll.Entities.Count > 0)
                        {
                            Entity updateproduct = productcoll.Entities[0];
                            updateproduct["hil_hsncode"] = new EntityReference("hil_hsncode", _hsnId);
                            updateproduct["hil_materialgroup"] = new EntityReference("product", _mgId);
                            _service.Update(updateproduct);

                            entTemp = new Entity("hil_producttaxdata", ent.Id);
                            entTemp["statecode"] = new OptionSetValue(1);
                            _service.Update(entTemp);
                        }
                        else
                        {
                            Console.WriteLine("Product Doesn't exist: " + Material);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }
                    Console.WriteLine("Record Updated.. " + i++.ToString() + "|" + Material);
                }
            }
        }
        static void MG_HSN_ProductUpdateExcel()
        {
            string filePath = @"C:\Users\33632\OneDrive - Havells\Desktop\D365 Projects\HSN and Tax rate\motor.xlsx";
            string conn = string.Empty;
            Application excelApp = new Application();
            if (excelApp != null)
            {
                Workbook excelWorkbook = excelApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet excelWorksheet = (Worksheet)excelWorkbook.Sheets[1];
                Range excelRange = excelWorksheet.UsedRange;
                Range range;

                string Material = string.Empty;
                string HSN_Code = string.Empty;
                string HSN_Name = string.Empty;
                string HSN_Tax_Rate = string.Empty;
                string MG_Code = string.Empty;
                string MG_Name = string.Empty;
                string Div_Code = string.Empty;
                string Div_Name = string.Empty;
                Guid _hsnId = Guid.Empty;
                Guid _mgId = Guid.Empty;
                Guid _divId = Guid.Empty;
                QueryExpression query = null;
                EntityCollection entcoll = null;
                for (int i = 2; i <= excelRange.Rows.Count; i++)
                {
                    try
                    {
                        #region reading Values from Excel file and declaration of local variables 
                        range = (excelWorksheet.Cells[i, 1] as Range);
                        Material = range.Value.ToString();

                        range = (excelWorksheet.Cells[i, 2] as Range);
                        HSN_Code = range.Value.ToString();

                        range = (excelWorksheet.Cells[i, 3] as Range);
                        HSN_Name = range.Value.ToString();

                        range = (excelWorksheet.Cells[i, 4] as Range);
                        HSN_Tax_Rate = range.Value.ToString();

                        range = (excelWorksheet.Cells[i, 5] as Range);
                        MG_Code = range.Value.ToString();

                        range = (excelWorksheet.Cells[i, 6] as Range);
                        MG_Name = range.Value.ToString();

                        range = (excelWorksheet.Cells[i, 7] as Range);
                        Div_Code = range.Value.ToString();

                        range = (excelWorksheet.Cells[i, 8] as Range);
                        Div_Name = range.Value.ToString();

                        _hsnId = Guid.Empty;
                        _mgId = Guid.Empty;
                        _divId = Guid.Empty;
                        #endregion

                        #region Product Division Validation
                        query = new QueryExpression("product");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_sapcode", ConditionOperator.Equal, Div_Code);
                        entcoll = _service.RetrieveMultiple(query);
                        if (entcoll.Entities.Count >0)
                        {
                            _divId = entcoll.Entities[0].Id;
                        }
                        else
                        {
                            Console.WriteLine("Product Division Doesn't exist: " + Material + "|" + Div_Code);
                            continue;
                        }
                        #endregion

                        #region HSN Code Validation
                        query = new QueryExpression("hil_hsncode");
                        query.ColumnSet = new ColumnSet(false);
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_name", ConditionOperator.Equal, HSN_Code);

                        entcoll = _service.RetrieveMultiple(query);
                        if (entcoll.Entities.Count == 0)
                        {
                            Console.Write("HSN Code doesnt exitst in system: " + HSN_Code);
                            Entity newhsn = new Entity("hil_hsncode");
                            newhsn["hil_name"] = HSN_Code;
                            newhsn["hil_taxtypetext"] = HSN_Name;
                            newhsn["hil_taxpercentage"] = Convert.ToDecimal(HSN_Tax_Rate);
                            _hsnId = _service.Create(newhsn);
                        }
                        else
                        {
                            _hsnId = entcoll.Entities[0].Id;
                        }
                        #endregion

                        #region Meterial Group Validation
                        query = new QueryExpression("product");
                        query.ColumnSet = new ColumnSet("hil_productcode");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_productcode", ConditionOperator.Equal, MG_Code);
                        query.Criteria.AddCondition("hil_hierarchylevel", ConditionOperator.Equal, 3); // Material Group

                        entcoll = _service.RetrieveMultiple(query);

                        if (entcoll.Entities.Count == 0)
                        {
                            Console.Write("Material Group doesn't exitst in system: " + MG_Code);
                            Entity newhsn = new Entity("product");
                            newhsn["hil_productcode"] = MG_Code;
                            newhsn["productnumber"] = MG_Code;
                            newhsn["name"] = MG_Name;
                            newhsn["hil_hierarchylevel"] = new OptionSetValue(3);// Material Group
                            newhsn["producttypecode"] = new OptionSetValue(1); // Finished Good
                            newhsn["description"] = MG_Name;
                            newhsn["hil_division"] = new EntityReference("product", _divId);
                            _mgId = _service.Create(newhsn);
                        }
                        else
                        {
                            _mgId = entcoll.Entities[0].Id;
                        }
                        #endregion


                        query = new QueryExpression("product");
                        query.ColumnSet = new ColumnSet("hil_productcode");
                        query.Criteria = new FilterExpression(LogicalOperator.And);
                        query.Criteria.AddCondition("hil_productcode", ConditionOperator.Equal, Material);

                        EntityCollection productcoll = _service.RetrieveMultiple(query);
                        if (productcoll.Entities.Count > 0)
                        {
                            Entity updateproduct = productcoll.Entities[0];
                            updateproduct["hil_hsncode"] = new EntityReference("hil_hsncode", _hsnId);
                            updateproduct["hil_materialgroup"] = new EntityReference("product", _mgId);
                            //updateproduct["hil_division"] = new EntityReference("product", _divId);
                            _service.Update(updateproduct);
                            
                        }
                        else
                        {
                            Console.WriteLine("Product Doesn't exist: " + Material);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }
                    Console.WriteLine("Record Updated.. " + i.ToString() + "|" + Material);
                }
            }
        }
        //static void ChannelPartnerDataUpdatesFromSamparkExcel()
        //{
        //    string filePath = @"C:\Kuldeep khare\RetailerMaster1.xlsx";
        //    string conn = string.Empty;

        //    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
        //    if (excelApp != null)
        //    {
        //        Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(filePath, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
        //        Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];

        //        Excel.Range excelRange = excelWorksheet.UsedRange;
        //        Excel.Range range;
        //        string _mobileNumber, _userType, _userSubType, _customerCode, _fullName, _businessName, _businessAddress, _businessAddress2, _pincode, _emailid, _mdmCode;
        //        QueryExpression Query1;
        //        EntityCollection entcoll, entcoll1, entcoll2;
        //        Entity jobEntity;
        //        string _fetchXML;

        //        for (int i = 1001; i <= excelRange.Rows.Count; i++)
        //        {
        //            try
        //            {
        //                Guid _accountId = Guid.Empty;

        //                Console.WriteLine("Processing Row# " + i.ToString());
        //                range = (excelWorksheet.Cells[i, 1] as Excel.Range);
        //                _mobileNumber = range.Value.ToString();

        //                range = (excelWorksheet.Cells[i, 2] as Excel.Range);
        //                _userType = range.Value.ToString();

        //                range = (excelWorksheet.Cells[i, 3] as Excel.Range);
        //                _userSubType = range.Value != null ? range.Value.ToString() : "";

        //                range = (excelWorksheet.Cells[i, 4] as Excel.Range);
        //                _customerCode = range.Value != null ? range.Value.ToString() : "";

        //                range = (excelWorksheet.Cells[i, 5] as Excel.Range);
        //                _fullName = range.Value != null ? range.Value.ToString() : "";

        //                range = (excelWorksheet.Cells[i, 6] as Excel.Range);
        //                _businessName = range.Value != null ? range.Value.ToString() : "";

        //                range = (excelWorksheet.Cells[i, 7] as Excel.Range);
        //                _businessAddress = range.Value != null ? range.Value.ToString() : "";

        //                range = (excelWorksheet.Cells[i, 8] as Excel.Range);
        //                _businessAddress2 = range.Value != null ? range.Value.ToString() : "";

        //                range = (excelWorksheet.Cells[i, 9] as Excel.Range);
        //                _pincode = range.Value != null ? range.Value.ToString() : "";

        //                range = (excelWorksheet.Cells[i, 10] as Excel.Range);
        //                _emailid = range.Value != null ? range.Value.ToString() : "";

        //                range = (excelWorksheet.Cells[i, 11] as Excel.Range);
        //                _mdmCode = range.Value != null ? range.Value.ToString() : "";

        //                Console.WriteLine("Row Data " + _mobileNumber + "/" + _userType + "/" + _userSubType + "/" + _customerCode + "/" + _fullName + "/" + _businessName + "/" + _businessAddress + "/" + _businessAddress2 + "/" + _pincode + "/" + _emailid + "/" + _mdmCode);

        //                Query1 = new QueryExpression("account");
        //                Query1.ColumnSet = new ColumnSet("hil_vendorcode");
        //                Query1.Criteria = new FilterExpression(LogicalOperator.And);
        //                Query1.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, _customerCode);
        //                Query1.AddOrder("createdon", OrderType.Descending);
        //                Query1.TopCount = 1;
        //                entcoll = _service.RetrieveMultiple(Query1);
        //                if (entcoll.Entities.Count > 0)
        //                {
        //                    if (entcoll.Entities[0].Contains("hil_vendorcode"))
        //                    {
        //                        Console.WriteLine(i.ToString() + "/" + excelRange.Rows.Count.ToString() + " Channel Partner is already exist.");
        //                        continue;
        //                    }
        //                    _accountId = entcoll.Entities[0].Id;
        //                }
        //                else
        //                {
        //                    Query1 = new QueryExpression("account");
        //                    Query1.ColumnSet = new ColumnSet("hil_vendorcode");
        //                    Query1.Criteria = new FilterExpression(LogicalOperator.And);
        //                    Query1.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, _mdmCode);
        //                    Query1.AddOrder("createdon", OrderType.Descending);
        //                    Query1.TopCount = 1;
        //                    entcoll = _service.RetrieveMultiple(Query1);
        //                    if (entcoll.Entities.Count > 0)
        //                    {
        //                        if (entcoll.Entities[0].Contains("hil_vendorcode"))
        //                        {
        //                            Console.WriteLine(i.ToString() + "/" + excelRange.Rows.Count.ToString() + " Channel Partner is already exist.");
        //                            continue;
        //                        }
        //                        _accountId = entcoll.Entities[0].Id;
        //                    }
        //                }

        //                if (_accountId != Guid.Empty)
        //                {
        //                    jobEntity = new Entity("account", _accountId);
        //                    jobEntity["name"] = string.IsNullOrEmpty(_businessName) || string.IsNullOrWhiteSpace(_businessName) ? _fullName : _businessName;
        //                    jobEntity["emailaddress1"] = _emailid;
        //                    jobEntity["hil_vendorcode"] = _mdmCode;
        //                    if (!string.IsNullOrEmpty(_userSubType) && !string.IsNullOrWhiteSpace(_userSubType))
        //                    {
        //                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                          <entity name='hil_usertype'>
        //                            <attribute name='hil_usertypeid' />
        //                            <attribute name='hil_parentusertype' />
        //                            <filter type='and'>
        //                              <condition attribute='hil_parentusertypename' operator='like' value='%{_userType}%' />
        //                              <condition attribute='hil_name' operator='like' value='%{_userSubType}%' />
        //                            </filter>
        //                          </entity>
        //                        </fetch>";
        //                        entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                        if (entcoll2.Entities.Count > 0)
        //                        {
        //                            jobEntity["hil_usertype"] = entcoll2.Entities[0].GetAttributeValue<EntityReference>("hil_parentusertype");
        //                            jobEntity["hil_usersubtype"] = entcoll2.Entities[0].ToEntityReference();
        //                        }
        //                        else
        //                        {
        //                            _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                              <entity name='hil_usertype'>
        //                                <attribute name='hil_usertypeid' />
        //                                <filter type='and'>
        //                                  <condition attribute='hil_name' operator='like' value='%{_userType}%' />
        //                                </filter>
        //                              </entity>
        //                            </fetch>";
        //                            entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                            if (entcoll2.Entities.Count > 0)
        //                            {
        //                                jobEntity["hil_usertype"] = entcoll2.Entities[0].ToEntityReference();
        //                            }
        //                            //
        //                            _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                              <entity name='hil_usertype'>
        //                                <attribute name='hil_usertypeid' />
        //                                <filter type='and'>
        //                                  <condition attribute='hil_name' operator='like' value='%{_userSubType}%' />
        //                                </filter>
        //                              </entity>
        //                            </fetch>";
        //                            entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                            if (entcoll2.Entities.Count > 0)
        //                            {
        //                                jobEntity["hil_usersubtype"] = entcoll2.Entities[0].ToEntityReference();
        //                            }
        //                        }
        //                    }
        //                    else if (!string.IsNullOrEmpty(_userType) && !string.IsNullOrWhiteSpace(_userType))
        //                    {
        //                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                              <entity name='hil_usertype'>
        //                                <attribute name='hil_usertypeid' />
        //                                <filter type='and'>
        //                                  <condition attribute='hil_name' operator='like' value='%{_userType}%' />
        //                                </filter>
        //                              </entity>
        //                            </fetch>";
        //                        entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                        if (entcoll2.Entities.Count > 0)
        //                        {
        //                            jobEntity["hil_usertype"] = entcoll2.Entities[0].ToEntityReference();
        //                        }
        //                    }
        //                    _service.Update(jobEntity);
        //                    Console.WriteLine(i.ToString() + "/" + excelRange.Rows.Count.ToString() + " Updating Channel Partner " + _customerCode + "/" + _mdmCode);
        //                    continue;
        //                }
        //                _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                <entity name='hil_businessmapping'>
        //                <attribute name='hil_businessmappingid' />
        //                <attribute name='hil_pincode' />
        //                <attribute name='hil_salesoffice' />
        //                <attribute name='hil_subterritory' />
        //                <attribute name='hil_region' />
        //                <attribute name='hil_state' />
        //                <attribute name='hil_branch' />
        //                <attribute name='hil_city' />
        //                <attribute name='hil_area' />
        //                <attribute name='hil_district' />
        //                <attribute name='hil_name' />
        //                <attribute name='createdon' />
        //                <order attribute='hil_name' descending='false' />
        //                <filter type='and'>
        //                    <condition attribute='statecode' operator='eq' value='0' />
        //                    <condition attribute='hil_pincodename' operator='like' value='%{_pincode}%' />
        //                </filter>
        //                </entity>
        //                </fetch>";

        //                entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                if (entcoll1.Entities.Count > 0)
        //                {
        //                    jobEntity = new Entity("account");

        //                    if (!string.IsNullOrEmpty(_userSubType) && !string.IsNullOrWhiteSpace(_userSubType))
        //                    {
        //                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                          <entity name='hil_usertype'>
        //                            <attribute name='hil_usertypeid' />
        //                            <attribute name='hil_parentusertype' />
        //                            <filter type='and'>
        //                              <condition attribute='hil_parentusertypename' operator='like' value='%{_userType}%' />
        //                              <condition attribute='hil_name' operator='like' value='%{_userSubType}%' />
        //                            </filter>
        //                          </entity>
        //                        </fetch>";
        //                        entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                        if (entcoll2.Entities.Count > 0)
        //                        {
        //                            jobEntity["hil_usertype"] = entcoll2.Entities[0].GetAttributeValue<EntityReference>("hil_parentusertype");
        //                            jobEntity["hil_usersubtype"] = entcoll2.Entities[0].ToEntityReference();
        //                        }
        //                        else
        //                        {
        //                            _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                              <entity name='hil_usertype'>
        //                                <attribute name='hil_usertypeid' />
        //                                <filter type='and'>
        //                                  <condition attribute='hil_name' operator='like' value='%{_userType}%' />
        //                                </filter>
        //                              </entity>
        //                            </fetch>";
        //                            entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                            if (entcoll2.Entities.Count > 0)
        //                            {
        //                                jobEntity["hil_usertype"] = entcoll2.Entities[0].ToEntityReference();
        //                            }
        //                            //
        //                            _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                              <entity name='hil_usertype'>
        //                                <attribute name='hil_usertypeid' />
        //                                <filter type='and'>
        //                                  <condition attribute='hil_name' operator='like' value='%{_userSubType}%' />
        //                                </filter>
        //                              </entity>
        //                            </fetch>";
        //                            entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                            if (entcoll2.Entities.Count > 0)
        //                            {
        //                                jobEntity["hil_usersubtype"] = entcoll2.Entities[0].ToEntityReference();
        //                            }
        //                        }
        //                    }
        //                    else if (!string.IsNullOrEmpty(_userType) && !string.IsNullOrWhiteSpace(_userType))
        //                    {
        //                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                              <entity name='hil_usertype'>
        //                                <attribute name='hil_usertypeid' />
        //                                <filter type='and'>
        //                                  <condition attribute='hil_name' operator='like' value='%{_userType}%' />
        //                                </filter>
        //                              </entity>
        //                            </fetch>";
        //                        entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                        if (entcoll2.Entities.Count > 0)
        //                        {
        //                            jobEntity["hil_usertype"] = entcoll2.Entities[0].ToEntityReference();
        //                        }
        //                    }
        //                    jobEntity["accountnumber"] = string.IsNullOrEmpty(_customerCode) || string.IsNullOrWhiteSpace(_customerCode) ? _mdmCode : _customerCode; ;
        //                    jobEntity["hil_vendorcode"] = _mdmCode;
        //                    jobEntity["name"] = string.IsNullOrEmpty(_businessName) || string.IsNullOrWhiteSpace(_businessName) ? _fullName : _businessName;
        //                    jobEntity["telephone1"] = _mobileNumber;
        //                    jobEntity["address1_line1"] = _businessAddress;
        //                    jobEntity["address1_line2"] = _businessAddress2;

        //                    jobEntity["hil_pincode"] = entcoll1.Entities[0]["hil_pincode"];
        //                    jobEntity["hil_state"] = entcoll1.Entities[0]["hil_state"];
        //                    jobEntity["hil_district"] = entcoll1.Entities[0]["hil_district"];
        //                    jobEntity["hil_city"] = entcoll1.Entities[0]["hil_city"];
        //                    jobEntity["hil_branch"] = entcoll1.Entities[0]["hil_branch"];
        //                    jobEntity["hil_salesoffice"] = entcoll1.Entities[0]["hil_salesoffice"];
        //                    jobEntity["hil_region"] = entcoll1.Entities[0]["hil_region"];
        //                    jobEntity["hil_area"] = entcoll1.Entities[0]["hil_area"];
        //                    jobEntity["hil_subterritory"] = entcoll1.Entities[0]["hil_subterritory"];

        //                    jobEntity["emailaddress1"] = _emailid;
        //                    _service.Create(jobEntity);
        //                }
        //                else
        //                {
        //                    Console.WriteLine(i.ToString() + "/" + excelRange.Rows.Count.ToString() + " Pincode does not exist. " + _pincode);
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("ERROR!!! Row# " + i.ToString() + "/" + excelRange.Rows.Count.ToString() + ex.Message);
        //            }
        //        }
        //    }

        //}
        //static void ChannelPartnerDataUpdatesFromSampark()
        //{
        //    var _startIndex = ConfigurationManager.AppSettings["StartIndex"].ToString();
        //    var _endIndex = ConfigurationManager.AppSettings["EndIndex"].ToString();
        //    string _mobileNumber, _userType, _userSubType, _customerCode, _fullName, _businessName, _businessAddress, _businessAddress2, _pincode, _emailid, _mdmCode;
        //    QueryExpression Query1;
        //    EntityCollection entcoll, entcoll1, entcoll2;
        //    Entity jobEntity;
        //    string _fetchXML;

        //    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //            <entity name='hil_samparkchannelpartner'>
        //            <attribute name='hil_name' />
        //            <attribute name='hil_usertype' />
        //            <attribute name='hil_usersubtype' />
        //            <attribute name='hil_pincode' />
        //            <attribute name='modifiedby' />
        //            <attribute name='hil_mdmcode' />
        //            <attribute name='hil_index' />
        //            <attribute name='hil_fullname' />
        //            <attribute name='hil_emailid' />
        //            <attribute name='hil_customercode' />
        //            <attribute name='hil_businessname' />
        //            <attribute name='hil_businessaddress2' />
        //            <attribute name='hil_businessaddress' />
        //            <attribute name='hil_samparkchannelpartnerid' />
        //            <order attribute='hil_index' descending='false' />
        //            <filter type='and'>
        //                <condition attribute='hil_rowstatus' operator='ne' value='1' />
        //                <condition attribute='hil_index' operator='ge' value='{_startIndex}' />
        //                <condition attribute='hil_index' operator='le' value='{_endIndex}' />
        //            </filter>
        //            </entity>
        //        </fetch>";
        //    int i = 1;
        //    Entity _entUpdate;
        //    entcoll = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //    foreach (Entity ent in entcoll.Entities)
        //    {
        //        try
        //        {
        //            Guid _accountId = Guid.Empty;

        //            i = ent.GetAttributeValue<int>("hil_index");

        //            Console.WriteLine("Processing Index# " + i.ToString());

        //            _mobileNumber = ParseValue(ent, "hil_name");

        //            _userType = ParseValue(ent, "hil_usertype");

        //            _userSubType = ParseValue(ent, "hil_usersubtype");

        //            _customerCode = ParseValue(ent, "hil_customercode");

        //            _fullName = ParseValue(ent, "hil_fullname");

        //            _businessName = ParseValue(ent, "hil_businessname");

        //            _businessAddress = ParseValue(ent, "hil_businessaddress");

        //            _businessAddress2 = ParseValue(ent, "hil_businessaddress2");

        //            _pincode = ParseValue(ent, "hil_pincode");

        //            _emailid = ParseValue(ent, "hil_emailid");

        //            _mdmCode = ParseValue(ent, "hil_mdmcode");

        //            Console.WriteLine("Row Data " + _mobileNumber + "/" + _userType + "/" + _userSubType + "/" + _customerCode + "/" + _fullName + "/" + _businessName + "/" + _businessAddress + "/" + _businessAddress2 + "/" + _pincode + "/" + _emailid + "/" + _mdmCode);

        //            Query1 = new QueryExpression("account");
        //            Query1.ColumnSet = new ColumnSet("hil_vendorcode");
        //            Query1.Criteria = new FilterExpression(LogicalOperator.And);
        //            Query1.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, _customerCode);
        //            Query1.AddOrder("createdon", OrderType.Descending);
        //            Query1.TopCount = 1;
        //            EntityCollection entcol2 = _service.RetrieveMultiple(Query1);
        //            if (entcol2.Entities.Count > 0)
        //            {
        //                if (entcol2.Entities[0].Contains("hil_vendorcode"))
        //                {
        //                    Console.WriteLine(i.ToString() + "/" + entcoll.Entities.Count.ToString() + " Channel Partner is already exist.");
        //                    _entUpdate = new Entity("hil_samparkchannelpartner", ent.Id);
        //                    _entUpdate["hil_rowstatus"] = true;
        //                    _service.Update(_entUpdate);
        //                    continue;
        //                }
        //                _accountId = entcol2.Entities[0].Id;
        //            }
        //            else
        //            {
        //                Query1 = new QueryExpression("account");
        //                Query1.ColumnSet = new ColumnSet("hil_vendorcode");
        //                Query1.Criteria = new FilterExpression(LogicalOperator.And);
        //                Query1.Criteria.AddCondition("accountnumber", ConditionOperator.Equal, _mdmCode);
        //                Query1.AddOrder("createdon", OrderType.Descending);
        //                Query1.TopCount = 1;
        //                entcol2 = _service.RetrieveMultiple(Query1);
        //                if (entcol2.Entities.Count > 0)
        //                {
        //                    if (entcol2.Entities[0].Contains("hil_vendorcode"))
        //                    {
        //                        Console.WriteLine(i.ToString() + "/" + entcoll.Entities.Count.ToString() + " Channel Partner is already exist.");
        //                        _entUpdate = new Entity("hil_samparkchannelpartner", ent.Id);
        //                        _entUpdate["hil_rowstatus"] = true;
        //                        _service.Update(_entUpdate);
        //                        continue;
        //                    }
        //                    _accountId = entcol2.Entities[0].Id;
        //                }
        //            }

        //            if (_accountId != Guid.Empty)
        //            {
        //                jobEntity = new Entity("account", _accountId);
        //                jobEntity["name"] = string.IsNullOrEmpty(_businessName) || string.IsNullOrWhiteSpace(_businessName) ? _fullName : _businessName;
        //                jobEntity["emailaddress1"] = _emailid;
        //                jobEntity["hil_vendorcode"] = _mdmCode;
        //                if (!string.IsNullOrEmpty(_userSubType) && !string.IsNullOrWhiteSpace(_userSubType))
        //                {
        //                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                          <entity name='hil_usertype'>
        //                            <attribute name='hil_usertypeid' />
        //                            <attribute name='hil_parentusertype' />
        //                            <filter type='and'>
        //                              <condition attribute='hil_parentusertypename' operator='like' value='%{_userType}%' />
        //                              <condition attribute='hil_name' operator='like' value='%{_userSubType}%' />
        //                            </filter>
        //                          </entity>
        //                        </fetch>";
        //                    entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                    if (entcoll2.Entities.Count > 0)
        //                    {
        //                        jobEntity["hil_usertype"] = entcoll2.Entities[0].GetAttributeValue<EntityReference>("hil_parentusertype");
        //                        jobEntity["hil_usersubtype"] = entcoll2.Entities[0].ToEntityReference();
        //                    }
        //                    else
        //                    {
        //                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                              <entity name='hil_usertype'>
        //                                <attribute name='hil_usertypeid' />
        //                                <filter type='and'>
        //                                  <condition attribute='hil_name' operator='like' value='%{_userType}%' />
        //                                </filter>
        //                              </entity>
        //                            </fetch>";
        //                        entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                        if (entcoll2.Entities.Count > 0)
        //                        {
        //                            jobEntity["hil_usertype"] = entcoll2.Entities[0].ToEntityReference();
        //                        }
        //                        //
        //                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                              <entity name='hil_usertype'>
        //                                <attribute name='hil_usertypeid' />
        //                                <filter type='and'>
        //                                  <condition attribute='hil_name' operator='like' value='%{_userSubType}%' />
        //                                </filter>
        //                              </entity>
        //                            </fetch>";
        //                        entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                        if (entcoll2.Entities.Count > 0)
        //                        {
        //                            jobEntity["hil_usersubtype"] = entcoll2.Entities[0].ToEntityReference();
        //                        }
        //                    }
        //                }
        //                else if (!string.IsNullOrEmpty(_userType) && !string.IsNullOrWhiteSpace(_userType))
        //                {
        //                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                              <entity name='hil_usertype'>
        //                                <attribute name='hil_usertypeid' />
        //                                <filter type='and'>
        //                                  <condition attribute='hil_name' operator='like' value='%{_userType}%' />
        //                                </filter>
        //                              </entity>
        //                            </fetch>";
        //                    entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                    if (entcoll2.Entities.Count > 0)
        //                    {
        //                        jobEntity["hil_usertype"] = entcoll2.Entities[0].ToEntityReference();
        //                    }
        //                }
        //                _service.Update(jobEntity);

        //                _entUpdate = new Entity("hil_samparkchannelpartner", ent.Id);
        //                _entUpdate["hil_rowstatus"] = true;
        //                _service.Update(_entUpdate);

        //                Console.WriteLine(i.ToString() + "/" + entcoll.Entities.Count.ToString() + " Updating Channel Partner " + _customerCode + "/" + _mdmCode);
        //                continue;
        //            }
        //            _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                <entity name='hil_businessmapping'>
        //                <attribute name='hil_businessmappingid' />
        //                <attribute name='hil_pincode' />
        //                <attribute name='hil_salesoffice' />
        //                <attribute name='hil_subterritory' />
        //                <attribute name='hil_region' />
        //                <attribute name='hil_state' />
        //                <attribute name='hil_branch' />
        //                <attribute name='hil_city' />
        //                <attribute name='hil_area' />
        //                <attribute name='hil_district' />
        //                <attribute name='hil_name' />
        //                <attribute name='createdon' />
        //                <order attribute='hil_name' descending='false' />
        //                <filter type='and'>
        //                    <condition attribute='statecode' operator='eq' value='0' />
        //                    <condition attribute='hil_pincodename' operator='like' value='%{_pincode}%' />
        //                </filter>
        //                </entity>
        //                </fetch>";

        //            entcoll1 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //            if (entcoll1.Entities.Count > 0)
        //            {
        //                jobEntity = new Entity("account");

        //                if (!string.IsNullOrEmpty(_userSubType) && !string.IsNullOrWhiteSpace(_userSubType))
        //                {
        //                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                          <entity name='hil_usertype'>
        //                            <attribute name='hil_usertypeid' />
        //                            <attribute name='hil_parentusertype' />
        //                            <filter type='and'>
        //                              <condition attribute='hil_parentusertypename' operator='like' value='%{_userType}%' />
        //                              <condition attribute='hil_name' operator='like' value='%{_userSubType}%' />
        //                            </filter>
        //                          </entity>
        //                        </fetch>";
        //                    entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                    if (entcoll2.Entities.Count > 0)
        //                    {
        //                        jobEntity["hil_usertype"] = entcoll2.Entities[0].GetAttributeValue<EntityReference>("hil_parentusertype");
        //                        jobEntity["hil_usersubtype"] = entcoll2.Entities[0].ToEntityReference();
        //                    }
        //                    else
        //                    {
        //                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                              <entity name='hil_usertype'>
        //                                <attribute name='hil_usertypeid' />
        //                                <filter type='and'>
        //                                  <condition attribute='hil_name' operator='like' value='%{_userType}%' />
        //                                </filter>
        //                              </entity>
        //                            </fetch>";
        //                        entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                        if (entcoll2.Entities.Count > 0)
        //                        {
        //                            jobEntity["hil_usertype"] = entcoll2.Entities[0].ToEntityReference();
        //                        }
        //                        //
        //                        _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                              <entity name='hil_usertype'>
        //                                <attribute name='hil_usertypeid' />
        //                                <filter type='and'>
        //                                  <condition attribute='hil_name' operator='like' value='%{_userSubType}%' />
        //                                </filter>
        //                              </entity>
        //                            </fetch>";
        //                        entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                        if (entcoll2.Entities.Count > 0)
        //                        {
        //                            jobEntity["hil_usersubtype"] = entcoll2.Entities[0].ToEntityReference();
        //                        }
        //                    }
        //                }
        //                else if (!string.IsNullOrEmpty(_userType) && !string.IsNullOrWhiteSpace(_userType))
        //                {
        //                    _fetchXML = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        //                              <entity name='hil_usertype'>
        //                                <attribute name='hil_usertypeid' />
        //                                <filter type='and'>
        //                                  <condition attribute='hil_name' operator='like' value='%{_userType}%' />
        //                                </filter>
        //                              </entity>
        //                            </fetch>";
        //                    entcoll2 = _service.RetrieveMultiple(new FetchExpression(_fetchXML));
        //                    if (entcoll2.Entities.Count > 0)
        //                    {
        //                        jobEntity["hil_usertype"] = entcoll2.Entities[0].ToEntityReference();
        //                    }
        //                }
        //                jobEntity["accountnumber"] = string.IsNullOrEmpty(_customerCode) || string.IsNullOrWhiteSpace(_customerCode) ? _mdmCode : _customerCode; ;
        //                jobEntity["hil_vendorcode"] = _mdmCode;
        //                jobEntity["name"] = string.IsNullOrEmpty(_businessName) || string.IsNullOrWhiteSpace(_businessName) ? _fullName : _businessName;
        //                jobEntity["telephone1"] = _mobileNumber;
        //                jobEntity["address1_line1"] = _businessAddress;
        //                jobEntity["address1_line2"] = _businessAddress2;

        //                jobEntity["hil_pincode"] = entcoll1.Entities[0]["hil_pincode"];
        //                jobEntity["hil_state"] = entcoll1.Entities[0]["hil_state"];
        //                jobEntity["hil_district"] = entcoll1.Entities[0]["hil_district"];
        //                jobEntity["hil_city"] = entcoll1.Entities[0]["hil_city"];
        //                jobEntity["hil_branch"] = entcoll1.Entities[0]["hil_branch"];
        //                jobEntity["hil_salesoffice"] = entcoll1.Entities[0]["hil_salesoffice"];
        //                jobEntity["hil_region"] = entcoll1.Entities[0]["hil_region"];
        //                jobEntity["hil_area"] = entcoll1.Entities[0]["hil_area"];
        //                jobEntity["hil_subterritory"] = entcoll1.Entities[0]["hil_subterritory"];

        //                jobEntity["emailaddress1"] = _emailid;
        //                _service.Create(jobEntity);
        //            }
        //            else
        //            {
        //                Console.WriteLine(i.ToString() + "/" + entcoll.Entities.Count.ToString() + " Pincode does not exist. " + _pincode);
        //            }
        //            Entity _ent = new Entity("hil_samparkchannelpartner", ent.Id);
        //            _ent["hil_rowstatus"] = true;
        //            _service.Update(_ent);
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine("ERROR!!! Row# " + i.ToString() + "/" + entcoll.Entities.Count.ToString() + ex.Message);
        //        }
        //    }
        //}
        static string ParseValue(Entity _obj, string _fieldName)
        {
            if (_obj.Contains(_fieldName))
                return _obj.GetAttributeValue<string>(_fieldName);
            else
                return string.Empty;
        }
        static void ClosePendingWorkDoneJobs()
        {

            try
            {
                var DateBefore = ConfigurationManager.AppSettings["DateBefore"].ToString();
                var Orderby = ConfigurationManager.AppSettings["Orderby"].ToString();
                //string finalString = string.Format(connStr, CrmURL);
                EntityCollection jobsColl = new EntityCollection();

                string fetchquery = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1000'>
                    <entity name='msdyn_workorder'>
                    <attribute name='msdyn_name'/>
                    <order attribute='hil_jobclosuredon' descending='{Orderby}'/>
                    <filter type='and'>
                    <condition attribute='msdyn_substatus' operator='eq' value='2927FA6C-FA0F-E911-A94E-000D3AF060A1'/>
                    <condition attribute='hil_jobclosuredon' operator='on-or-before' value='{DateBefore}'/>
                    <condition attribute='hil_callsubtype' operator='ne' value='55A71A52-3C0B-E911-A94E-000D3AF06CD4'/>
                    </filter>
                    </entity>
                    </fetch>";
                EntityCollection jobs;
                Entity entWorkorder;
                int i = 1;
                while (true)
                {
                    jobs = _service.RetrieveMultiple(new FetchExpression(fetchquery));
                    if (jobs.Entities.Count == 0) { break; }
                    foreach (Entity item in jobs.Entities)
                    {
                        try
                        {
                            entWorkorder = new Entity("msdyn_workorder", item.Id);
                            entWorkorder["hil_closeticket"] = true;
                            entWorkorder["hil_kkgcode_sms"] = new OptionSetValue(910590006);
                            entWorkorder["hil_webclosureremarks"] = "Closed from backend as per approval";
                            entWorkorder["hil_closureremarks"] = "Closed from backend as per approval";
                            _service.Update(entWorkorder);
                            Console.WriteLine("Processing... " + i++.ToString() + " Jobs : " + item.Attributes["msdyn_name"].ToString());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Processing... " + i++.ToString() + " Error : " + item.Attributes["msdyn_name"].ToString() + "/" + ex.Message);
                            //StreamWriter log;
                            //if (!File.Exists(@"C:\Kuldeep khare\Errors\logfile.xls"))
                            //{
                            //    log = new StreamWriter(@"C:\Kuldeep khare\Errors\logfile.xls");
                            //}
                            //else
                            //{
                            //    log = File.AppendText(@"C:\Kuldeep khare\Errors\logfile.xls");
                            //}
                            //log.Write(item.Attributes["msdyn_name"].ToString() + "||" + ex.Message);
                            //log.WriteLine();
                            //log.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:- " + ex.Message);
            }
        }
        #region App Setting Load/CRM Connection
        static IOrganizationService ConnectToCRM(string connectionString)
        {
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);
                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Creating Conn: " + ex.Message);
            }
            return service;

        }
        #endregion
    }
}
