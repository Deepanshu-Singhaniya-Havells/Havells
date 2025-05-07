using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.ProductAndPriceListSync
{
    public class SyncProducts
    {
        public static void syncProducts(IOrganizationService service, string _syncDatetime)
        {
            Console.WriteLine("*******************************  syncProducts Started ******************************* ");
            try
            {
                WebClient webClient = new WebClient();
                Integration intConf = Models.GetIntegration(service, "Product_Material");

                string _enquiryFromDatetime = _syncDatetime != "" ? _syncDatetime : Models.getTimeStamp(service);
                DateTime today = DateTime.Now.AddHours(6);
                string _enquiryToDatetime = today.Year.ToString() + today.Month.ToString().PadLeft(2, '0') + today.Day.ToString().PadLeft(2, '0') + today.Hour.ToString().PadLeft(2, '0') + today.Minute.ToString().PadLeft(2, '0') + (today.Second + 1).ToString().PadLeft(2, '0');
                Console.WriteLine("start downloading from dtate " + _enquiryFromDatetime + " Till " + _enquiryToDatetime);
                string uri = intConf.uri;
                string authInfo = intConf.userName + ":" + intConf.passWord;
                authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo));
                webClient.Headers["Authorization"] = "Basic " + authInfo;
                uri = uri + "?fromDate=" + _enquiryFromDatetime + "&toDate=" + _enquiryToDatetime;

                Console.WriteLine("starting download " + DateTime.Now);
                Console.WriteLine("final Url " + uri);
                try
                {
                    var jsonData = webClient.DownloadData(uri);
                    DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(RootObjectProduct));
                    RootObjectProduct rootObject = (RootObjectProduct)ser.ReadObject(new MemoryStream(jsonData));

                    Console.WriteLine("data downloaded " + DateTime.Now);

                    int recCount = rootObject.Results.Count;
                    int iDone = 0;
                    foreach (Product obj in rootObject.Results)
                    {
                        Console.WriteLine("DateTime: " + obj.CTIMESTAMP + " : " + obj.MATNR);
                        //if (obj.MATNR != "CHHELLDYHK03240")
                        //    continue;
                        try
                        {
                            iDone++;
                            string dm_Div = String.Empty;
                            if (obj.dm_Div != "")
                                dm_Div = obj.dm_Div;
                            string EAN11 = String.Empty;
                            if (obj.EAN11 != "")
                                EAN11 = obj.EAN11;
                            string EWBEZ = String.Empty;
                            if (obj.EWBEZ != "")
                                EWBEZ = obj.EWBEZ;
                            string EXTWG = String.Empty;
                            if (obj.EXTWG != "")
                                EXTWG = obj.EXTWG;
                            string GEWEI = String.Empty;
                            if (obj.GEWEI != "")
                                GEWEI = obj.GEWEI;

                            string KONDM = String.Empty;
                            if (obj.KONDM != "")
                                KONDM = obj.KONDM;

                            string MAKTX = String.Empty;
                            if (obj.MAKTX != "")
                                MAKTX = obj.MAKTX;

                            string MATNR = String.Empty;
                            if (obj.MATNR != "")
                                MATNR = obj.MATNR;

                            string MHDHB = String.Empty;
                            if (obj.MHDHB != "")
                                MHDHB = obj.MHDHB;

                            string MSTAE = String.Empty;
                            if (obj.MSTAE != "")
                                MSTAE = obj.MSTAE;

                            string MTART = String.Empty;
                            if (obj.MTART != "")
                                MTART = obj.MTART;

                            string MVGR1 = String.Empty;
                            if (obj.MVGR1 != "")
                                MVGR1 = obj.MVGR1;

                            string MVGR1_DESC = String.Empty;
                            if (obj.MVGR1_DESC != "")
                                MVGR1_DESC = obj.MVGR1_DESC;

                            string MVGR2 = String.Empty;
                            if (obj.MVGR2 != "")
                                MVGR2 = obj.MVGR2;
                            string MVGR2_DESC = String.Empty;
                            if (obj.MVGR2_DESC != "")
                                MVGR2_DESC = obj.MVGR2_DESC;
                            string MVGR3 = String.Empty;
                            if (obj.MVGR3 != "")
                                MVGR3 = obj.MVGR3;
                            string MVGR3_DESC = String.Empty;
                            if (obj.MVGR3_DESC != "")
                                MVGR3_DESC = obj.MVGR3_DESC;
                            String MVGR4 = String.Empty;//
                            if (obj.MVGR4 != null)
                                MVGR4 = obj.MVGR4.ToString();
                            String MVGR4_DESC = String.Empty;
                            if (obj.MVGR4_DESC != null)
                                MVGR4_DESC = obj.MVGR4_DESC.ToString();
                            string MVGR5 = String.Empty;
                            if (obj.MVGR5 != "")
                                MVGR5 = obj.MVGR5;
                            string MVGR5_DESC = String.Empty;
                            if (obj.MVGR5_DESC != "")
                                MVGR5_DESC = obj.MVGR5_DESC;
                            string NTGEW = String.Empty;
                            if (obj.NTGEW != "")
                                NTGEW = obj.NTGEW;
                            string ProductLine_EWBEZ = String.Empty;
                            if (obj.ProductLine_EWBEZ != "")
                                ProductLine_EWBEZ = obj.ProductLine_EWBEZ;
                            string SAP_DIV = String.Empty;
                            if (obj.SAP_DIV != "")
                                SAP_DIV = obj.SAP_DIV;
                            string sap_div_desc = String.Empty;
                            if (obj.sap_div_desc != "")
                                sap_div_desc = obj.sap_div_desc;
                            string SERNP = String.Empty;
                            if (obj.SERNP != "")
                                SERNP = obj.SERNP;
                            string STEUC = String.Empty;
                            if (obj.STEUC != "")
                                STEUC = obj.STEUC;
                            string VKORG = String.Empty;
                            if (obj.VKORG != "")
                                VKORG = obj.VKORG;
                            string delete_flag = String.Empty;
                            if (obj.DELETE_FLAG != "")
                                delete_flag = obj.DELETE_FLAG;
                            string CreatedBY = String.Empty;
                            if (obj.CREATEDBY != "")
                                CreatedBY = obj.CREATEDBY;
                            string WGBEZ = String.Empty;
                            if (obj.WGBEZ != "")
                                WGBEZ = obj.WGBEZ;
                            string MATKL = String.Empty;
                            if (obj.MATKL != "")
                                MATKL = obj.MATKL;
                            Guid productId = Models.GetGuidbyNameCommon("product",
                                  "hil_uniquekey", obj.MATNR, service, 1);

                            if (delete_flag == "X")
                            {
                                Console.WriteLine("product fetch " + DateTime.Now);
                                if (productId != Guid.Empty)
                                {
                                    SetStateRequest state = new SetStateRequest();
                                    state.State = new OptionSetValue(1);
                                    state.Status = new OptionSetValue(2);
                                    state.EntityMoniker = new EntityReference("product", productId);
                                    SetStateResponse stateSet = (SetStateResponse)service.Execute(state);
                                    continue;
                                }
                            }

                            Entity product = new Entity("product");
                            EntityReference er;
                            string UqKey = obj.MATNR;
                            if (productId == Guid.Empty)
                            {
                                if (MATNR != String.Empty)
                                {
                                    product.Attributes["hil_uniquekey"] = UqKey;
                                    product.Attributes["name"] = MATNR;
                                    product.Attributes["productnumber"] = MATNR;
                                    product.Attributes["hil_productcode"] = MATNR;
                                }
                                if (STEUC != String.Empty)
                                {
                                    EntityReference HSN = Models.GetHSNCodeRefrence(STEUC, service);
                                    product.Attributes["hil_staginghsncode"] = STEUC;
                                    if (HSN != null)
                                    {
                                        product.Attributes["hil_hsncode"] = HSN;
                                    }
                                }
                                if (ProductLine_EWBEZ != String.Empty)
                                    product.Attributes["hil_stagingsbu"] = ProductLine_EWBEZ;
                                if (MTART != String.Empty)
                                    product.Attributes["hil_stagingmtart"] = MTART;
                                if (MVGR5_DESC != String.Empty)
                                {
                                    Guid mg5 = Models.GetGuidbyNameCommon("hil_materialgroup5",
                                        "hil_name", obj.MVGR5_DESC, service, 1);
                                    if (mg5 == Guid.Empty && obj.MVGR5 != string.Empty)
                                    {
                                        mg5 = Models.GetGuidbyNameCommon("hil_materialgroup5",
                                        "hil_code", obj.MVGR5, service, 1);
                                    }
                                    if (mg5 != Guid.Empty) product.Attributes["hil_materialgroup5"] = new EntityReference("hil_materialgroup5", mg5);
                                }
                                if (MVGR4_DESC != String.Empty)
                                {
                                    Guid mg4 = Models.GetGuidbyNameCommon("hil_materialgroup4",
                                         "hil_name", obj.MVGR4_DESC, service, 1);
                                    if (mg4 == Guid.Empty && obj.MVGR4 != string.Empty)
                                    {
                                        mg4 = Models.GetGuidbyNameCommon("hil_materialgroup4",
                                        "hil_code", obj.MVGR4, service, 1);
                                    }
                                    if (mg4 != Guid.Empty)
                                        product.Attributes["hil_materialgroup4"] = new EntityReference("hil_materialgroup4",
                                            mg4);
                                }
                                if (MVGR3_DESC != String.Empty)
                                {
                                    Guid mg3 = Models.GetGuidbyNameCommon("hil_materialgroup3",
                                        "hil_name", obj.MVGR3_DESC, service, 1);
                                    if (mg3 == Guid.Empty && obj.MVGR3 != string.Empty)
                                    {
                                        mg3 = Models.GetGuidbyNameCommon("hil_materialgroup3",
                                        "hil_code", obj.MVGR3, service, 1);
                                    }
                                    if (mg3 != Guid.Empty)
                                        product.Attributes["hil_materialgroup3"] = new EntityReference("hil_materialgroup3",
                                            mg3);
                                }
                                if (MVGR2_DESC != String.Empty)
                                {
                                    Guid mg2 = Models.GetGuidbyNameCommon("hil_materialgroup2",
                                        "hil_name", obj.MVGR2_DESC, service, 1);
                                    if (mg2 == Guid.Empty && obj.MVGR2 != string.Empty)
                                    {
                                        mg2 = Models.GetGuidbyNameCommon("hil_materialgroup2",
                                        "hil_code", obj.MVGR2, service, 1);
                                    }
                                    if (mg2 != Guid.Empty)
                                        product.Attributes["hil_materialgroup2"] = new EntityReference("hil_materialgroup2",
                                            mg2);
                                }
                                if (MAKTX != String.Empty)
                                {
                                    product.Attributes["description"] = MAKTX;
                                    product.Attributes["suppliername"] = MAKTX;
                                }
                                product.Attributes["defaultuomscheduleid"] = new EntityReference("uomschedule", new Guid("af39a94c-f79f-4e6d-9a9e-20f2948fe185"));
                                product.Attributes["defaultuomid"] = new EntityReference("uom", new Guid("0359d51b-d7cf-43b1-87f6-fc13a2c1dec8"));
                                if (dm_Div != String.Empty)
                                {
                                    Guid _divId = Models.GetGuidbyNameCommon("product", "hil_stagingdivision", obj.dm_Div, service, 2);
                                    if (_divId != Guid.Empty)
                                    {
                                        product.Attributes["hil_division"] = new EntityReference("product", _divId);
                                    }
                                }
                                if (MVGR1 != String.Empty)
                                {
                                    Guid _matGroup1Id = Models.GetGuidbyNameCommon("product", "name", obj.MVGR1, service, 1);
                                    if (_matGroup1Id != Guid.Empty)
                                    {
                                        product.Attributes["hil_materialgroup1"] = new EntityReference("product", _matGroup1Id);
                                    }
                                }
                                if (obj.MATKL != String.Empty)
                                {
                                    Guid _matGroupId = Models.GetGuidbyNameCommon("product", "productnumber", obj.MATKL, service, 1);
                                    if (_matGroupId != Guid.Empty)
                                    {
                                        product.Attributes["hil_materialgroup"] = new EntityReference("product", _matGroupId);
                                    }
                                }
                                if (EXTWG != String.Empty)
                                {
                                    Guid _sbuId = Models.GetGuidbyNameCommon("product", "productnumber", obj.EXTWG, service, 1);
                                    if (_sbuId != Guid.Empty)
                                    {
                                        product.Attributes["hil_sbu"] = new EntityReference("product", _sbuId);
                                    }
                                }
                                if (MVGR4_DESC != String.Empty)
                                    product.Attributes["hil_stagingmaterialgroup4"] = MVGR4_DESC;
                                if (MVGR3_DESC != String.Empty)
                                    product.Attributes["hil_stagingmaterialgroup3"] = MVGR3_DESC;
                                if (MVGR2_DESC != String.Empty)
                                    product.Attributes["hil_stagingmaterialgroup2"] = MVGR2_DESC;
                                if (MVGR1_DESC != String.Empty)
                                    product.Attributes["hil_stagingmaterialgroup1"] = MVGR1_DESC;
                                if (sap_div_desc != String.Empty)
                                    product.Attributes["hil_stagingdivision"] = sap_div_desc;
                                if (MVGR5_DESC != String.Empty)
                                    product.Attributes["hil_stagingmaterialgroup5"] = MVGR5_DESC;

                                product.Attributes["hil_hierarchylevel"] = new OptionSetValue(5);

                                if (obj.MTIMESTAMP == null)
                                    product.Attributes["hil_mdmtimestamp"] = Models.ConvertToDateTime(obj.CTIMESTAMP);
                                else
                                    product.Attributes["hil_mdmtimestamp"] = Models.ConvertToDateTime(obj.MTIMESTAMP);
                                try
                                {
                                    service.Create(product);
                                    Console.WriteLine(MATNR + " Product is created");
                                }
                                catch (Exception ex) { Console.WriteLine(MATNR + " Record has an Error. " + ex.Message); }
                            }
                            else
                            {

                                if (MAKTX != String.Empty)
                                {
                                    Entity entity = new Entity("product", productId);
                                    entity.Attributes["description"] = MAKTX;
                                    try
                                    {
                                        if (STEUC != String.Empty)
                                        {
                                            EntityReference HSN = Models.GetHSNCodeRefrence(STEUC, service);
                                            entity.Attributes["hil_staginghsncode"] = STEUC;
                                            if (HSN != null)
                                            {
                                                product.Attributes["hil_hsncode"] = HSN;
                                            }
                                        }
                                        service.Update(entity);
                                        Console.WriteLine(MATNR + " Record is updated.");
                                    }
                                    catch (Exception ex) { Console.WriteLine(MATNR + " Record has an Error. " + ex.Message); }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("ERROR!!! " + ex.Message);
                        }
                        Console.WriteLine("Records has ben processed " + iDone.ToString());
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.WriteLine("******************************* syncProducts Ended ******************************* ");
        }
        
    }
}
