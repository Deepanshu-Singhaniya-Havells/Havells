using System;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using System.Globalization;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using Havells_Plugin.HelperIntegration;

namespace Havells_Plugin.StagingJSONPartner
{
    public class PostCreate_Async : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            #region SetUp
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity
                    && context.PrimaryEntityName.ToLower() == new_integrationstaging.EntityLogicalName && context.MessageName.ToUpper() == "CREATE")
                {
                    Entity entity = (Entity)context.InputParameters["Target"];
                    new_integrationstaging Staging = entity.ToEntity<new_integrationstaging>();
                    switch (Staging.hil_Type.Value)
                    {
                        case 1: 
                        ResolvePartner(Staging, service);
                            break;
                        case 3:
                            ResolveBIZGeoMapping(Staging, service);
                            break;

                    }
                   
                }
            }
            catch(Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.StagingJSONPartner.Execute : "+ex.Message);
            }
            
        }

        public static void ResolvePartner(new_integrationstaging Stage , IOrganizationService service)
        {
            try
            {
                if (Stage.new_Message != null)
                {
                        var jsonData = Stage.new_Message;
                        PartnerRootObject rootObject = JsonConvert.DeserializeObject<PartnerRootObject>(jsonData);
                    if (rootObject.KUNNR != null)
                        {
                            string _iCode = string.Empty;
                            string _oCode = string.Empty;
                            string _pCode = rootObject.KUNNR;
                            if (_pCode.StartsWith("F") || _pCode.StartsWith("f"))
                            {
                                _oCode = _pCode;
                                for (int i = 1; i < _pCode.Length; i++) { _iCode = _iCode + _pCode[i]; }
                            }  
                            else
                            {
                                _iCode = _pCode;
                                _oCode = "F" + _pCode;
                            }
                        if (_iCode != null)
                            {
                                bool IfExists = CheckIfPartnerExists(service, _iCode, _oCode);
                            if (!IfExists)
                                {
                                    Guid AccId = CreateAccount(service, rootObject);
                                if (AccId != Guid.Empty)
                                    {

                                    }
                                }
                            }
                        }
                    }
                }
            catch (Exception ex)
            { }

            }


        public static bool CheckIfPartnerExists(IOrganizationService service, string _iCode, string _oCode)
        {
            bool IfExists = false;
            try
            {
                QueryExpression Query = new QueryExpression(Account.EntityLogicalName);
                Query.ColumnSet = new ColumnSet(false);
                Query.Criteria = new FilterExpression(LogicalOperator.Or);
                Query.Criteria.AddCondition(new ConditionExpression("hil_inwarrantycustomersapcode", ConditionOperator.Equal, _iCode));
                Query.Criteria.AddCondition(new ConditionExpression("hil_outwarrantycustomersapcode", ConditionOperator.Equal, _oCode));
                EntityCollection Found = service.RetrieveMultiple(Query);
                if (Found.Entities.Count > 0)
                {
                    IfExists = true;
                    Account _tPrtnr = (Account)Found.Entities[0];
                    _tPrtnr.hil_InWarrantyCustomerSAPCode = _iCode;
                    _tPrtnr.hil_OutWarrantyCustomerSAPCode = _oCode;
                    service.Update(_tPrtnr);
                }
            }
            catch(Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.StagingJSONPartner.CheckIfPartnerExists : " + ex.Message);
            }
            return IfExists;
        }
        public static Guid CreateAccount(IOrganizationService service, PartnerRootObject rootObject)
        {
            Guid AccountId = new Guid();
            try
            {
                Guid Area = Guid.Empty;
                if(rootObject.DM_AREA!="" && rootObject.DM_AREA!=null)
                Area = Helper.GetGuidbyName(hil_area.EntityLogicalName, "hil_areacode", rootObject.DM_AREA, service);
                Guid Region = Helper.GetGuidbyName(hil_region.EntityLogicalName, "hil_regioncode", rootObject.DM_REGION, service);
                Guid Branch = Helper.GetGuidbyName(hil_branch.EntityLogicalName, "hil_branchcode", rootObject.DM_BRANCH, service);
                Guid SalesOffice = Helper.GetGuidbyName(hil_salesoffice.EntityLogicalName, "hil_salesofficecode", rootObject.DM_SALES_OFFICE, service);
                Guid City = Helper.GetGuidbyName(hil_city.EntityLogicalName, "hil_citycode", rootObject.dm_city, service);
                Guid State = Helper.GetGuidbyName(hil_state.EntityLogicalName, "hil_statecode", rootObject.dm_state, service);
                Guid PinCode = Helper.GetGuidbyName(hil_pincode.EntityLogicalName, "hil_name", rootObject.dm_pin, service);
                Guid District = Helper.GetGuidbyName(hil_district.EntityLogicalName, "hil_districtcode", rootObject.dm_dist, service);
                Account _tPrtner = new Account();
                _tPrtner.CustomerTypeCode = new OptionSetValue(6);
                if(rootObject.VTXTM != null)
                    _tPrtner.Name = rootObject.VTXTM;
                if(rootObject.STREET != null)
                    _tPrtner.Address1_Line1 = rootObject.STREET;
                if(rootObject.STR_SUPPL3 != null)
                    _tPrtner.Address1_Line2 = rootObject.STR_SUPPL3;
                if(rootObject.ADDRESS3 != null)
                    _tPrtner.Address1_Line3 = rootObject.ADDRESS3;
                if(rootObject.SMTP_ADDR != null)
                    _tPrtner.EMailAddress1 = rootObject.SMTP_ADDR;
                if(rootObject.MOB_NUMBER != null)
                    _tPrtner.Telephone1 = rootObject.MOB_NUMBER;
                if(rootObject.J_1IPANNO != null)
                    _tPrtner.hil_pan = rootObject.J_1IPANNO;
                if(Area != Guid.Empty)
                {
                    _tPrtner.hil_area = new EntityReference(hil_area.EntityLogicalName, Area);
                }
                if (Region != Guid.Empty)
                {
                    _tPrtner.hil_region = new EntityReference(hil_region.EntityLogicalName, Region);
                }
                if (Branch != Guid.Empty)
                {
                    _tPrtner.hil_branch = new EntityReference(hil_branch.EntityLogicalName, Branch);
                }
                if (SalesOffice != Guid.Empty)
                {
                    _tPrtner.hil_salesoffice = new EntityReference(hil_salesoffice.EntityLogicalName, SalesOffice);
                }
                if (City != Guid.Empty)
                {
                    _tPrtner.hil_city = new EntityReference(hil_city.EntityLogicalName, City);
                }
                if (State != Guid.Empty)
                {
                    _tPrtner.hil_state = new EntityReference(hil_state.EntityLogicalName, State);
                }
                if (PinCode != Guid.Empty)
                {
                    _tPrtner.hil_pincode = new EntityReference(hil_pincode.EntityLogicalName, PinCode);
                }
                if (District != Guid.Empty)
                {
                    _tPrtner.hil_district = new EntityReference(hil_district.EntityLogicalName, District);
                }
                if (rootObject.KUNNR.StartsWith("F"))
                {
                    _tPrtner.hil_OutWarrantyCustomerSAPCode = rootObject.KUNNR;
                }
                else
                {
                    _tPrtner.hil_InWarrantyCustomerSAPCode = rootObject.KUNNR;
                }
                AccountId = service.Create(_tPrtner);
            }
            catch(Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.StagingJSONPartner.CreateAccount : " + ex.Message);
            }
            return AccountId;
        }

        public static void ResolveBIZGeoMapping(new_integrationstaging Record, IOrganizationService service)
        {
            try
            {
                using (OrganizationServiceContext orgContext = new OrganizationServiceContext(service))
                {
                    if (Record.new_Message != null)
                    {
                        var jsonData = Record.new_Message;
                        BIZGeoMapModelClass obj = JsonConvert.DeserializeObject<BIZGeoMapModelClass>(jsonData);

                        string dm_pin_efdt = String.Empty;
                        if (obj.dm_pin_efdt != "")
                            dm_pin_efdt = obj.dm_pin_efdt;
                        string DM_AREA_EFDT = String.Empty;
                        if (obj.DM_AREA_EFDT != "")
                            DM_AREA_EFDT = obj.DM_AREA_EFDT;
                        string dm_city_EFDT = String.Empty;
                        if (obj.dm_city_EFDT != "")
                            dm_city_EFDT = obj.dm_city_EFDT;
                        string dm_dist_EFDT = String.Empty;
                        if (obj.dm_dist_EFDT != "")
                            dm_dist_EFDT = obj.dm_dist_EFDT;
                        string DM_SALES_OFFICE_EFDT = String.Empty;
                        if (obj.DM_SALES_OFFICE_EFDT != "")
                            DM_SALES_OFFICE_EFDT = obj.DM_SALES_OFFICE_EFDT;

                        string DM_BRANCH_EFDT = String.Empty;
                        if (obj.DM_BRANCH_EFDT != "")
                            DM_BRANCH_EFDT = obj.DM_BRANCH_EFDT;

                        string dm_region_EFDT = String.Empty;
                        if (obj.dm_region_EFDT != "")
                            dm_region_EFDT = obj.dm_region_EFDT;

                        string eff_frmdt = String.Empty;
                        if (obj.eff_frmdt != "")
                            eff_frmdt = obj.eff_frmdt;

                        string eff_todt = String.Empty;
                        if (obj.eff_todt != "")
                            eff_todt = obj.eff_todt;

                        string dm_area_desc = String.Empty;
                        if (obj.dm_area_desc != "")
                            dm_area_desc = obj.dm_area_desc;

                        string sap_city_desc = String.Empty;
                        if (obj.sap_city_desc != "")
                            sap_city_desc = obj.sap_city_desc;

                        string sap_dist_desc = String.Empty;
                        if (obj.sap_dist_desc != "")
                            sap_dist_desc = obj.sap_dist_desc;

                        string SAP_SALES_OFFICE_DESC = String.Empty;
                        if (obj.SAP_SALES_OFFICE_DESC != "")
                            SAP_SALES_OFFICE_DESC = obj.SAP_SALES_OFFICE_DESC;

                        string dm_branch_desc = String.Empty;
                        if (obj.dm_branch_desc != "")
                            dm_branch_desc = obj.dm_branch_desc;
                        string sap_region_desc = String.Empty;
                        if (obj.sap_region_desc != "")
                            sap_region_desc = obj.sap_region_desc;
                        string sap_state_desc = String.Empty;
                        if (obj.sap_state_desc != "")
                            sap_state_desc = obj.sap_state_desc;
                        string DM_SUB_TER_DESC = String.Empty;
                        if (obj.DM_SUB_TER_DESC != "")
                            DM_SUB_TER_DESC = obj.DM_SUB_TER_DESC;
                        String dm_sub_ter_efdt = String.Empty;//
                        if (obj.dm_sub_ter_efdt != null)
                            dm_sub_ter_efdt = obj.dm_sub_ter_efdt.ToString();
                        String dm_state_efdt = String.Empty;
                        if (obj.dm_state_efdt != null)
                            dm_state_efdt = obj.dm_state_efdt.ToString();
                        string dm_pin = String.Empty;
                        if (obj.dm_pin != "")
                            dm_pin = obj.dm_pin;
                        string DM_AREA = String.Empty;
                        if (obj.DM_AREA != "")
                            DM_AREA = obj.DM_AREA;
                        string dm_city = String.Empty;
                        if (obj.dm_city != "")
                            dm_city = obj.dm_city;
                        string dm_dist = String.Empty;
                        if (obj.dm_dist != "")
                            dm_dist = obj.dm_dist;
                        string DM_SALES_OFFICE = String.Empty;
                        if (obj.DM_SALES_OFFICE != "")
                            DM_SALES_OFFICE = obj.DM_SALES_OFFICE;
                        string DM_BRANCH = String.Empty;
                        if (obj.DM_BRANCH != "")
                            DM_BRANCH = obj.DM_BRANCH;
                        string dm_region = String.Empty;
                        if (obj.dm_region != "")
                            dm_region = obj.dm_region;
                        string dm_state = String.Empty;
                        if (obj.dm_state != "")
                            dm_state = obj.dm_state;
                        string dm_sub_ter = String.Empty;
                        if (obj.dm_sub_ter != "")
                            dm_sub_ter = obj.dm_sub_ter;
                        string delete_flag = String.Empty;
                        if (obj.delete_flag != "")
                            delete_flag = obj.delete_flag;
                        string CreatedBY = String.Empty;
                        if (obj.CreatedBY != "")
                            CreatedBY = obj.CreatedBY;

                        if (delete_flag == "X")
                        {
                            Guid fsBigGeoId = GetGuidbyNameCommon(hil_businessmapping.EntityLogicalName, "hil_uniquekey", dm_pin + dm_pin_efdt + DM_AREA + DM_AREA_EFDT + eff_frmdt, service, -1);
                            if (fsBigGeoId != Guid.Empty)
                            {
                                //statuscode active 1 inactive 0  statecode active 0 inactive 1

                                SetStateRequest state = new SetStateRequest();
                                state.State = new OptionSetValue(1);
                                state.Status = new OptionSetValue(0);
                                state.EntityMoniker = new EntityReference(hil_businessmapping.EntityLogicalName, fsBigGeoId);
                                SetStateResponse stateSet = (SetStateResponse)service.Execute(state);

                            }
    }
                        else
                        {

                            hil_businessmapping crBigGeoMap = new hil_businessmapping();

                            if (eff_frmdt != String.Empty)
                            {
                                DateTime dt = DateTime.ParseExact(eff_frmdt, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                                crBigGeoMap.hil_effectivefromdate = dt;
                            }
                            if (eff_todt != String.Empty)
                            {
                                DateTime dt = DateTime.ParseExact(eff_todt, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                                crBigGeoMap.hil_effectivetodate = dt;
                            }
                            crBigGeoMap.hil_uniquekey = dm_pin + dm_pin_efdt + DM_AREA + DM_AREA_EFDT + eff_frmdt;
                            if (dm_pin != String.Empty)
                                crBigGeoMap.hil_StagingPIN = dm_pin;
                            if (dm_pin != String.Empty)
                                crBigGeoMap.hil_StagingPIN = dm_pin;
                            if (dm_pin_efdt != String.Empty)
                                crBigGeoMap.hil_StagingPINeffDate = dm_pin_efdt;
                            if (DM_AREA != String.Empty)
                                crBigGeoMap.hil_StagingArea = DM_AREA;
                            if (DM_AREA_EFDT != String.Empty)
                                crBigGeoMap.hil_StagingAreaEffDate = DM_AREA_EFDT;
                            if (dm_city != String.Empty)
                                crBigGeoMap.hil_StagingCity = dm_city;
                            if (dm_city_EFDT != String.Empty)//
                                crBigGeoMap.hil_StagingCItyEffDate = dm_city_EFDT;
                            if (dm_dist != String.Empty)
                                crBigGeoMap.hil_StagingDistrict = dm_dist;
                            if (dm_dist_EFDT != String.Empty)
                                crBigGeoMap.hil_StagingDistrictEffDate = dm_dist_EFDT;
                            if (dm_sub_ter != String.Empty)
                                crBigGeoMap.hil_StagingTerritory = dm_sub_ter;
                            if (dm_sub_ter_efdt != String.Empty)
                                crBigGeoMap.hil_StagingTerritoryEffDt = dm_sub_ter_efdt;
                            if (dm_state != String.Empty)
                                crBigGeoMap.hil_StagingState = dm_state;
                            if (dm_city != String.Empty)
                                crBigGeoMap.hil_StagingCity = dm_city;
                            if (dm_state_efdt != String.Empty)
                                crBigGeoMap.hil_StagingStateEffDt = dm_state_efdt;
                            if (DM_SALES_OFFICE != String.Empty)
                                crBigGeoMap.hil_StagingSalesOffice = DM_SALES_OFFICE;
                            if (DM_SALES_OFFICE_EFDT != String.Empty)
                                crBigGeoMap.hil_StagingSalesOfficeEffDt = DM_SALES_OFFICE_EFDT;
                            if (DM_SALES_OFFICE_EFDT != String.Empty)
                                crBigGeoMap.hil_StagingSalesOfficeEffDt = DM_SALES_OFFICE_EFDT;//
                            if (DM_BRANCH != String.Empty)
                                crBigGeoMap.hil_StagingBranch = DM_BRANCH;
                            if (DM_BRANCH_EFDT != String.Empty)
                                crBigGeoMap.hil_StagingBranchEffFrDate = DM_BRANCH_EFDT;
                            if (dm_region != String.Empty)
                                crBigGeoMap.hil_StagingRegion = dm_region;
                            if (dm_region_EFDT != String.Empty)
                                crBigGeoMap.hil_StagingRegionEffDt = dm_region_EFDT;
                            if (DM_SALES_OFFICE_EFDT != String.Empty)
                                crBigGeoMap.hil_StagingSalesOfficeEffDt = DM_SALES_OFFICE_EFDT;

                            crBigGeoMap.hil_StagingPINUniqueKey = dm_pin + dm_pin_efdt;
                            crBigGeoMap.hil_StagingareaUniqueKey = DM_AREA + DM_AREA_EFDT;
                            crBigGeoMap.hil_StagingCityUniqueKey = dm_city + dm_city_EFDT;
                            crBigGeoMap.hil_StagingDistrictUniqueKey = dm_dist + dm_dist_EFDT;
                            crBigGeoMap.hil_StagingTerritoryUniqueKey = dm_sub_ter;
                            crBigGeoMap.hil_StagingStateUniqueKey = dm_state;
                            crBigGeoMap.hil_StagingSalesOfficeUniqueKey = DM_SALES_OFFICE + DM_SALES_OFFICE_EFDT;
                            crBigGeoMap.hil_StagingBranchUniqieKey = DM_BRANCH + DM_BRANCH_EFDT;
                            crBigGeoMap.hil_regionUniqieKey = dm_region + dm_region_EFDT;

                            Guid fsPinCodeId = GetGuidbyNameCommon(hil_pincode.EntityLogicalName, "hil_uniquekey", dm_pin + dm_pin_efdt, service, 1);
                            Guid fsAreaId = GetGuidbyNameCommon(hil_area.EntityLogicalName, "hil_uniquekey", DM_AREA + DM_AREA_EFDT, service, 1);
                            Guid fsCItyId = GetGuidbyNameCommon(hil_city.EntityLogicalName, "hil_uniquekey", dm_city + dm_city_EFDT, service, 1);
                            Guid fsDistrictId = GetGuidbyNameCommon(hil_district.EntityLogicalName, "hil_uniquekey", dm_dist + dm_dist_EFDT, service, 1);

                            Guid fsTerritoryId = GetGuidbyNameCommon(hil_subterritory.EntityLogicalName, "hil_subterritorycode", dm_sub_ter, service, 1);
                            Guid fsStateId = GetGuidbyNameCommon(hil_state.EntityLogicalName, "hil_statecode", dm_state, service, 1);

                            //fix populate 
                            Guid fsSalesOfficeId = GetGuidbyNameCommon(hil_salesoffice.EntityLogicalName, "hil_uniquekey", DM_SALES_OFFICE + DM_SALES_OFFICE_EFDT, service, 1);
                            Guid fsBranchId = GetGuidbyNameCommon(hil_branch.EntityLogicalName, "hil_uniquekey", DM_BRANCH + DM_BRANCH_EFDT, service, 1);
                            Guid fsRegionId = GetGuidbyNameCommon(hil_region.EntityLogicalName, "hil_uniquekey", dm_region + dm_region_EFDT, service, 1);

                            if (fsPinCodeId != Guid.Empty)
                                crBigGeoMap.hil_pincode = new EntityReference(hil_pincode.EntityLogicalName, fsPinCodeId);
                            if (fsAreaId != Guid.Empty)
                                crBigGeoMap.hil_area = new EntityReference(hil_area.EntityLogicalName, fsAreaId);
                            if (fsCItyId != Guid.Empty)
                                crBigGeoMap.hil_city = new EntityReference(hil_city.EntityLogicalName, fsCItyId);
                            if (fsDistrictId != Guid.Empty)
                                crBigGeoMap.hil_district = new EntityReference(hil_district.EntityLogicalName, fsDistrictId);
                            if (fsTerritoryId != Guid.Empty)
                                crBigGeoMap.hil_subterritory = new EntityReference(hil_subterritory.EntityLogicalName, fsTerritoryId);
                            if (fsStateId != Guid.Empty)
                                crBigGeoMap.hil_state = new EntityReference(hil_state.EntityLogicalName, fsStateId);
                            if (fsSalesOfficeId != Guid.Empty)
                                crBigGeoMap.hil_salesoffice = new EntityReference(hil_pincode.EntityLogicalName, fsSalesOfficeId);
                            if (fsBranchId != Guid.Empty)
                                crBigGeoMap.hil_branch = new EntityReference(hil_branch.EntityLogicalName, fsBranchId);
                            if (fsRegionId != Guid.Empty)
                                crBigGeoMap.hil_region = new EntityReference(hil_region.EntityLogicalName, fsRegionId);
                            crBigGeoMap.statuscode = new OptionSetValue(1);//statuscode active 1 inactive 1  statecode active 0 inactive 1
                         Guid g=   service.Create(crBigGeoMap);

                        }


                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        public static Guid GetGuidbyNameCommon(String sEntityName, String sFieldName, String sFieldValue, IOrganizationService service, int iStatusCode = 0)
        {
            Guid fsResult = Guid.Empty;
            try
            {
                QueryExpression qe = new QueryExpression(sEntityName);
                qe.Criteria.AddCondition(sFieldName, ConditionOperator.Equal, sFieldValue);
                qe.AddOrder("createdon", OrderType.Descending);
                if (iStatusCode >= 0)
                {
                    //qe.Criteria.AddCondition("statecode", ConditionOperator.Equal, iStatusCode);
                    qe.Criteria.AddCondition("statuscode", ConditionOperator.Equal, iStatusCode);
                }
                //qe.Criteria.AddCondition("hil_deleteflag", ConditionOperator.NotEqual, 1);
                EntityCollection enColl = service.RetrieveMultiple(qe);
                if (enColl.Entities.Count > 0)
                {
                    fsResult = enColl.Entities[0].Id;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Havells_Plugin.Helper.GetGuidbyName" + ex.Message);
            }
            return fsResult;
        }

    }
} 
