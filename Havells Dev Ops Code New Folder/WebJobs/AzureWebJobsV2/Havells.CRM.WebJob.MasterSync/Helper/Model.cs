using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.WebJob.MasterSync.Helper
{
    public class IntegrationConfig
    {
        public string uri { get; set; }
        public string Auth { get; set; }
    }
    #region state
    public class StateResult
    {
        public string eff_frmdt { get; set; }
        public string eff_todt { get; set; }
        public string GSTStateName { get; set; }
        public string GSTStateCode { get; set; }
        public string dm_state { get; set; }
        public string SAP_state { get; set; }
        public string sap_state_desc { get; set; }
        public string delete_flag { get; set; }
        public string CreatedBY { get; set; }
        public string Ctimestamp { get; set; }
        public string ModifyBy { get; set; }
        public string Mtimestamp { get; set; }
    }
    public class StateRoot
    {
        public StateResult Result { get; set; }
        public List<StateResult> Results { get; set; }
        public bool Success { get; set; }
        public string Message { get; set; }
        public string ErrorCode { get; set; }
    }
    #endregion state

    #region city
    public class CityResult
    {
        public string eff_frmdt { get; set; }
        public string eff_todt { get; set; }
        public string dm_city { get; set; }
        public string SAP_city { get; set; }
        public string sap_city_desc { get; set; }
        public string delete_flag { get; set; }
        public string CreatedBY { get; set; }
        public string Ctimestamp { get; set; }
        public string ModifyBy { get; set; }
        public string Mtimestamp { get; set; }
        public object ZPOPU { get; set; }
    }

    public class CityRoot
    {
        public object Result { get; set; }
        public List<CityResult> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }
    #endregion city

    #region District
    public class DistrictResult
    {
        public string eff_frmdt { get; set; }
        public string eff_todt { get; set; }
        public string dm_dist { get; set; }
        public string SAP_dist { get; set; }
        public string sap_dist_desc { get; set; }
        public string delete_flag { get; set; }
        public string CreatedBY { get; set; }
        public string Ctimestamp { get; set; }
        public string ModifyBy { get; set; }
        public string Mtimestamp { get; set; }
    }

    public class DistrictRoot
    {
        public object Result { get; set; }
        public List<DistrictResult> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }

    #endregion District

    #region PIN
    public class PinResult
    {
        public string eff_frmdt { get; set; }
        public string eff_todt { get; set; }
        public string dm_pin { get; set; }
        public string delete_flag { get; set; }
        public string CreatedBY { get; set; }
        public string Ctimestamp { get; set; }
        public string ModifyBy { get; set; }
        public string Mtimestamp { get; set; }
    }

    public class PinRoot
    {
        public object Result { get; set; }
        public List<PinResult> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }
    #endregion PIN

    #region salesoffice
    public class SalesOfficeResult
    {
        public string DM_SALES_OFFICE { get; set; }
        public DateTime EFF_FRMDT { get; set; }
        public string SAP_SALES_OFFICE { get; set; }
        public string SAP_SALES_OFFICE_DESC { get; set; }
        public DateTime EFF_TODT { get; set; }
        public string DELETE_FLAG { get; set; }
        public string CREATEDBY { get; set; }
        public string CTIMESTAMP { get; set; }
        public string MODIFYBY { get; set; }
        public string MTIMESTAMP { get; set; }
    }

    public class SalesOfficeRoot
    {
        public object Result { get; set; }
        public List<SalesOfficeResult> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }
    #endregion salesoffice

    #region subterritory
    public class SubTerritoryResult
    {
        public string EFF_FRMDT { get; set; }
        public string EFF_TODT { get; set; }
        public string DM_SUB_TER { get; set; }
        public string DM_SUB_TER_DESC { get; set; }
        public string DELETE_FLAG { get; set; }
        public string CREATEDBY { get; set; }
        public string CTIMESTAMP { get; set; }
        public string MODIFYBY { get; set; }
        public string MTIMESTAMP { get; set; }
    }

    public class SubTerritoryRoot
    {
        public object Result { get; set; }
        public List<SubTerritoryResult> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }

    #endregion subterritory

    #region Region
    public class RegionResult
    {
        public string eff_frmdt { get; set; }
        public string eff_todt { get; set; }
        public string dm_region { get; set; }
        public string SAP_region { get; set; }
        public string sap_region_desc { get; set; }
        public string delete_flag { get; set; }
        public string CreatedBY { get; set; }
        public string Ctimestamp { get; set; }
        public string ModifyBy { get; set; }
        public string Mtimestamp { get; set; }
    }

    public class RegionRoot
    {
        public object Result { get; set; }
        public List<RegionResult> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }
    #endregion Region

    #region Branch
    public class BranchResult
    {
        public string eff_frmdt { get; set; }
        public string eff_todt { get; set; }
        public string dm_branch { get; set; }
        public string dm_branch_desc { get; set; }
        public string delete_flag { get; set; }
        public string CreatedBY { get; set; }
        public string Ctimestamp { get; set; }
        public string ModifyBy { get; set; }
        public string Mtimestamp { get; set; }
    }

    public class BranchRoot
    {
        public object Result { get; set; }
        public List<BranchResult> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }
    #endregion Branch

    #region Area
    public class AreaResult
    {
        public string eff_frmdt { get; set; }
        public string eff_todt { get; set; }
        public string DM_AREA { get; set; }
        public string dm_area_desc { get; set; }
        public string delete_flag { get; set; }
        public string CreatedBY { get; set; }
        public string Ctimestamp { get; set; }
        public string ModifyBy { get; set; }
        public string Mtimestamp { get; set; }
    }

    public class AreaRoot
    {
        public object Result { get; set; }
        public List<AreaResult> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }

    #endregion Area

    #region BusinessMapping
    public class BusinessMappingResult
    {
        public string dm_pin_efdt { get; set; }
        public string DM_AREA_EFDT { get; set; }
        public string dm_city_EFDT { get; set; }
        public string dm_dist_EFDT { get; set; }
        public string DM_SALES_OFFICE_EFDT { get; set; }
        public string DM_BRANCH_EFDT { get; set; }
        public string dm_region_EFDT { get; set; }
        public string eff_frmdt { get; set; }
        public string eff_todt { get; set; }
        public string dm_area_desc { get; set; }
        public string sap_city_desc { get; set; }
        public string sap_dist_desc { get; set; }
        public string SAP_SALES_OFFICE_DESC { get; set; }
        public string dm_branch_desc { get; set; }
        public string sap_region_desc { get; set; }
        public string sap_state_desc { get; set; }
        public string DM_SUB_TER_DESC { get; set; }
        public object dm_sub_ter_efdt { get; set; }
        public object dm_state_efdt { get; set; }
        public string dm_pin { get; set; }
        public string DM_AREA { get; set; }
        public string dm_city { get; set; }
        public string dm_dist { get; set; }
        public string DM_SALES_OFFICE { get; set; }
        public string DM_BRANCH { get; set; }
        public string dm_region { get; set; }
        public string dm_state { get; set; }
        public string dm_sub_ter { get; set; }
        public string delete_flag { get; set; }
        public string CreatedBY { get; set; }
        public string Ctimestamp { get; set; }
        public string ModifyBy { get; set; }
        public string Mtimestamp { get; set; }
    }

    public class BusinessMappingRoot
    {
        public object Result { get; set; }
        public List<BusinessMappingResult> Results { get; set; }
        public bool Success { get; set; }
        public object Message { get; set; }
        public object ErrorCode { get; set; }
    }

    #endregion BusinessMapping
}
