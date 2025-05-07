using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.D365.Entities.Incident.Entity
{
    public class dtoIncidents
    {
        public string createdby { get; set; }
        public string createdbyname { get; set; }
        public string createdbyyominame { get; set; }
        public DateTime? createdon { get; set; }
        public string createdonbehalfby { get; set; }
        public string createdonbehalfbyname { get; set; }
        public string createdonbehalfbyyominame { get; set; }
        public string hil_customer { get; set; }
        public string hil_customeridtype { get; set; }
        public string hil_customername { get; set; }
        public string hil_customeryominame { get; set; }
        public int? hil_inputwatertds { get; set; }
        public bool? hil_iswarrantyvoid { get; set; }
        public string hil_iswarrantyvoidname { get; set; }
        public string hil_modelcode { get; set; }
        public string hil_modelcodename { get; set; }
        public string hil_modeldescription { get; set; }
        public string hil_modelname { get; set; }
        public string hil_natureofcomplaint { get; set; }
        public string hil_natureofcomplaintname { get; set; }
        public string hil_noofpeopleinhome { get; set; }
        public string hil_observation { get; set; }
        public string hil_observationname { get; set; }
        public string hil_ph { get; set; }
        public string hil_productcategory { get; set; }
        public string hil_productcategoryname { get; set; }
        public int? hil_productreplacement { get; set; }
        public string hil_productreplacementname { get; set; }
        public string hil_productsubcategory { get; set; }
        public string hil_productsubcategoryname { get; set; }
        public int? hil_purewatertds { get; set; }
        public int? hil_quantity { get; set; }
        public int? hil_rejectwatertds { get; set; }
        public string hil_serialnumber { get; set; }
        public double? hil_tds { get; set; }
        public int? hil_warrantystatus { get; set; }
        public string hil_warrantystatusname { get; set; }
        public string hil_warrantyvoidreasoncode { get; set; }
        public int? hil_watersource { get; set; }
        public string hil_watersourcename { get; set; }
        public int? hil_waterstoragetype { get; set; }
        public string hil_waterstoragetypename { get; set; }
        public int? importsequencenumber { get; set; }
        public string modifiedby { get; set; }
        public string modifiedbyname { get; set; }
        public string modifiedbyyominame { get; set; }
        public DateTime? modifiedon { get; set; }
        public string modifiedonbehalfby { get; set; }
        public string modifiedonbehalfbyname { get; set; }
        public string modifiedonbehalfbyyominame { get; set; }
        public string msdyn_agreementbookingincident { get; set; }
        public string msdyn_agreementbookingincidentname { get; set; }
        public string msdyn_customerasset { get; set; }
        public string msdyn_customerassetname { get; set; }
        public string msdyn_description { get; set; }
        public int? msdyn_estimatedduration { get; set; }
        public bool? msdyn_incidentresolved { get; set; }
        public string msdyn_incidentresolvedname { get; set; }
        public string msdyn_incidenttype { get; set; }
        public string msdyn_incidenttypename { get; set; }
        public string msdyn_internalflags { get; set; }
        public bool? msdyn_ismobile { get; set; }
        public string msdyn_ismobilename { get; set; }
        public bool? msdyn_isprimary { get; set; }
        public string msdyn_isprimaryname { get; set; }
        public bool? msdyn_itemspopulated { get; set; }
        public string msdyn_itemspopulatedname { get; set; }
        public string msdyn_name { get; set; }
        public string msdyn_resourcerequirement { get; set; }
        public string msdyn_resourcerequirementname { get; set; }
        public float? msdyn_taskspercentcompleted { get; set; }
        public string msdyn_workorder { get; set; }
        public string msdyn_workorderincidentid { get; set; }
        public string msdyn_workordername { get; set; }
        public DateTime? new_warrantyenddate { get; set; }
        public DateTime? overriddencreatedon { get; set; }
        public string ownerid { get; set; }
        public string owneridname { get; set; }
        public string owneridtype { get; set; }
        public string owneridyominame { get; set; }
        public string owningbusinessunit { get; set; }
        public string owningteam { get; set; }
        public string owninguser { get; set; }
        public int? statecode { get; set; }
        public string statecodename { get; set; }
        public int? statuscode { get; set; }
        public string statuscodename { get; set; }
        public int? timezoneruleversionnumber { get; set; }
        public int? utcconversiontimezonecode { get; set; }
        public Int64? versionnumber { get; set; }
        public DateTime? CreatedDateWh { get; set; }


    }
}
