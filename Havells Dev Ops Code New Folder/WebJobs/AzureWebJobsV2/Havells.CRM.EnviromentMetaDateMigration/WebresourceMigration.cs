using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Havells.CRM.EnviromentMetaDateMigration
{
    public class WebresourceMigration
    {
        public static void UpsertWebResource(IOrganizationService _SourceService, IOrganizationService _DestinationService)
        {
            Console.WriteLine("********************** WEBRESOURCE MIGRATION IS STARTED **********************");
            string web = "lat_/CRMRESTBuilder/scripts/Sdk.Soap.min.js;dyn_result_32;dyn_result_16;ispl_AutoNumberIcon_32;ispl_AutoNumberIcon_16;hil_WorkOrderNew.js;hil_WorkOrderProduct;hil_WorkOrder;hil_WOIncident;hil_WOClaims;hil_Webresource/JS/Case.js;hil_WebResource/JS/AdvisoryEnqueryLine.js;hil_WebResource/JS/AdvisoryEnquery.js;hil_WebResource/HTML/Attachment/WebPage.html;hil_WarrantlyDetail;hil_validateAMC;hil_UserSecurityRoleExtension;hil_TimeSlot;hil_TestQuickCreate;hil_Tendertechno1;hil_Tendertechno;hil_TenderProduct;hil_TenderPo1;hil_TenderPo;hil_TenderPaymentdetailsJs;hil_Tendergtpuploaded1;hil_Tendergtpuploaded;hil_Tenderfinaltcp1;hil_Tenderfinaltcp;hil_TenderFinalgtpuploaded1;hil_TenderFinalgtpuploaded;hil_TenderFinal1;hil_TenderFinal;hil_TenderDesigntaskcompleted1;hil_TenderDesigntaskcompleted;hil_TenderButton;hil_BankGuarnteeDetails;hil_tenderbankguarntee1;hil_TenderFormEvents;hil_tenderFieldEvents;hil_technicianJS;hil_TechnicalSpecification;hil_TechnicalLocker;hil_SubmittoDesignteam1;hil_SubmittoDesignteam;hil_style;hil_SolarServey;hil_SMSTemplate;hil_slick;hil_SetSubStatusCanceled_LockJobForm;hil_SetJobIDonload;hil_Scheme_validate;hil_sawActivityApproval;new_Reversegeocoding;hil_ReturnLine;hil_ReturnHeader;hil_WebResource/HTML/RemotePay.html;hil_Reject.svg;hil_qrcode.min.js;hil_qr-code.jpg;hil_purchaseorderimage;hil_ProjectUtils.js;hil_profileimageguidelines;hil_ProductionSupport;hil_ProductRequestsValidation;hil_ProductRequest;hil_PostOrderCheckList;hil_JSPOStatusTracker;hil_POStatusHistoryTracker;hil_PMSvalidate;hil_pmsconfiglines;hil_photo2;hil_PhoneCall;hil_PerformaInvoice.js;hil_part_validate;hil_OrderCheckListProductJS;hil_OrderCheckListJS;hil_OAProductJS;hil_OALineinspectionStatus;hil_OAHeaderAttachment;hil_oAHeader;hil_namejs;hil_MinimunStockLevel;new_map;hil_LocalPurchase;hil_lead;hil_labor_validate;hil_Jquery3.6.0;hil_jquery_1.9.1.min.js;hil_jquery.min.js;hil_JobEstimates;hil_JobCancelIMAGE;hil_JobServices;hil_JobProducts;hil_IncomingCall;hil_imageview;hil_yellowIcon.png;hil_Subgridformatting;hil_redIcon.png;hil_greenIcon.png;hil_HavellsJSLib.js;hil_Havells.WarrantyTemplate.js;hil_Havells.WarrantyTemplate.FieldEvents.js;hil_Havells.Utility.js;hil_GRNLine;new_GetAllEntityList;hil_GeoFencingMapView;hil_EnquiryStatus;hil_ems;hil_DigiLockerTest;hil_DigiLocker;hil_DesignDashboard;hil_DailyHealthIndicatorLine;hil_hil_CustomerAssetPortal;hil_Customerasset;hil_custom;hil_ContactNew.js;hil_Contact;hil_ClaimOverheadLineJs;hil_claimRemarks;hil_Characterstics;hil_care-360.jpg;hil_care-360;hil_CampaignMainLibrary;hil_Bootstrap_min_js_3.3.7;hil_Bootstrap_min_css_3.3.7;hil_bootstrap.min.js;hil_bootstrap.min;hil_BankGurabteeValidation;hil_WebResource/HTML/AuditLog/AuditLogTable.html;hil_Attachment;hil_ArchivedJobView;hil_ArchivedJobJS;hil_AquaParametersValidation;hil_Approve.svg;hil_Approve;hil_Approval;hil_address;hil_ACInstallationCheckList;hil_Account;hil_/css/alert.css;hil_communication;hil_EnquiryBomLine;hil_HealthIndicatorIcon;hil_HavellsLogo;hil_group-logo.jpg;";
            //"hil_validateAMC";// "adx_scripts /jquery1.9.1.min.js;adx_scripts/jquery1.9.1.min.js;msdyn_SDK.REST.js;mag_/js/process.js";
            string[] webs = web.Split(';');

            try
            {
                var query = new QueryExpression("webresource");
                query.Criteria.AddCondition("name", ConditionOperator.In, webs);
                query.ColumnSet = new ColumnSet(true);
                EntityCollection _SourceWebResourceColl = _SourceService.RetrieveMultiple(query);
                int count = _SourceWebResourceColl.Entities.Count;
                int done = 1;
                int error = 1;
                foreach (Entity _SourceWebresource in _SourceWebResourceColl.Entities)
                {
                    try
                    {
                        var query1 = new QueryExpression("webresource");
                        query1.Criteria.AddCondition("name", ConditionOperator.Equal, _SourceWebresource["name"]);
                        query1.ColumnSet = new ColumnSet(true);
                        EntityCollection _DestinationWebResourceColl = _DestinationService.RetrieveMultiple(query1);
                        if (_DestinationWebResourceColl.Entities.Count == 0)
                        {
                            _DestinationService.Create(_SourceWebresource);
                            Console.WriteLine("Done " + done + "/" + count + " Record Created. " + _SourceWebresource["name"]);
                            done++;
                        }
                        else
                        {
                            _DestinationWebResourceColl[0]["content"] = _SourceWebresource["content"];
                            //entity["webresourceidunique"] = results1[0].Id;
                            //entity.Id = results1[0].Id;
                            _DestinationService.Update(_DestinationWebResourceColl[0]);
                            Console.WriteLine("Done " + done + "/" + count + " Record updated. " + _SourceWebresource["name"]);
                            done++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error with webresource name " + _SourceWebresource["name"] + " Error is " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error with webresource name Error is " + ex.Message);
            }

            Console.WriteLine("********************** WEBRESOURCE MIGRATION IS ENDED **********************");

        }

    }
}
