using Microsoft.Crm.Sdk.Messages;
using Microsoft.Office.Interop.Excel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Threading.Tasks;
using Label = Microsoft.Xrm.Sdk.Label;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.WebServiceClient;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.IO;
using System.IO.Packaging;
using System.IO.Compression;
using System.Xml;
using CrmPluginAttributes;
using System.Web.Services.Description;
using Microsoft.Xrm.Sdk.Client;

namespace CreateFieldsFromExcelToD365
{
    public class Program
    {
        public const string exportFolderSer = "C:\\Users\\35405\\Downloads\\Service\\";
        public const string exportFolderPrd = "C:\\Users\\35405\\Downloads\\Prd\\";
        public const int _languageCode = 1033;
        private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";
        public static string preFix = null;
        public static Guid PublisherSer = new Guid("D21AAB71-79E7-11DD-8874-00188B01E34F");
        public static Guid PublisherPrd = new Guid("D21AAB71-79E7-11DD-8874-00188B01E34F");
        public static EntityReference SystemAdmin = new EntityReference("systemuser", new Guid("1a8fc0e8-7e48-ed11-bba2-6045bdac5a88"));

        static void Main(string[] args)
        {
            preFix = ConfigurationManager.AppSettings["PreFix"].ToString();

            IOrganizationService service = ConnectToCRM(string.Format(connStr, "https://havellscrmqa.crm8.dynamics.com"));
            IOrganizationService serviceSer = ConnectToCRM(string.Format(connStr, "https://havells.crm8.dynamics.com"));


            var query = new QueryExpression("savedquery");

            query.Criteria.AddCondition("savedqueryid", ConditionOperator.Equal, new Guid("73D44A1A-ADD0-4731-AD10-621AADFFB3E3"));

            query.ColumnSet = new ColumnSet(true);
            EntityCollection savedQueries = serviceSer.RetrieveMultiple(query);


            //IOrganizationService serviceSer = ConnectToCRM(string.Format(connStr, "https://havellsdemo.crm8.dynamics.com"));
            //IOrganizationService serviceSer = ConnectToCRM(string.Format(connStr, "https://havellscrmrebuild.crm8.dynamics.com"));


            //RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
            //{
            //    EntityFilters = EntityFilters.All,
            //    LogicalName = "hil_address"
            //};
            //RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
            //EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;

            //var StateEntityReal = StateEntity.OneToManyRelationships;
            //var StateEntityRealN = StateEntity.ManyToOneRelationships;
            //var StateEntityRealNN = StateEntity.ManyToManyRelationships;
            //Console.Clear();

            //foreach(var rel in StateEntityReal)
            //{
            //    Console.WriteLine(rel.ReferencingEntity);
            //}
            //Console.WriteLine("+++++++++++++++++++++++++++++++++++++++++++++");
            //foreach (var rel in StateEntityRealN)
            //{
            //    Console.WriteLine(rel.ReferencedEntity);
            //}


            Console.WriteLine("ddd");

            // RibbonMigration.mainFunction(service, serviceSer);

            //MigrateKeys.mainFunction(service, serviceSer);



            //ReadXML.mainFunction(service, serviceSer);// updatePrimaryField(service, serviceSer);



            //DashboardMigration.mainFunction(service, serviceSer);
            //ViewsModificationNew.mainFunction(service, serviceSer);

            //PluginMigrationService.retivePluginAssemblly(serviceSer, "GenerateOTPWorkFlow");

            //string pluginName = "Autonumber_Setup;DialyHealthIndicator;GenerateOTPWorkFlow;Havells_Plugin;Havells_POStatusTracker;HavellsNewPlugin;HomeAdvisory;IDGenerator;Warranty Management;WorkOrderVirtualEntity;EncryptionWorkflow";
            //string[] entities = pluginName.Split(';');
            //foreach (string ent in entities)
            //{
            //    PluginMigration.retivePluginAssemblly(service, serviceSer, ent);
            //}
            //string EntityLists = "hil_warrantyperiod;hil_warrantyscheme;hil_warrantytemplate;hil_warrantyvoidreason;hil_whatsappproductdivisionconfig;hil_workorderarch;hil_wrongclosurepenalty;hil_yammersettings";

            //string[] entities = EntityLists.Split(';');
            //foreach (string entityName in entities)
            //{
            //    try
            //    {
            //        RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
            //        {
            //            EntityFilters = EntityFilters.Entity,
            //            LogicalName = entityName
            //        };
            //        RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
            //        EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;

            //        RetrieveEntityRequest retrieveEntityRequestDemo = new RetrieveEntityRequest
            //        {
            //            EntityFilters = EntityFilters.Entity,
            //            LogicalName = entityName
            //        };
            //        RetrieveEntityResponse retrieveAccountEntityResponseDemo = (RetrieveEntityResponse)serviceSer.Execute(retrieveEntityRequestDemo);
            //        EntityMetadata StateEntityDemo = retrieveAccountEntityResponseDemo.EntityMetadata;

            //        FormModfication.CreateForm(StateEntity, serviceSer, service, StateEntityDemo.ObjectTypeCode);
            //        Console.WriteLine(StateEntity.DisplayName.LocalizedLabels[0].Label + "|" + StateEntity.LogicalName + "|" + StateEntity.ObjectTypeCode);
            //    }
            //    catch (Exception ex)
            //    {
            //        Console.WriteLine("#N/A" + "|" + "#N/A" + "|" + "#N/A");
            //    }

            //}
            //Console.Clear();
            //string pluginName = "test";
            // "Active AdvisorMaster;Advisor Master Advanced Find View;Advisor Master Associated View;Advisor Master Lookup View;Inactive AdvisorMaster;Quick Find Active AdvisorMaster";


            //string[] entities = pluginName.Split(';');
            //foreach (string ent in entities)
            //{
            //    var query = new QueryExpression("savedquery");

            //    query.Criteria.AddCondition("name", ConditionOperator.Equal, ent);

            //    query.ColumnSet = new ColumnSet(true);
            //    EntityCollection savedQueries = serviceSer.RetrieveMultiple(query);
            //    //RetrieveMultipleRequest retrieveSavedQueriesRequest = new RetrieveMultipleRequest { Query = query };
            //    //RetrieveMultipleResponse retrieveSavedQueriesResponse = (RetrieveMultipleResponse)serviceSer.Execute(retrieveSavedQueriesRequest);
            //    //DataCollection<Entity> savedQueries = retrieveSavedQueriesResponse.EntityCollection.Entities;
            //    ////Display the Retrieved views

            //    Console.WriteLine("totla Views Count " + savedQueries.Entities.Count);
            //    Guid guid = savedQueries[0].Id;
            //    Entity entity = new Entity(savedQueries.EntityName, guid);
            //    entity["isdefault"] = false;// "Channel Partner Advanced Find View";
            //    serviceSer.Update(entity);
            //    Console.WriteLine(ent + "||" + guid);
            //}
            // ReadXML cc = new ReadXML();
            //cc.mainFunction(service, serviceSer);

            //ProcessMigration.MainFunction(service, serviceSer);
            Console.WriteLine("done");
            //CreateEmailTemplate(service, serviceSer);
            //ProcessMigration.craeteworkflow(service, serviceSer);

            //string ProcessNamesList = "AdvisoryLineSync;AMC WC Send From Email to Customer;AMC WC Send From SMS to Customer;Approve Approvals;Approve Manual PO;Approve Partner Application Workflow;Approve Reject Credit Limit;Assign Assignment Matrix;Assign Inventory to Owner Account;Assign position;Assign work order from Power app;Assignment Matrix : Temp;Attachment Type Bank Guarantee;Auto Assign System User Created Jobs;Auto mailers for MG4 creation;Bank Guarantee Return Date Validation;BIGeo Mapping Map Pincode and City entity to city and state;Bill Document uploaded;BOE Acceptance Letter;BOE Receipt Acceptance;Calculate Partner Capacity Workflow;Case ID Generation;Change  Request Updation;Change associated web role based on partner contact role field on contact;Change partner application status when contact expresses interest;Clear For Dispatch;Cleared For Billing;Close Enquiry Line;Communication on Tender Creation;Consumer alert for Job Cancellation due to COVID-19;Consumer App : Customer Validation OTP;Copy Managing Partner from parent account - On change of Contact;Create Job Extension from Job Create;Create PO Status Tracker Line;Create Technical Specification;Creating Unit Warranty Lines if Customer Asset is Approved;Customer Asset populate Model Description;Customer Asset Validation;Customer Communication;Deactivate AMC Staging Doc;Deactivate Bulk Job Uploader Record if status is successfull;Delete Created PMS Rows from PMS Up loader;De-linked PO No. - Populated over Job;Email Generated Service Call;Email to Accounts Team after EMD Aapproved;Escalation SMS;Escallation/Remainder Count on Job;Estimates : On Customer Approval;Excute Assignment Engine for WhatsApp Jobs on change of Product Details;Forgot Password;Havells : Claim Management : Restrict User to add Claim Overheads once Claim approved.;Inactivate Technician Profile;Integration_Delink Remove SBU Lookup from Division;Inventory Request : On Branch Head Approval;Inventory Request : On Submit for Approval;IPGeo Mapping Map Pincode and City entity to city and state;Job : Latitude n Longitude Flow in Contact;Job : OCR;Job Cancellation Request Process;Job Class Reject by Branch Head;Job Closure Time Stamp;Job Estimate : Calculate Charges;Job First Response Time Stamp;JOB PRODUCT : ON UPDATE REPLACED PART;Jobs: Reminder Call SMS;KKG Audit Failed - Notify BSH;LC Document submission intimation on document type BOE;Mail On Design team Change;Mail On Stack Holder Change;Mail to Account after Tender DD approval;Manual Product Request  Approval;MarkCompletedSMS;number;On Create : Aqua Parameters;On Demand PO Push to SAP;On Inspection Report Send Mail to Sales Team;On Inventory Request Create for Manual GRN;On OA Header Creation;On PO Header Create;On Success, Deactivate Customer Asset staging Line;OTPContact_OTP for ForgotPassword;Password Hash and Security Stramp On Contact Create;PhoneCall Creation SMS;Populate Division on Material Group from Material;Populate number NPS NSH number from user;Populate Product Sub Cat from Staging Division Mat group Mapping;PopulateDeactivatedOnForCase;PR Number (Text) to be populated over Job;PR Request Get PO Header Number;Prevent Franchise/DSE to Approve Product Request;Product Registration : Emails and SMS;Product Replacement Approval/Rejection Email to All Levels;Product Request Header - Revenue PO;Product Request Header _ Revenue PO On Division Update;PRODUCT_REQ_HEADER : IV USAGE LOGIC;readiness date update in tender Product;Refresh Job Status;Rerun Job Assignment Engine if is in registered mode;Resend KKG Code to Customer;resend payment link set as No;Reset Claim Process Status on Claim approval;Restrict delete and creation of tender product;Restrict OCL in case of Motor and category is Dealer Stock;Restrict Updation Of Order Check List Product Once OA Created;Resync AMC Call Data with SAP;Re-Sync PR;Resync Product Request;ristrict save after send;Send Adhoc Delivery Intimation To PPC;Send Appointment SMS to Customer;Send Charges SMS;Send Email and SMS on Job Creation;Send Enquiry Response SMS;Send Mail On OA Creation;Send Mail Stake Holder Change In OA;Send Mail Stake Holder Change In OCL;Send mail to Approvar regarding AMC;Send Satisfaction/KKG Code to Customer on Appointment Set;Send SMS on Emergency Call tagged as Yes;Send SMS on Payment Collection;Send SMS to Customer and Technician ;Send SMS to Customer on incorrect msg format;Send SMS/Email to customer on Home advisory enquiry;Service Bom - On Demand For Parts& Populate part and model text;Set Approval Status Cancel;Set Asset type on create;Set Cancelled status on Job on 3rd phone call;Set Category and Sub Category Text on Staging Division Mapping;Set Contact on address if doesn't exist;Set Customer on Customer Asset;Set Customer on Job Incident on create;Set Email Regarding On Job Created from Email;Set Franchisee and DE on owneraccount;Set JobIncident COunt on job;Set KKG Code Required : Yes if Hierarchy Level is Division or Material Group;Set Observation based on cause;Set Product subcategory on change of productcatsubcatmapping;Set Product Type in case of material;Set Regarding when Email Recieved;Set SLA Start Date in Job;Set Status Code;Set Warranty Void On Job Product;SortJobsBasedOnAppointmentTime;status update( mark used);Tasks;Time Off Request : Set owner Account;Update Address in job;Update Alternate Part Description;Update Approver Time Stamp on SAW Activity Approval;Update Asset Purchase Date in Job once Job Work done;Update Asset Sub Category- Serialize;Update Assignment Matrix Upload Status;Update Consumers Alternate Number from Job;Update Customer Asset Status;Update Department  on Tender Product;Update Department on OA Product;Update Department On OrderchecklistProduct;Update Department On Survey;Update Inventory Owner Account;Update Job Closed on when ISOCR = true;Update Job Extension in Job;Update Job Status & Create Child Job - KKG Audit Failed;Update managing partner when created by (portal contact) field is populated.;Update Order Link on Tender;UPDATE PO QTY;Update PO Status Tracker ID in Product Request;Update Product Price;Update Product subcategory & spare Part Family in SBOM;Update related Franchisee on Job (Reallocation);Update Sales Order No. in PO Status tracker;Update Scheme District Exclusion;Update Technical Specification;Update Technician Name on User Card;Update Tolerance and Plant code On OrdercheckList Product;Update Upcountry Flag in Job;User;Warranty Template - On Demand;Work order incident flow product category, sub category,Nature of complaint and observation";
            //string[] entities = ProcessNamesList.Split(';');
            //foreach (string ent in entities)
            //{
            //    ProcessMigration.UpdateWorkflow(service, ent, serviceSer);
            //}





            //string EntityNames = "product;systemuser;task";
            ////"account;appointment;bookableresourcebooking;campaign;characteristic;contact;email;incident;lead;msdyn_customerasset;msdyn_incidenttype;msdyn_incidenttypeproduct;msdyn_resourcerequirement;msdyn_surveyresponse;msdyn_timeoffrequest;msdyn_workorder;msdyn_workorderincident;msdyn_workorderproduct;msdyn_workorderservice;msdyn_workordersubstatus;phonecall;product;systemuser;task;systemuser;task;hil_address;hil_advancefind;hil_advisormaster;hil_advisoryenquiry;hil_alerts;hil_alternatepart;hil_amcdiscountmatrix;hil_amcplan;hil_amcstaging;hil_approval;hil_approvalmatrix;hil_aquaparameters;hil_archivedactivity;hil_archvedjobdatasource;hil_area;hil_assignmentmatrix;hil_assignmentmatrixupload;hil_attachment;hil_attachmentdocumenttype;hil_auditlog;hil_autonumber;hil_bdtatdetails;hil_bdtatownership;hil_bdteam;hil_bdteammember;hil_bizgeomappingstagigng;hil_bomcategory;hil_bomsubcategory;hil_branch;hil_bulkjobsuploader;hil_businessmapping;hil_callsubtype;hil_calltype;hil_campaigndivisions;hil_campaignenquirytypes;hil_campaignwebsitesetup;hil_cancellationreason;hil_casecategory;hil_channelpartnercountryclassification;hil_city;hil_claimallocator;hil_claimcategory;hil_claimheader;hil_claimline;hil_claimlines;hil_claimoverheadline;hil_claimperiod;hil_claimpostingsetup;hil_claimsummary;hil_claimtype;hil_constructiontype;hil_consumerappbridge;hil_consumercategory;hil_consumernpssurvey;hil_consumertype;hil_country;hil_customerassetstagging;hil_customerfeedback;hil_customerwishlist;hil_dealerstockverificationheader;hil_deliveryschedule;hil_deliveryschedulemaster;hil_designteambranchmapping;hil_despatchteammaster;hil_discountmatrix;hil_distributionchannel;hil_district;hil_effeciency;hil_enclosuretype;hil_enquerysegment;hil_enquirybusinessprocessflow;hil_enquirydepartment;hil_enquirydocumenttype;hil_enquirylostreason;hil_enquiryproductsegment;hil_enquirysegmentdcmapping;hil_enquirytype;hil_entityfieldsmetadata;hil_errorcode;hil_escalation;hil_escalationmatrix;hil_estimate;hil_feedback;hil_fixedcompensation;hil_forgotpassword;hil_frequency;hil_grnline;hil_healthindicatorheader;hil_homeadvisoryline;hil_hsncode;hil_icauploader;hil_incentivetable;hil_industrysubtype;hil_industrytype;hil_inspectiontype;hil_installationchecklist;hil_integrationconfiguration;hil_integrationjob;hil_integrationjobrun;hil_integrationjobrundetail;hil_integrationtrace;hil_inventory;hil_inventoryjournal;hil_inventoryrequest;hil_invoice;hil_jobbasecharge;hil_jobcancellationrequest;hil_joberrorcode;hil_jobestimation;hil_jobreassignreason;hil_jobsextension;hil_jobtat;hil_jobtatcategory;hil_labor;hil_leadproduct;hil_materialgroup;hil_materialgroup2;hil_materialgroup3;hil_materialgroup4;hil_materialgroup5;hil_materialreturn;hil_minimumguarantee;hil_minimumstocklevel;hil_mobileappbanner;hil_natureofcomplaint;hil_npssetup;hil_oaheader;hil_oaproduct;hil_oareadinessdatestaging;hil_observation;hil_orderchecklist;hil_orderchecklistproduct;hil_part;hil_partnerdepartment;hil_partnerdivisionmapping;hil_paymentstatus;hil_pincode;hil_plantmaster;hil_plantordertypesetup;hil_pmsconfiguration;hil_pmsconfigurationlines;hil_pmsjobuploader;hil_pmsscheduleconfiguration;hil_politicalmapping;hil_postatustracker;hil_postatustrackerupload;hil_postorder;hil_preferredlanguageforcommunication;hil_productcatalog;hil_production;hil_productrequest;hil_productrequestheader;hil_productrequestheaderarchived;hil_productstaging;hil_productvideos;hil_propertytype;hil_refreshjobs;hil_region;hil_remarks;hil_returnheader;hil_returnline;hil_salesoffice;hil_salesofficebranchheadmapping;hil_salesofficeplantmapping;hil_salestatmaster;hil_sawactivity;hil_sawactivityapproval;hil_sawcategoryapprovals;hil_sbubranchmapping;hil_schemedistrictexclusion;hil_schemeincentive;hil_schemeline;hil_securityroleextension;hil_serialnumber;hil_serviceactionwork;hil_serviceactionworksetup;hil_servicebom;hil_servicecallrequest;hil_servicecallrequestdetail;hil_serviceengineergeocode;hil_smsconfiguration;hil_smstemplates;hil_solarservey;hil_specialincentive;hil_staging;hil_stagingdivisonmaterialgroupmapping;hil_stagingintegrationjson;hil_stagingpricingmapping;hil_stagingsbudivisionmapping;hil_startingmethod;hil_state;hil_statustransitionmatrix;hil_subterritory;hil_tatachievementslabmaster;hil_tatbreachpenalty;hil_tatincentive;hil_technicalspecfication;hil_technician;hil_technicianhealthindicator;hil_tender;hil_tenderattachmentdoctype;hil_tenderattachmentmanager;hil_tenderbankguarantee;hil_tenderbomlineitem;hil_tenderdocs;hil_tenderemailersetup;hil_tenderpaymentdetail;hil_tenderproduct;hil_tolerencesetup;hil_travelclosureincentive;hil_travelexpense;hil_typeofcustomer;hil_typeofproduct;hil_unitwarranty;hil_upcountrydataupdate;hil_upcountrytravelcharge;hil_usageindicator;hil_userbranchmapping;hil_usersecurityroleextension;hil_usersimeibinding;hil_voltage;hil_warrantyperiod;hil_warrantyscheme;hil_warrantytemplate;hil_warrantyvoidreason;hil_whatsappproductdivisionconfig;hil_workorderarch;hil_wrongclosurepenalty;hil_yammersettings;dyn_advancefindconfig;ispl_autonumberconfiguration;new_integrationstaging;plt_idgenerator";
            //string[] entities = EntityNames.Split(';');
            //Console.WriteLine("totla Entity Count " + entities.Length);
            //int totalCount = entities.Length;
            //int done = 1;
            //int error = 0;
            //foreach (string ent in entities)
            //{
            //    MigrateRibbon(service, serviceSer, ent);
            //}

        }
        #region old Mail Function
        //static void Main(string[] args)
        //{
        //    preFix = ConfigurationManager.AppSettings["PreFix"].ToString();

        //    IOrganizationService service = ConnectToCRM(string.Format(connStr, "https://havellscrmqa.crm8.dynamics.com"));
        //    //IOrganizationService serviceSer = ConnectToCRM(string.Format(connStr, "https://havellsservice.crm8.dynamics.com"));
        //    IOrganizationService serviceSer = ConnectToCRM(string.Format(connStr, "https://havellsdemo.crm8.dynamics.com"));



        //    //CreateGlobalOptionSet(service, serviceSer);
        //    //EntityMigration(service, serviceSer);
        //    //  UpdateEntityMetaData(service, serviceSer);
        //    //UpdateAttribute(service, serviceSer, "contact");

        //    //string EntityNames = "account;appointment;bookableresourcebooking;campaign;characteristic;contact;email;incident;lead;msdyn_customerasset;msdyn_incidenttype;msdyn_incidenttypeproduct;msdyn_resourcerequirement;msdyn_surveyresponse;msdyn_timeoffrequest;msdyn_workorder;msdyn_workorderincident;msdyn_workorderproduct;msdyn_workorderservice;msdyn_workordersubstatus;phonecall;product;systemuser;task;systemuser;task;hil_address;hil_advancefind;hil_advisormaster;hil_advisoryenquiry;hil_alerts;hil_alternatepart;hil_amcdiscountmatrix;hil_amcplan;hil_amcstaging;hil_approval;hil_approvalmatrix;hil_aquaparameters;hil_archivedactivity;hil_archvedjobdatasource;hil_area;hil_assignmentmatrix;hil_assignmentmatrixupload;hil_attachment;hil_attachmentdocumenttype;hil_auditlog;hil_autonumber;hil_bdtatdetails;hil_bdtatownership;hil_bdteam;hil_bdteammember;hil_bizgeomappingstagigng;hil_bomcategory;hil_bomsubcategory;hil_branch;hil_bulkjobsuploader;hil_businessmapping;hil_callsubtype;hil_calltype;hil_campaigndivisions;hil_campaignenquirytypes;hil_campaignwebsitesetup;hil_cancellationreason;hil_casecategory;hil_channelpartnercountryclassification;hil_city;hil_claimallocator;hil_claimcategory;hil_claimheader;hil_claimline;hil_claimlines;hil_claimoverheadline;hil_claimperiod;hil_claimpostingsetup;hil_claimsummary;hil_claimtype;hil_constructiontype;hil_consumerappbridge;hil_consumercategory;hil_consumernpssurvey;hil_consumertype;hil_country;hil_customerassetstagging;hil_customerfeedback;hil_customerwishlist;hil_dealerstockverificationheader;hil_deliveryschedule;hil_deliveryschedulemaster;hil_designteambranchmapping;hil_despatchteammaster;hil_discountmatrix;hil_distributionchannel;hil_district;hil_effeciency;hil_enclosuretype;hil_enquerysegment;hil_enquirybusinessprocessflow;hil_enquirydepartment;hil_enquirydocumenttype;hil_enquirylostreason;hil_enquiryproductsegment;hil_enquirysegmentdcmapping;hil_enquirytype;hil_entityfieldsmetadata;hil_errorcode;hil_escalation;hil_escalationmatrix;hil_estimate;hil_feedback;hil_fixedcompensation;hil_forgotpassword;hil_frequency;hil_grnline;hil_healthindicatorheader;hil_homeadvisoryline;hil_hsncode;hil_icauploader;hil_incentivetable;hil_industrysubtype;hil_industrytype;hil_inspectiontype;hil_installationchecklist;hil_integrationconfiguration;hil_integrationjob;hil_integrationjobrun;hil_integrationjobrundetail;hil_integrationtrace;hil_inventory;hil_inventoryjournal;hil_inventoryrequest;hil_invoice;hil_jobbasecharge;hil_jobcancellationrequest;hil_joberrorcode;hil_jobestimation;hil_jobreassignreason;hil_jobsextension;hil_jobtat;hil_jobtatcategory;hil_labor;hil_leadproduct;hil_materialgroup;hil_materialgroup2;hil_materialgroup3;hil_materialgroup4;hil_materialgroup5;hil_materialreturn;hil_minimumguarantee;hil_minimumstocklevel;hil_mobileappbanner;hil_natureofcomplaint;hil_npssetup;hil_oaheader;hil_oaproduct;hil_oareadinessdatestaging;hil_observation;hil_orderchecklist;hil_orderchecklistproduct;hil_part;hil_partnerdepartment;hil_partnerdivisionmapping;hil_paymentstatus;hil_pincode;hil_plantmaster;hil_plantordertypesetup;hil_pmsconfiguration;hil_pmsconfigurationlines;hil_pmsjobuploader;hil_pmsscheduleconfiguration;hil_politicalmapping;hil_postatustracker;hil_postatustrackerupload;hil_postorder;hil_preferredlanguageforcommunication;hil_productcatalog;hil_production;hil_productrequest;hil_productrequestheader;hil_productrequestheaderarchived;hil_productstaging;hil_productvideos;hil_propertytype;hil_refreshjobs;hil_region;hil_remarks;hil_returnheader;hil_returnline;hil_salesoffice;hil_salesofficebranchheadmapping;hil_salesofficeplantmapping;hil_salestatmaster;hil_sawactivity;hil_sawactivityapproval;hil_sawcategoryapprovals;hil_sbubranchmapping;hil_schemedistrictexclusion;hil_schemeincentive;hil_schemeline;hil_securityroleextension;hil_serialnumber;hil_serviceactionwork;hil_serviceactionworksetup;hil_servicebom;hil_servicecallrequest;hil_servicecallrequestdetail;hil_serviceengineergeocode;hil_smsconfiguration;hil_smstemplates;hil_solarservey;hil_specialincentive;hil_staging;hil_stagingdivisonmaterialgroupmapping;hil_stagingintegrationjson;hil_stagingpricingmapping;hil_stagingsbudivisionmapping;hil_startingmethod;hil_state;hil_statustransitionmatrix;hil_subterritory;hil_tatachievementslabmaster;hil_tatbreachpenalty;hil_tatincentive;hil_technicalspecfication;hil_technician;hil_technicianhealthindicator;hil_tender;hil_tenderattachmentdoctype;hil_tenderattachmentmanager;hil_tenderbankguarantee;hil_tenderbomlineitem;hil_tenderdocs;hil_tenderemailersetup;hil_tenderpaymentdetail;hil_tenderproduct;hil_tolerencesetup;hil_travelclosureincentive;hil_travelexpense;hil_typeofcustomer;hil_typeofproduct;hil_unitwarranty;hil_upcountrydataupdate;hil_upcountrytravelcharge;hil_usageindicator;hil_userbranchmapping;hil_usersecurityroleextension;hil_usersimeibinding;hil_voltage;hil_warrantyperiod;hil_warrantyscheme;hil_warrantytemplate;hil_warrantyvoidreason;hil_whatsappproductdivisionconfig;hil_workorderarch;hil_wrongclosurepenalty;hil_yammersettings;dyn_advancefindconfig;ispl_autonumberconfiguration;new_integrationstaging;plt_idgenerator";
        //    //string[] entities = EntityNames.Split(';');
        //    //Console.WriteLine("totla Entity Count " + entities.Length);
        //    //int totalCount = entities.Length;
        //    //int done = 1;
        //    //int error = 0;
        //    //Console.WriteLine("Entity Creation Started");
        //    #region CreateSystemView

        //    //CreateWebResource(service, serviceSer);

        //    Console.WriteLine("View Creation Started");
        //    string EntityNames = "systemuser".ToLower();// "hil_address;hil_advancefind;hil_advisormaster;hil_advisoryenquiry;hil_alerts;hil_alternatepart;hil_amcdiscountmatrix;hil_amcplan;hil_amcstaging;hil_approval;hil_approvalmatrix;hil_aquaparameters;hil_archivedactivity;hil_archvedjobdatasource;hil_area;hil_assignmentmatrix;hil_assignmentmatrixupload;hil_attachment;hil_attachmentdocumenttype;hil_auditlog;hil_autonumber;hil_bdtatdetails;hil_bdtatownership;hil_bdteam;hil_bdteammember;hil_bizgeomappingstagigng;hil_bomcategory;hil_bomsubcategory;hil_branch;hil_bulkjobsuploader;hil_businessmapping;hil_callsubtype;hil_calltype;hil_campaigndivisions;hil_campaignenquirytypes;hil_campaignwebsitesetup;hil_cancellationreason;hil_casecategory;hil_channelpartnercountryclassification;hil_city;hil_claimallocator;hil_claimcategory;hil_claimheader;hil_claimline;hil_claimlines;hil_claimoverheadline;hil_claimperiod;hil_claimpostingsetup;hil_claimsummary;hil_claimtype;hil_constructiontype;hil_consumerappbridge;hil_consumercategory;hil_consumernpssurvey;hil_consumertype;hil_country;hil_customerassetstagging;hil_customerfeedback;hil_customerwishlist;hil_dealerstockverificationheader;hil_deliveryschedule;hil_deliveryschedulemaster;hil_designteambranchmapping;hil_despatchteammaster;hil_discountmatrix;hil_distributionchannel;hil_district;hil_effeciency;hil_enclosuretype;hil_enquerysegment;hil_enquirybusinessprocessflow;hil_enquirydepartment;hil_enquirydocumenttype;hil_enquirylostreason;hil_enquiryproductsegment;hil_enquirysegmentdcmapping;hil_enquirytype;hil_entityfieldsmetadata;hil_errorcode;hil_escalation;hil_escalationmatrix;hil_estimate;hil_feedback;hil_fixedcompensation;hil_forgotpassword;hil_frequency;hil_grnline;hil_healthindicatorheader;hil_homeadvisoryline;hil_hsncode;hil_icauploader;hil_incentivetable;hil_industrysubtype;hil_industrytype;hil_inspectiontype;hil_installationchecklist;hil_integrationconfiguration;hil_integrationjob;hil_integrationjobrun;hil_integrationjobrundetail;hil_integrationtrace;hil_inventory;hil_inventoryjournal;hil_inventoryrequest;hil_invoice;hil_jobbasecharge;hil_jobcancellationrequest;hil_joberrorcode;hil_jobestimation;hil_jobreassignreason;hil_jobsextension;hil_jobtat;hil_jobtatcategory;hil_labor;hil_leadproduct;hil_materialgroup;hil_materialgroup2;hil_materialgroup3;hil_materialgroup4;hil_materialgroup5;hil_materialreturn;hil_minimumguarantee;hil_minimumstocklevel;hil_mobileappbanner;hil_natureofcomplaint;hil_npssetup;hil_oaheader;hil_oaproduct;hil_oareadinessdatestaging;hil_observation;hil_orderchecklist;hil_orderchecklistproduct;hil_part;hil_partnerdepartment;hil_partnerdivisionmapping;hil_paymentstatus;hil_pincode;hil_plantmaster;hil_plantordertypesetup;hil_pmsconfiguration;hil_pmsconfigurationlines;hil_pmsjobuploader;hil_pmsscheduleconfiguration;hil_politicalmapping;hil_postatustracker;hil_postatustrackerupload;hil_postorder;hil_preferredlanguageforcommunication;hil_productcatalog;hil_production;hil_productrequest;hil_productrequestheader;hil_productrequestheaderarchived;hil_productstaging;hil_productvideos;hil_propertytype;hil_refreshjobs;hil_region;hil_remarks;hil_returnheader;hil_returnline;hil_salesoffice;hil_salesofficebranchheadmapping;hil_salesofficeplantmapping;hil_salestatmaster;hil_sawactivity;hil_sawactivityapproval;hil_sawcategoryapprovals;hil_sbubranchmapping;hil_schemedistrictexclusion;hil_schemeincentive;hil_schemeline;hil_securityroleextension;hil_serialnumber;hil_serviceactionwork;hil_serviceactionworksetup;hil_servicebom;hil_servicecallrequest;hil_servicecallrequestdetail;hil_serviceengineergeocode;hil_smsconfiguration;hil_smstemplates;hil_solarservey;hil_specialincentive;hil_staging;hil_stagingdivisonmaterialgroupmapping;hil_stagingintegrationjson;hil_stagingpricingmapping;hil_stagingsbudivisionmapping;hil_startingmethod;hil_state;hil_statustransitionmatrix;hil_subterritory;hil_tatachievementslabmaster;hil_tatbreachpenalty;hil_tatincentive;hil_technicalspecfication;hil_technician;hil_technicianhealthindicator;hil_tender;hil_tenderattachmentdoctype;hil_tenderattachmentmanager;hil_tenderbankguarantee;hil_tenderbomlineitem;hil_tenderdocs;hil_tenderemailersetup;hil_tenderpaymentdetail;hil_tenderproduct;hil_tolerencesetup;hil_travelclosureincentive;hil_travelexpense;hil_typeofcustomer;hil_typeofproduct;hil_unitwarranty;hil_upcountrydataupdate;hil_upcountrytravelcharge;hil_usageindicator;hil_userbranchmapping;hil_usersecurityroleextension;hil_usersimeibinding;hil_voltage;hil_warrantyperiod;hil_warrantyscheme;hil_warrantytemplate;hil_warrantyvoidreason;hil_whatsappproductdivisionconfig;hil_workorderarch;hil_wrongclosurepenalty;hil_yammersettings;dyn_advancefindconfig;ispl_autonumberconfiguration;new_integrationstaging;plt_idgenerator;account;appointment;bookableresourcebooking;campaign;characteristic;contact;email;incident;lead;msdyn_customerasset;msdyn_incidenttype;msdyn_incidenttypeproduct;msdyn_resourcerequirement;msdyn_surveyresponse;msdyn_timeoffrequest;msdyn_workorder;msdyn_workorderincident;msdyn_workorderproduct;msdyn_workorderservice;msdyn_workordersubstatus;phonecall;product;systemuser;task";
        //    string[] entities = EntityNames.Split(';');
        //    Console.WriteLine("totla Entity Count " + entities.Length);
        //    int totalCount = entities.Length;
        //    int done = 1;
        //    int error = 0;
        //    //Console.WriteLine("Entity Creation Started");

        //    #region CreateForm
        //    Console.WriteLine("Forms Creation Started");
        //    foreach (string ent in entities)
        //    {
        //        RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
        //        {
        //            EntityFilters = EntityFilters.Entity,
        //            LogicalName = ent
        //        };
        //        RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
        //        EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;

        //        RetrieveEntityRequest retrieveEntityRequestDemo = new RetrieveEntityRequest
        //        {
        //            EntityFilters = EntityFilters.Entity,
        //            LogicalName = ent
        //        };
        //        RetrieveEntityResponse retrieveAccountEntityResponseDemo = (RetrieveEntityResponse)serviceSer.Execute(retrieveEntityRequestDemo);
        //        EntityMetadata StateEntityDemo = retrieveAccountEntityResponseDemo.EntityMetadata;

        //        CreateForm(StateEntity, serviceSer, service, StateEntityDemo.ObjectTypeCode);
        //    }
        //    #endregion
        //    #endregion
        //    Console.WriteLine("View Creation Ended");



        //    //string ent = "HIL_INTEGRATIONTRACE".ToLower();
        //    //try
        //    //{
        //    //    RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
        //    //    {
        //    //        EntityFilters = EntityFilters.All,
        //    //        LogicalName = ent
        //    //    };
        //    //    RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
        //    //    EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;
        //    //    createNto1RelationShip(StateEntity, serviceSer); 
        //    //    //CreateFieldsExceptLookUp(StateEntity, serviceSer);
        //    //    Console.WriteLine("Entity Created " + " / ");

        //    //}
        //    //catch (Exception ex)
        //    //{
        //    //    Console.WriteLine("-------------------------Failuer in entity MetaData creation with name " + ent + " Error is:-  " + ex.Message);
        //    //}


        //    //var query1 = new QueryExpression("webresource");
        //    //query1.Criteria.AddCondition("name", ConditionOperator.In, "hil_lead");
        //    //query1.ColumnSet = new ColumnSet(true);
        //    //var results1 = service.RetrieveMultiple(query1);
        //    //if (results1.Entities.Count != 0)
        //    //{
        //    //    serviceSer.Create(results1[0]);
        //    //}

        //    //string EntityNames = "hil_address;hil_advancefind;hil_advisormaster;hil_advisoryenquiry;hil_alerts;hil_alternatepart;hil_amcdiscountmatrix;hil_amcplan;hil_amcstaging;hil_approval;hil_approvalmatrix;hil_aquaparameters;hil_area;hil_assignmentmatrix;hil_assignmentmatrixupload;hil_attachment;hil_attachmentdocumenttype;hil_auditlog;hil_autonumber;hil_bdtatdetails;hil_bdtatownership;hil_bdteam;hil_bdteammember;hil_bizgeomappingstagigng;hil_bomcategory;hil_bomsubcategory;hil_branch;hil_bulkjobsuploader;hil_businessmapping;hil_callsubtype;hil_calltype;hil_campaigndivisions;hil_campaignenquirytypes;hil_campaignwebsitesetup;hil_cancellationreason;hil_casecategory;hil_channelpartnercountryclassification;hil_city;hil_claimallocator;hil_claimcategory;hil_claimheader;hil_claimline;hil_claimlines;hil_claimoverheadline;hil_claimperiod;hil_claimpostingsetup;hil_claimsummary;hil_claimtype;hil_constructiontype;hil_consumerappbridge;hil_consumercategory;hil_consumernpssurvey;hil_consumertype;hil_country;hil_customerassetstagging;hil_customerfeedback;hil_customerwishlist;hil_dealerstockverificationheader;hil_deliveryschedule;hil_deliveryschedulemaster;hil_designteambranchmapping;hil_despatchteammaster;hil_discountmatrix;hil_distributionchannel;hil_district;hil_effeciency;hil_enclosuretype;hil_enquerysegment;hil_enquirybusinessprocessflow;hil_enquirydepartment;hil_enquirydocumenttype;hil_enquirylostreason;hil_enquiryproductsegment;hil_enquirysegmentdcmapping;hil_enquirytype;hil_entityfieldsmetadata;hil_errorcode;hil_escalation;hil_escalationmatrix;hil_estimate;hil_feedback;hil_fixedcompensation;hil_forgotpassword;hil_frequency;hil_grnline;hil_healthindicatorheader;hil_homeadvisoryline;hil_hsncode;hil_icauploader;hil_incentivetable;hil_industrysubtype;hil_industrytype;hil_inspectiontype;hil_installationchecklist;hil_integrationconfiguration;hil_integrationjob;hil_integrationjobrun;hil_integrationjobrundetail;hil_integrationtrace;hil_inventory;hil_inventoryjournal;hil_inventoryrequest;hil_invoice;hil_jobbasecharge;hil_jobcancellationrequest;hil_joberrorcode;hil_jobestimation;hil_jobreassignreason;hil_jobsextension;hil_jobtat;hil_jobtatcategory;hil_labor;hil_leadproduct;hil_materialgroup;hil_materialgroup2;hil_materialgroup3;hil_materialgroup4;hil_materialgroup5;hil_materialreturn;hil_minimumguarantee;hil_minimumstocklevel;hil_mobileappbanner;hil_natureofcomplaint;hil_npssetup;hil_oaheader;hil_oaproduct;hil_oareadinessdatestaging;hil_observation;hil_orderchecklist;hil_orderchecklistproduct;hil_part;hil_partnerdepartment;hil_partnerdivisionmapping;hil_paymentstatus;hil_pincode;hil_plantmaster;hil_plantordertypesetup;hil_pmsconfiguration;hil_pmsconfigurationlines;hil_pmsjobuploader;hil_pmsscheduleconfiguration;hil_politicalmapping;hil_postatustracker;hil_postatustrackerupload;hil_postorder;hil_preferredlanguageforcommunication;hil_productcatalog;hil_production;hil_productrequest;hil_productrequestheader;hil_productstaging;hil_productvideos;hil_propertytype;hil_refreshjobs;hil_region;hil_remarks;hil_returnheader;hil_returnline;hil_salesoffice;hil_salesofficebranchheadmapping;hil_salesofficeplantmapping;hil_salestatmaster;hil_sawactivity;hil_sawactivityapproval;hil_sawcategoryapprovals;hil_sbubranchmapping;hil_schemedistrictexclusion;hil_schemeincentive;hil_schemeline;hil_securityroleextension;hil_serialnumber;hil_serviceactionwork;hil_serviceactionworksetup;hil_servicebom;hil_servicecallrequest;hil_servicecallrequestdetail;hil_serviceengineergeocode;hil_smsconfiguration;hil_smstemplates;hil_solarservey;hil_specialincentive;hil_staging;hil_stagingdivisonmaterialgroupmapping;hil_stagingintegrationjson;hil_stagingpricingmapping;hil_stagingsbudivisionmapping;hil_startingmethod;hil_state;hil_statustransitionmatrix;hil_subterritory;hil_tatachievementslabmaster;hil_tatbreachpenalty;hil_tatincentive;hil_technicalspecfication;hil_technician;hil_technicianhealthindicator;hil_tender;hil_tenderattachmentdoctype;hil_tenderattachmentmanager;hil_tenderbankguarantee;hil_tenderbomlineitem;hil_tenderdocs;hil_tenderemailersetup;hil_tenderpaymentdetail;hil_tenderproduct;hil_tolerencesetup;hil_travelclosureincentive;hil_travelexpense;hil_typeofcustomer;hil_typeofproduct;hil_unitwarranty;hil_upcountrydataupdate;hil_upcountrytravelcharge;hil_usageindicator;hil_userbranchmapping;hil_usersecurityroleextension;hil_usersimeibinding;hil_voltage;hil_warrantyperiod;hil_warrantyscheme;hil_warrantytemplate;hil_warrantyvoidreason;hil_whatsappproductdivisionconfig;hil_workorderarch;hil_wrongclosurepenalty;hil_yammersettings;task;systemuser;product;plt_idgenerator;phonecall;new_integrationstaging;msdyn_workordersubstatus;msdyn_workorderservice;msdyn_workorderproduct;msdyn_workorderincident;msdyn_workorder;msdyn_timeoffrequest;msdyn_surveyresponse;msdyn_resourcerequirement;msdyn_incidenttypeproduct;msdyn_incidenttype;msdyn_customerasset;lead;incident;email;contact;characteristic;campaign;bookableresourcebooking;appointment;account";
        //    //string[] entities = EntityNames.Split(';');
        //    //Console.WriteLine("totla Entity Count " + entities.Length);
        //    //foreach (string entityName1 in entities)
        //    //{
        //    //    RetrieveEntityRequest retrieveEntityRequest12 = new RetrieveEntityRequest
        //    //    {
        //    //        EntityFilters = EntityFilters.All,
        //    //        LogicalName = entityName1
        //    //    };
        //    //    RetrieveEntityResponse retrieveAccountEntityResponse12 = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest12);
        //    //    EntityMetadata StateEntity12 = retrieveAccountEntityResponse12.EntityMetadata;

        //    //    Console.WriteLine(StateEntity12.LogicalName + "||" + StateEntity12.DisplayName.LocalizedLabels[0].Label);
        //    //}


        //    //RetriveViewDepndency(serviceSer, "hil_city");



        //    //string entityName = "account".ToLower();

        //    //createEntity(service, serviceSer, entityName);

        //    //RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
        //    //{
        //    //    EntityFilters = EntityFilters.All,
        //    //    LogicalName = entityName
        //    //};
        //    //RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
        //    //EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;


        //    //RetrieveEntityRequest retrieveEntityRequestDemo = new RetrieveEntityRequest
        //    //{
        //    //    EntityFilters = EntityFilters.Entity,
        //    //    LogicalName = entityName
        //    //};
        //    //RetrieveEntityResponse retrieveAccountEntityResponseDemo = (RetrieveEntityResponse)serviceSer.Execute(retrieveEntityRequestDemo);
        //    //EntityMetadata StateEntityDemo = retrieveAccountEntityResponseDemo.EntityMetadata;
        //    //CreateSystemView(service, serviceSer, StateEntity.ObjectTypeCode);
        //    //CreateUserView(service, serviceSer, StateEntity.ObjectTypeCode);
        //    //createNto1RelationShip(StateEntity, serviceSer);
        //    //CreateSystemView(service, serviceSer, StateEntity.ObjectTypeCode);

        //    //UpdateExistingSystemView(service, serviceSer, StateEntity.ObjectTypeCode, StateEntityDemo.ObjectTypeCode);
        //    //CreateForm(StateEntity, serviceSer, service, StateEntityDemo.ObjectTypeCode);

        //    //Console.WriteLine("ddd");

        //    // RetriveSitemap(service, serviceSer);

        //    //string entityName = "hil_tender";
        //    //craeteworkflow(service, serviceSer);
        //    //MigrateRibbon(service, serviceSer, entityName);
        //    //CreateEmailTemplate(service, serviceSer);
        //    //UpdateGlobalOptionSet(service, serviceSer);
        //    //// CreateWebResource(service, serviceSer);
        //    //CreateGlobalOptionSet(service, serviceSer);

        //    //UpdateGlobalOptionSet(service, serviceSer);
        //    //UpdateEntityMetaData(service, serviceSer);
        //}
        #endregion
        static void updatePrimaryField(IOrganizationService service, IOrganizationService serviceSer)
        {
            string EntityNames = "hil_address;hil_advancefind;dyn_advancefindconfig;hil_advisormaster;hil_advisoryenquiry;hil_homeadvisoryline;hil_alerts;hil_alternatepart;hil_amcdiscountmatrix;hil_amcplan;hil_amcstaging;hil_approvalmatrix;hil_approval;hil_aquaparameters;hil_area;hil_assignmentmatrix;hil_assignmentmatrixupload;hil_attachment;hil_tenderattachmentdoctype;hil_attachmentdocumenttype;hil_tenderattachmentmanager;hil_auditlog;ispl_autonumberconfiguration;hil_autonumber;hil_tenderbankguarantee;hil_jobbasecharge;hil_bdtatdetails;hil_bdtatownership;hil_bdteam;hil_bdteammember;hil_bizgeomappingstagigng;hil_bomcategory;hil_bomsubcategory;hil_branch;hil_bulkjobsuploader;hil_businessmapping;hil_callsubtype;hil_calltype;hil_campaignwebsitesetup;hil_campaigndivisions;hil_campaignenquirytypes;hil_cancellationreason;hil_casecategory;hil_channelpartnercountryclassification;hil_city;hil_claimallocator;hil_claimcategory;hil_claimline;hil_claimlines;hil_claimoverheadline;hil_claimperiod;hil_claimpostingsetup;hil_claimsummary;hil_claimtype;hil_constructiontype;hil_consumerappbridge;hil_consumercategory;hil_consumernpssurvey;hil_consumertype;hil_country;hil_customerassetstagging;hil_remarks;hil_customerfeedback;hil_customerwishlist;hil_healthindicatorheader;hil_technicianhealthindicator;hil_dealerstockverificationheader;hil_materialreturn;hil_deliveryschedule;hil_deliveryschedulemaster;hil_despatchteammaster;hil_discountmatrix;hil_distributionchannel;hil_district;hil_effeciency;hil_enclosuretype;hil_enquirybusinessprocessflow;hil_enquirydepartment;hil_enquirydocumenttype;hil_enquirylostreason;hil_enquiryproductsegment;hil_enquerysegment;hil_enquirysegmentdcmapping;hil_enquirytype;hil_entityfieldsmetadata;hil_errorcode;hil_escalation;hil_escalationmatrix;hil_estimate;hil_feedback;hil_fixedcompensation;hil_forgotpassword;hil_frequency;hil_grnline;hil_hsncode;hil_icauploader;plt_idgenerator;hil_incentivetable;hil_industrysubtype;hil_industrytype;hil_inspectiontype;hil_installationchecklist;hil_integrationconfiguration;hil_integrationjob;hil_integrationjobrun;hil_integrationjobrundetail;new_integrationstaging;hil_staging;hil_integrationtrace;hil_inventory;hil_inventoryjournal;hil_inventoryrequest;hil_invoice;hil_jobcancellationrequest;hil_joberrorcode;hil_jobestimation;hil_jobreassignreason;hil_jobtat;hil_jobtatcategory;hil_jobsextension;hil_travelclosureincentive;hil_labor;hil_leadproduct;hil_materialgroup;hil_materialgroup2;hil_materialgroup3;hil_materialgroup4;hil_materialgroup5;hil_minimumguarantee;hil_minimumstocklevel;hil_mobileappbanner;hil_natureofcomplaint;hil_npssetup;hil_oaheader;hil_oaproduct;hil_oareadinessdatestaging;hil_observation;hil_orderchecklist;hil_orderchecklistproduct;hil_part;hil_partnerdepartment;hil_partnerdivisionmapping;hil_paymentstatus;hil_claimheader;hil_pincode;hil_plantmaster;hil_plantordertypesetup;hil_pmsconfiguration;hil_pmsconfigurationlines;hil_pmsjobuploader;hil_pmsscheduleconfiguration;hil_postatustracker;hil_postatustrackerupload;hil_politicalmapping;hil_postorder;hil_preferredlanguageforcommunication;hil_productcatalog;hil_productrequest;hil_productrequestheader;hil_productstaging;hil_stagingdivisonmaterialgroupmapping;hil_productvideos;hil_production;hil_propertytype;hil_refreshjobs;hil_region;hil_returnheader;hil_returnline;hil_salesoffice;hil_salesofficebranchheadmapping;hil_salesofficeplantmapping;hil_salestatmaster;hil_sawactivity;hil_sawactivityapproval;hil_serviceactionwork;hil_sawcategoryapprovals;hil_serviceactionworksetup;hil_sbubranchmapping;hil_schemedistrictexclusion;hil_schemeincentive;hil_schemeline;hil_securityroleextension;hil_serialnumber;hil_servicebom;hil_servicecallrequest;hil_servicecallrequestdetail;hil_serviceengineergeocode;hil_smsconfiguration;hil_smstemplates;hil_solarservey;hil_specialincentive;hil_stagingintegrationjson;hil_stagingpricingmapping;hil_stagingsbudivisionmapping;hil_startingmethod;hil_state;hil_statustransitionmatrix;hil_subterritory;hil_tatachievementslabmaster;hil_tatbreachpenalty;hil_tatincentive;hil_technicalspecfication;hil_technician;hil_tender;hil_tenderbomlineitem;hil_designteambranchmapping;hil_tenderdocs;hil_tenderemailersetup;hil_tenderpaymentdetail;hil_tenderproduct;hil_userbranchmapping;hil_tolerencesetup;hil_travelexpense;hil_typeofcustomer;hil_typeofproduct;hil_unitwarranty;hil_upcountrydataupdate;hil_upcountrytravelcharge;hil_usageindicator;hil_usersecurityroleextension;hil_usersimeibinding;hil_voltage;hil_warrantyperiod;hil_warrantyscheme;hil_warrantytemplate;hil_warrantyvoidreason;hil_whatsappproductdivisionconfig;hil_workorderarch;hil_wrongclosurepenalty;hil_yammersettings;account;appointment;bookableresourcebooking;campaign;characteristic;contact;email;incident;lead;msdyn_customerasset;msdyn_incidenttype;msdyn_incidenttypeproduct;msdyn_resourcerequirement;msdyn_timeoffrequest;msdyn_workorder;msdyn_workorderincident;msdyn_workorderproduct;msdyn_workorderservice;msdyn_workordersubstatus;phonecall;product;systemuser;task";
            string[] entities = EntityNames.Split(';');
            Console.WriteLine("totla Entity Count " + entities.Length);
            int totalCount = entities.Length;
            int done = 1;
            int error = 0;

            Console.WriteLine("Fields Creation Started");
            foreach (string ent in entities)
            {
                try
                {
                    RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
                    {
                        EntityFilters = EntityFilters.All,
                        LogicalName = ent
                    };
                    RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
                    EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;

                    RetrieveEntityResponse retrieveAccountEntityResponseSer = (RetrieveEntityResponse)serviceSer.Execute(retrieveEntityRequest);
                    EntityMetadata StateEntitySer = retrieveAccountEntityResponseSer.EntityMetadata;
                    updatefield(StateEntity, StateEntitySer, service, serviceSer);
                    Console.WriteLine(done + " / " + totalCount);
                    done++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("-------------------------Failuer in entity MetaData creation with name " + ent + " Error is:-  " + ex.Message);
                }
            }
        }
        static void updatefield(EntityMetadata prdEntity, EntityMetadata serviceEntity, IOrganizationService service, IOrganizationService serviceSer)
        {
            try
            {

                RetrieveAttributeRequest _PrdAttributeRequest = new RetrieveAttributeRequest
                {
                    EntityLogicalName = prdEntity.LogicalName,
                    LogicalName = prdEntity.PrimaryNameAttribute,
                    RetrieveAsIfPublished = true
                };
                // Execute the request
                RetrieveAttributeResponse _prdAttributeResponse = (RetrieveAttributeResponse)service.Execute(_PrdAttributeRequest);
                StringAttributeMetadata _prdPrimaryField = (StringAttributeMetadata)_prdAttributeResponse.AttributeMetadata;
                RetrieveAttributeRequest _ServiceAttributeRequest = new RetrieveAttributeRequest
                {
                    EntityLogicalName = prdEntity.LogicalName,
                    LogicalName = prdEntity.PrimaryNameAttribute,
                    RetrieveAsIfPublished = true
                };
                // Execute the request
                RetrieveAttributeResponse _ServiceAttributeResponse = (RetrieveAttributeResponse)serviceSer.Execute(_ServiceAttributeRequest);
                StringAttributeMetadata _servicePrimaryField = (StringAttributeMetadata)_ServiceAttributeResponse.AttributeMetadata;
                _servicePrimaryField.MaxLength = _prdPrimaryField.MaxLength;
                _servicePrimaryField.RequiredLevel = _prdPrimaryField.RequiredLevel;
                UpdateAttributeRequest updateRequest = new UpdateAttributeRequest
                {
                    Attribute = (AttributeMetadata)_servicePrimaryField,
                    EntityName = prdEntity.LogicalName,
                };
                serviceSer.Execute(updateRequest);
                Console.WriteLine("Primary Field Updated of entity " + prdEntity.LogicalName);

            }
            catch (Exception ex)
            {
                Console.WriteLine("@@@@@@@@@@@@@@ Entity " + prdEntity.LogicalName + " || Error " + ex.Message);
            }
        }
        public static void CreateEmailTemplate1(IOrganizationService service, IOrganizationService serviceSer)
        {
            QueryExpression query = new QueryExpression("template");
            //query.Criteria.AddCondition("parentbusinessunitid", ConditionOperator.Null);
            query.ColumnSet = new ColumnSet(true);
            query.NoLock = true;
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;
            EntityCollection entityCollection = service.RetrieveMultiple(query);
            do
            {
                foreach (Entity entity1 in entityCollection.Entities)
                {
                    try
                    {
                        if (entity1.GetAttributeValue<EntityReference>("businessunitid").Id == new Guid("574D775D-60EB-E811-A96C-000D3AF05828"))
                        {
                            entity1["businessunitid"] = new EntityReference("businessunit", new Guid("3D41A7E9-D037-ED11-9DB1-000D3AF071C9"));
                        }
                        serviceSer.Create(entity1);
                        Console.WriteLine("Done");
                    }
                    catch (Exception ex)
                    {
                        if (!ex.Message.Contains(" already exists on business unit"))
                            if (!ex.Message.Contains("Cannot insert duplicate key exception when executing non-query"))
                                Console.WriteLine("Error " + ex.Message);
                    }
                }
                //Console.WriteLine("createMordernDerivenApp with name " + entity.GetAttributeValue<string>("name") + " is created " + done + " / " + totalCount);
                //done++;
            }
            while (entityCollection.MoreRecords);
        }

        public static void UpdateExistingSystemView(IOrganizationService service, IOrganizationService serviceSer, int? ObjectTypeCode, int? ObjectTypeCodeSer)
        {
            Console.WriteLine("********************** VIEWs CREATION FOR ENTITY " + ObjectTypeCode.ToString().ToUpper() + " IS STARTED **********************");
            var query = new QueryExpression("savedquery");
            query.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, ObjectTypeCode);
            query.ColumnSet = new ColumnSet(true);
            RetrieveMultipleRequest retrieveSavedQueriesRequest = new RetrieveMultipleRequest { Query = query };
            RetrieveMultipleResponse retrieveSavedQueriesResponse = (RetrieveMultipleResponse)service.Execute(retrieveSavedQueriesRequest);
            DataCollection<Entity> savedQueries = retrieveSavedQueriesResponse.EntityCollection.Entities;
            //Display the Retrieved views
            Console.WriteLine("totla Views Count " + savedQueries.Count);
            int totalCount = savedQueries.Count;
            int done = 1;
            int error = 0;
            foreach (Entity ent in savedQueries)
            {
                try
                {
                    var query1 = new QueryExpression("savedquery");
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, ent.GetAttributeValue<string>("name"));
                    query.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, ObjectTypeCodeSer);
                    query1.ColumnSet = new ColumnSet(true);
                    RetrieveMultipleRequest retrieveSavedQueriesRequest1 = new RetrieveMultipleRequest { Query = query1 };
                    RetrieveMultipleResponse retrieveSavedQueriesResponse1 = (RetrieveMultipleResponse)serviceSer.Execute(retrieveSavedQueriesRequest1);
                    DataCollection<Entity> savedQueries1 = retrieveSavedQueriesResponse1.EntityCollection.Entities;
                    if (savedQueries1.Count > 0)
                    {
                        Entity entity = new Entity(savedQueries1[0].LogicalName, savedQueries1[0].Id);
                        entity["columnsetxml"] = savedQueries1[0].GetAttributeValue<string>("columnsetxml");
                        entity["fetchxml"] = savedQueries1[0].GetAttributeValue<string>("fetchxml");
                        entity["layoutjson"] = savedQueries1[0].GetAttributeValue<string>("layoutjson");
                        entity["layoutxml"] = savedQueries1[0].GetAttributeValue<string>("layoutxml");
                        serviceSer.Update(entity);
                        Console.WriteLine("View with name " + ent.GetAttributeValue<string>("name") + " is update " + done + " / " + totalCount);
                        done++;
                    }
                    else
                    {
                        ent["isdefault"] = false;
                        ent["description"] = ent.GetAttributeValue<string>("description") + " newView";
                        try
                        {
                            serviceSer.Create(ent);
                            Console.WriteLine("View with name " + ent.GetAttributeValue<string>("name") + " is created " + done + " / " + totalCount);
                            done++;
                        }
                        catch (Exception ex)
                        {
                            error++;
                            Console.WriteLine("-------------------------Failuer " + error + " in View creation with name " + ent.GetAttributeValue<string>("name") + ". Error is:-  " + ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    error++;
                    Console.WriteLine("-------------------------Failuer " + error + " in View Updation with name " + ent.GetAttributeValue<string>("name") + ". Error is:-  " + ex.Message);
                }

            }
            Console.WriteLine("********************** VIEWs CREATION FOR ENTITY " + ObjectTypeCode.ToString().ToUpper() + " IS ENDED **********************");
        }

        public static void RetriveViewDepndency(IOrganizationService serviceSer, string entityName)
        {
            Console.WriteLine("********************** VIEWs Dependenct FOR ENTITY ".ToUpper() + entityName.ToString().ToUpper() + " IS STARTED **********************");

            RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.All,
                LogicalName = entityName
            };
            RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)serviceSer.Execute(retrieveEntityRequest);
            EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;

            var query = new QueryExpression("savedquery");
            query.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, StateEntity.ObjectTypeCode);
            query.Criteria.AddCondition("iscustom", ConditionOperator.Equal, true);
            query.ColumnSet = new ColumnSet(true);
            RetrieveMultipleRequest retrieveSavedQueriesRequest = new RetrieveMultipleRequest { Query = query };
            RetrieveMultipleResponse retrieveSavedQueriesResponse = (RetrieveMultipleResponse)serviceSer.Execute(retrieveSavedQueriesRequest);
            DataCollection<Entity> savedQueries = retrieveSavedQueriesResponse.EntityCollection.Entities;
            //Display the Retrieved views
            Console.WriteLine("totla Views Count " + savedQueries.Count);
            int totalCount = savedQueries.Count;
            int done = 1;
            int error = 0;
            foreach (Entity ent in savedQueries)
            {
                var query1 = new QueryExpression("savedquery");
                query1.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, StateEntity.ObjectTypeCode);
                query1.Criteria.AddCondition("name", ConditionOperator.Equal, ent.GetAttributeValue<string>("name"));
                query1.Criteria.AddCondition("iscustom", ConditionOperator.Equal, false);
                query1.ColumnSet = new ColumnSet(true);
                RetrieveMultipleRequest retrieveSavedQueriesRequest1 = new RetrieveMultipleRequest { Query = query1 };
                RetrieveMultipleResponse retrieveSavedQueriesResponse1 = (RetrieveMultipleResponse)serviceSer.Execute(retrieveSavedQueriesRequest1);
                DataCollection<Entity> savedQueries1 = retrieveSavedQueriesResponse1.EntityCollection.Entities;
                if (savedQueries1.Count != 0)
                {
                    var dependenciesRequest = new RetrieveDependenciesForDeleteRequest
                    {
                        ObjectId = ent.Id,
                        ComponentType = 26// RetrieveDependenciesForDeleteRequest.ComponentType
                    };

                    var dependenciesResponse = (RetrieveDependenciesForDeleteResponse)serviceSer.Execute(dependenciesRequest);

                    foreach (Entity entity in dependenciesResponse.EntityCollection.Entities)
                    {
                        var componentType = entity.FormattedValues["dependentcomponenttype"];
                        var componentid = entity.GetAttributeValue<Guid>("dependentcomponentobjectid");
                        var dependncy = entity.FormattedValues["dependencytype"];
                        if (entity.GetAttributeValue<OptionSetValue>("dependentcomponenttype").Value == 60)
                        {
                            updateViewsOnForm(serviceSer, ent.Id, savedQueries1[0].Id, componentid);
                        }
                    }
                    Entity view = serviceSer.Retrieve(ent.LogicalName, ent.Id, new ColumnSet(true));

                    Entity newview = new Entity(ent.LogicalName, savedQueries1[0].Id);
                    if (ent.Contains("columnsetxml"))
                        newview["columnsetxml"] = ent["columnsetxml"];
                    newview["fetchxml"] = ent["fetchxml"];
                    newview["layoutjson"] = ent["layoutjson"];
                    newview["layoutxml"] = ent["layoutxml"];
                    try
                    {
                        serviceSer.Update(newview);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }


                    try
                    {
                        serviceSer.Delete(ent.LogicalName, ent.Id);
                    }
                    catch (Exception ex)
                    {

                        Console.WriteLine(ex.Message);
                    }

                    Console.WriteLine("View with name " + ent.GetAttributeValue<string>("name") + " is update " + done + " / " + totalCount);
                    done++;
                }
                else
                {
                    Console.WriteLine("Name No dependency found");
                }
            }
            Console.WriteLine("********************** VIEWs Dependenct FOR ENTITY ".ToUpper() + entityName.ToString().ToUpper() + " IS Ended **********************".ToUpper());
        }
        public static void updateViewsOnForm(IOrganizationService serviceSer, Guid oldView, Guid newView, Guid formId)
        {
            var query = new QueryExpression("systemform");
            query.Criteria.AddCondition("formid", ConditionOperator.Equal, formId);
            query.ColumnSet = new ColumnSet(true);

            var results = serviceSer.RetrieveMultiple(query);
            var formjson = results[0].GetAttributeValue<string>("formjson");
            string oldValue = "{" + oldView.ToString() + "}";
            string newValue = "{" + newView.ToString() + "}";
            formjson = formjson.Replace(oldValue.ToUpper(), newValue.ToUpper());
            results[0]["formjson"] = formjson;
            var formxml = results[0].GetAttributeValue<string>("formxml");
            formxml = formxml.Replace(oldValue.ToUpper(), newValue.ToUpper());
            results[0]["formxml"] = formxml;

            serviceSer.Update(results[0]);

            serviceSer.Execute(new PublishXmlRequest
            {

                ParameterXml = $@"
                    <importexportxml>
                        <entities>
                            <entity>{results[0].GetAttributeValue<string>("objecttypecode")}</entity>
                        </entities>
                    </importexportxml>"
            });
            Console.WriteLine("Name ");

        }
       public static void DeleteFile(string path)
        {
            System.IO.DirectoryInfo di = new DirectoryInfo(path);

            foreach (FileInfo file in di.GetFiles())
            {
                file.Delete();
            }
        }
        public static void RetriveSitemap(IOrganizationService service, IOrganizationService serviceSer)
        {
            var query = new QueryExpression("appmodule");
            //query.Criteria.AddCondition("name", ConditionOperator.BeginsWith, "hil_");
            query.ColumnSet = new ColumnSet(true);

            var results = service.RetrieveMultiple(query);
            Console.WriteLine("totla Form Count " + results.Entities.Count);
            int totalCount = results.Entities.Count;
            int done = 1;
            int error = 0;
            foreach (Entity entity in results.Entities)
            {
                try
                {
                    query = new QueryExpression("sitemap");
                    query.ColumnSet = new ColumnSet(true);
                    query.Criteria = new FilterExpression
                    {
                        FilterOperator = LogicalOperator.Or,
                        Conditions ={
           new ConditionExpression("sitemapnameunique", ConditionOperator.Null),
              new ConditionExpression("sitemapnameunique", ConditionOperator.Equal, entity.GetAttributeValue<string>("uniquename"))
       }
                    };

                    EntityCollection sitemaps = service.RetrieveMultiple(query);
                    foreach (Entity entSiteMap in sitemaps.Entities)
                    {
                        try
                        {
                            serviceSer.Create(entSiteMap);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("-------------------------Failuer " + error + " in webResource creation with name " + entity.GetAttributeValue<string>("name") + ". Error is:-  " + ex.Message);
                        }

                    }
                }
                catch (Exception ex)
                {
                    error++;
                    Console.WriteLine("-------------------------Failuer " + error + " in webResource creation with name " + entity.GetAttributeValue<string>("name") + ". Error is:-  " + ex.Message);
                }
            }
            Console.WriteLine("********************** createMordernDerivenApp MIGRATION IS ENDED **********************");


        }
        //public static void RetriveRibbon(IOrganizationService service, IOrganizationService serviceSer, string entityName)
        //{
        //    //Retrieve the Application Ribbon
        //    RetrieveApplicationRibbonRequest appribReq = new RetrieveApplicationRibbonRequest();
        //    RetrieveApplicationRibbonResponse appribResp = (RetrieveApplicationRibbonResponse)service.Execute(appribReq);

        //    System.String applicationRibbonPath = Path.GetFullPath(exportFolder + "\\applicationRibbon.xml");
        //    File.WriteAllBytes(applicationRibbonPath, unzipRibbon(appribResp.CompressedApplicationRibbonXml));

        //    //Import

        //}
        public static void ZipFiles(string SerPath, string outputFilePath, string password = null)
        {
            ZipFile.CreateFromDirectory(SerPath, outputFilePath, CompressionLevel.Fastest, false);

        }
        public static void ImportSolution(IOrganizationService service, string ManagedSolutionLocation)
        {

            byte[] fileBytes = File.ReadAllBytes(ManagedSolutionLocation);

            ImportSolutionRequest impSolReq = new ImportSolutionRequest()
            {
                CustomizationFile = fileBytes
            };

            service.Execute(impSolReq);

            Console.WriteLine("Imported Solution from {0}", ManagedSolutionLocation);
        }
        public static byte[] unzipRibbon(byte[] data)
        {
            System.IO.Packaging.ZipPackage package = null;
            MemoryStream memStream = null;

            memStream = new MemoryStream();
            memStream.Write(data, 0, data.Length);
            package = (ZipPackage)ZipPackage.Open(memStream, FileMode.Open);

            ZipPackagePart part = (ZipPackagePart)package.GetPart(new Uri("/customizations.xml", UriKind.Relative));
            using (Stream strm = part.GetStream())
            {
                long len = strm.Length;
                byte[] buff = new byte[len];
                strm.Read(buff, 0, (int)len);
                return buff;
            }
        }
        public static void CreateSecurityRole(IOrganizationService service, IOrganizationService serviceSer)
        {
            QueryExpression query = new QueryExpression("role");
            //query.Criteria.AddCondition("parentbusinessunitid", ConditionOperator.Null);
            query.ColumnSet = new ColumnSet(true);
            query.NoLock = true;
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;
            EntityCollection entityCollection = service.RetrieveMultiple(query);
            do
            {
                foreach (Entity entity1 in entityCollection.Entities)
                {
                    try
                    {
                        if (entity1.GetAttributeValue<EntityReference>("businessunitid").Id == new Guid("574D775D-60EB-E811-A96C-000D3AF05828"))
                        {
                            entity1["businessunitid"] = new EntityReference("businessunit", new Guid("3D41A7E9-D037-ED11-9DB1-000D3AF071C9"));
                        }
                        serviceSer.Create(entity1);
                        Console.WriteLine("Done");
                    }
                    catch (Exception ex)
                    {
                        if (!ex.Message.Contains(" already exists on business unit"))
                            if (!ex.Message.Contains("Cannot insert duplicate key exception when executing non-query"))
                                Console.WriteLine("Error " + ex.Message);
                    }
                }
                //Console.WriteLine("createMordernDerivenApp with name " + entity.GetAttributeValue<string>("name") + " is created " + done + " / " + totalCount);
                //done++;
            }
            while (entityCollection.MoreRecords);
        }
        public static void createMordernDerivenApp(IOrganizationService service, IOrganizationService serviceSer)
        {
            Console.WriteLine("********************** createMordernDerivenApp MIGRATION IS STARTED **********************");
            var query = new QueryExpression("appmodule");
            //query.Criteria.AddCondition("name", ConditionOperator.BeginsWith, "hil_");
            query.ColumnSet = new ColumnSet(true);

            var results = service.RetrieveMultiple(query);
            Console.WriteLine("totla Form Count " + results.Entities.Count);
            int totalCount = results.Entities.Count;
            int done = 1;
            int error = 0;
            foreach (Entity entity in results.Entities)
            {
                try
                {
                    query = new QueryExpression("appmodule");
                    query.Criteria.AddCondition("name", ConditionOperator.Equal, entity.GetAttributeValue<string>("name"));
                    query.ColumnSet = new ColumnSet(true);
                    var resultsApp = serviceSer.RetrieveMultiple(query);
                    if (resultsApp.Entities.Count == 0)
                    {
                        serviceSer.Create(entity);
                        //query = new QueryExpression("appmodulecomponent");
                        ////query.Criteria.AddCondition("appmoduleidunique", ConditionOperator.Equal, entity.Id);
                        //query.ColumnSet = new ColumnSet(true);
                        //query.NoLock = true;
                        //query.PageInfo = new PagingInfo();
                        //query.PageInfo.Count = 5000;
                        //query.PageInfo.PageNumber = 1;
                        //query.PageInfo.ReturnTotalRecordCount = true;
                        //EntityCollection entityCollection = service.RetrieveMultiple(query);
                        //do
                        //{
                        //    foreach (Entity entity1 in entityCollection.Entities)
                        //    {
                        //        if (entity1.GetAttributeValue<EntityReference>("appmoduleidunique").Name == entity.GetAttributeValue<string>("name"))
                        //        {
                        //            Console.WriteLine("ddd");
                        //            serviceSer.Create(entity1);
                        //        }
                        //    }
                        //    Console.WriteLine("createMordernDerivenApp with name " + entity.GetAttributeValue<string>("name") + " is created " + done + " / " + totalCount);
                        //    done++;
                        //}
                        //while (entityCollection.MoreRecords);
                    }
                    else
                    {
                        entity.Id = resultsApp[0].Id;
                        serviceSer.Create(entity);
                        Console.WriteLine("createMordernDerivenApp with name " + entity.GetAttributeValue<string>("name") + " is created " + done + " / " + totalCount);
                        done++;
                    }
                }
                catch (Exception ex)
                {
                    error++;
                    Console.WriteLine("-------------------------Failuer " + error + " in webResource creation with name " + entity.GetAttributeValue<string>("name") + ". Error is:-  " + ex.Message);
                }
            }
            Console.WriteLine("********************** createMordernDerivenApp MIGRATION IS ENDED **********************");
        }
        public static void createTransactionCurrency(IOrganizationService service, IOrganizationService serviceSer)
        {
            QueryExpression query = new QueryExpression("transactioncurrency");
            //query.Criteria.AddCondition("appmoduleidunique", ConditionOperator.Equal, entity.Id);
            query.ColumnSet = new ColumnSet(true);
            query.NoLock = true;
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;
            EntityCollection entityCollection = service.RetrieveMultiple(query);
            do
            {
                foreach (Entity entity1 in entityCollection.Entities)
                {
                    try
                    {
                        QueryExpression query1 = new QueryExpression("transactioncurrency");
                        //query.Criteria.AddCondition("appmoduleidunique", ConditionOperator.Equal, entity.Id);
                        query1.ColumnSet = new ColumnSet(true);
                        EntityCollection entityCollection1 = serviceSer.RetrieveMultiple(query1);
                        Entity entity = new Entity(entityCollection1[0].LogicalName, entityCollection1[0].Id);
                        entity["isocurrencycode"] = "ALL";
                        entity["currencyname"] = "NNN";
                        serviceSer.Update(entity);
                        serviceSer.Create(entity1);

                        Console.WriteLine("Done");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error " + ex.Message);
                    }
                }
                //Console.WriteLine("createMordernDerivenApp with name " + entity.GetAttributeValue<string>("name") + " is created " + done + " / " + totalCount);
                //done++;
            }
            while (entityCollection.MoreRecords);
        }
        public static void createBusinessUnit(IOrganizationService service, IOrganizationService serviceSer)
        {
            QueryExpression query = new QueryExpression("businessunit");
            query.Criteria.AddCondition("parentbusinessunitid", ConditionOperator.Null);
            query.ColumnSet = new ColumnSet(true);
            query.NoLock = true;
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;
            EntityCollection entityCollection = service.RetrieveMultiple(query);
            do
            {
                foreach (Entity entity1 in entityCollection.Entities)
                {
                    try
                    {
                        CreateChildBU(service, serviceSer, new Guid("574D775D-60EB-E811-A96C-000D3AF05828"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error " + ex.Message);
                    }


                }
                //Console.WriteLine("createMordernDerivenApp with name " + entity.GetAttributeValue<string>("name") + " is created " + done + " / " + totalCount);
                //done++;
            }
            while (entityCollection.MoreRecords);
        }
        public static void CreateChildBU(IOrganizationService service, IOrganizationService serviceSer, Guid parentBUID)
        {
            QueryExpression query = new QueryExpression("businessunit");
            query.Criteria.AddCondition("parentbusinessunitid", ConditionOperator.Equal, parentBUID);
            query.ColumnSet = new ColumnSet(true);
            query.NoLock = true;
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;
            EntityCollection entityCollection = service.RetrieveMultiple(query);
            do
            {
                //Entity[] entities = entityCollection.Entities.ToArray();
                // Parallel.ForEach(entities, entity =>
                foreach (Entity entity1 in entityCollection.Entities)
                {
                    try
                    {
                        //        Entity entity1 = entity;
                        if (entity1.GetAttributeValue<EntityReference>("parentbusinessunitid").Id == new Guid("574D775D-60EB-E811-A96C-000D3AF05828"))
                        {
                            QueryExpression query1 = new QueryExpression("businessunit");
                            //query.Criteria.AddCondition("appmoduleidunique", ConditionOperator.Equal, entity.Id);
                            query1.ColumnSet = new ColumnSet(true);
                            EntityCollection entityCollection1 = serviceSer.RetrieveMultiple(query1);
                            entity1["parentbusinessunitid"] = entityCollection1[0].ToEntityReference();
                            Console.WriteLine("Done");
                        }

                        entity1["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("a4705b2e-1b38-ed11-9db1-000d3af071c9"));
                        parentBUID = serviceSer.Create(entity1);
                        Console.WriteLine("Done");
                        CreateChildBU(service, serviceSer, parentBUID);

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error " + ex.Message);
                    }


                };
                //Console.WriteLine("createMordernDerivenApp with name " + entity.GetAttributeValue<string>("name") + " is created " + done + " / " + totalCount);
                //done++;
            }
            while (entityCollection.MoreRecords);
        }
        public static void CreateN2NRelationShip(EntityMetadata StateEntity, IOrganizationService serviceSer)
        {
            Console.WriteLine("********************** N to N RELATIONSHIP CREATION FOR ENTITY " + StateEntity.LogicalName.ToUpper() + " IS STARTED **********************");
            Console.WriteLine("totla Attribute Count " + StateEntity.ManyToManyRelationships.Length);
            int totalCount = StateEntity.ManyToManyRelationships.Length;
            int done = 1;
            int error = 0;
            foreach (object relationShip in StateEntity.ManyToManyRelationships)
            {
                ManyToManyRelationshipMetadata relation = (ManyToManyRelationshipMetadata)relationShip;
                try
                {
                    CreateManyToManyRequest createManyToManyRelationshipRequest = new CreateManyToManyRequest
                    {
                        IntersectEntitySchemaName = relation.IntersectEntityName,
                        ManyToManyRelationship = relation
                    };

                    CreateManyToManyResponse createManytoManyRelationshipResponse = (CreateManyToManyResponse)serviceSer.Execute(createManyToManyRelationshipRequest);
                    Console.WriteLine("RelationShip is Created " + done + " / " + totalCount + " with IntersectEntityName name " + relation.IntersectEntityName);
                }
                catch (Exception ex)
                {
                    if (!ex.Message.Contains("s not unique within an entity"))
                    {
                        error++;
                        Console.WriteLine("-------------------------Failuer " + error + " in attribute creation with IntersectEntityName name " + relation.IntersectEntityName + " Error is:-  " + ex.Message);
                    }
                }

            }
            Console.WriteLine("********************** N to N RELATIONSHIP CREATION FOR ENTITY " + StateEntity.LogicalName.ToUpper() + " IS ENDED **********************");

        }
        public static void UpdateGlobalOptionSet(IOrganizationService service, IOrganizationService serviceSer)
        {
            Console.WriteLine("********************** UPDATE GLOBAL OPTIONSET METHOD **********************");
            string optionsSetName = "socialprofile_isblocked;socialprofile_community;socialactivity_postmessagetype;sharepoint_generic_yes_no;serviceappointment_status;quote_createrevisedquote;purchasetimeframe;purchaseprocess;processstage_category;need;msdyncrm_targettype;msdyncrm_fittype;msdyncrm_borderstyletype;msdyncrm_backgroundwidthtype;msdyncrm_backgroundsizetype;msdyncrm_addtocalendarchoiceoptions;msdyn_wosystemstatus;msdyn_workstartlocationtype;msdyn_worklocation;msdyn_weekday;msdyn_wbsnodelevel;msdyn_visibility;msdyn_upgradestatus;msdyn_travelchargetype;msdyn_transactiontypecode;msdyn_transactionrole;msdyn_transactionclassification;msdyn_timescale;msdyn_timeoffrecordstatus;msdyn_timeentrytype;msdyn_timeentrystatus;msdyn_timeentrysourcetype;msdyn_surveystage;msdyn_surveyactiontype;msdyn_suggestiontype;msdyn_suggestioncontroltype;msdyn_srooptions;msdyn_solutionarea_type;msdyn_slatype;msdyn_sessiontypeoptions;msdyn_screenpoptimeout;msdyn_scoredefinition;msdyn_schedule;msdyn_salesplay_type;msdyn_salesmotion_type;msdyn_rtvsystemstatus;msdyn_rmasystemstatus;msdyn_rmaproductstatus;msdyn_rmaprocessingaction;msdyn_reviewstatus;msdyn_resulttype;msdyn_responseretrievaltype;msdyn_responsemappingset;msdyn_resourceschedulesource;msdyn_reservationtype;msdyn_requirementstatus;msdyn_requirementgrouporder;msdyn_relationship_cardinality;msdyn_ratingimagetype;msdyn_ratingimagesize;msdyn_quicknote_type;msdyn_questiongrouptype;msdyn_querystringset;msdyn_purchaseorderproductstatus;msdyn_projecttaskstatusindicators;msdyn_projectinvoicestatus;msdyn_projectcontractstatus;msdyn_projectcontractstate;msdyn_profitability;msdyn_productservicestatus;msdyn_productcostorder;msdyn_pricelistentity;msdyn_pricecalculation;msdyn_predictivescoringgrade;msdyn_predictivescoretrend;msdyn_posystemstatus;msdyn_postconversationsurveymode;msdyn_postconversationsurveyenable;msdyn_poshiptotype;msdyn_pooltype;msdyn_poapprovalstatus;msdyn_pipetypes;msdyn_personalmessage_localefield;msdyn_paymenttype;msdyn_partytype;msdyn_parametertype;msdyn_panelstateoptions;msdyn_outofstockoptions;msdyn_opportunityscoretrendoptset;msdyn_opportunitygradeoptset;msdyn_ocsystemmessagetype;msdyn_ocmessagereceiver;msdyn_occurrenceofweekday;msdyn_notificationtheme;msdyn_netpromoterscore;msdyn_module;msdyn_mltrainingstatus;msdyn_linkquestions;msdyn_linetype;msdyn_levelofimportance;msdyn_leadscoretrendoptset;msdyn_leadgradeoptset;msdyn_languagecodes;msdyn_jobstatus;msdyn_invoicestatus;msdyn_invoicerunstatus;msdyn_inventorytransactiontype;msdyn_inventoryjournaltype;msdyn_integrationtype;msdyn_integrationjobstatus;msdyn_inspectionstatus;msdyn_inspectionresult;msdyn_importactiontype;msdyn_generictype;msdyn_generateresponsedataoptions;msdyn_frequency;msdyn_findworkeventtype;msdyn_fieldserviceproducttype;msdyn_feedbackruntimelookuptype;msdyn_feedbackgenerationstatus;msdyn_feedbackattributetype;msdyn_feasibility;msdyn_facttype;msdyn_facialexpressiontype;msdyn_facialexpressionset;msdyn_facemodelset;msdyn_extensiontype;msdyn_exportstatus;msdyn_expensetypes;msdyn_expensestatus;msdyn_expensecategorybehavior;msdyn_eventtype;msdyn_estimateheadertype;msdyn_entitlementappliesto;msdyn_emailtemplatetype;msdyn_durationroundingpolicy;msdyn_distanceunit;msdyn_displaylogic;msdyn_deviceevent;msdyn_desktopnotificationvisibility;msdyn_daysofrun;msdyn_dayofmonth;msdyn_createsurveyresponsealert;msdyn_conversation_statuscode;msdyn_conversation_statecode;msdyn_connectortype;msdyn_computablefields;msdyn_competitive;msdyn_committype;msdyn_characteristictype;msdyn_changesource;msdyn_calculatesurveyscore;msdyn_budgetestimate;msdyn_bookingsystemstatus;msdyn_bookingsource;msdyn_bookingmethod;msdyn_bookingjournaltype;msdyn_bookableresourcetype;msdyn_billingtype;msdyn_billingstatus;msdyn_billingmethod;msdyn_autoupdatebookingtraveltype;msdyn_autocreateinvoices;msdyn_asyncoperationstatus;msdyn_approvalstate;msdyn_applicationtype;msdyn_apiversionoptions;msdyn_amountmethod;msdyn_alignment;msdyn_agreementsystemstatus;msdyn_agreementinvoicestatus;msdyn_agreementbookingstatus;msdyn_aggregationtype;msdyn_agentinputlanguage;msdyn_adjustmentstatus;msdyn_activitylinktype;initialcommunication;incident_caseorigincode;identifypursuitteam;identifycustomercontacts;identifycompetitors;fullsyncstate;evaluatefit;entitytype;decisionmaker;convert_campaign_response_to_lead_qualify_status;convert_campaign_response_to_lead_option;convert_campaign_response_to_lead_disqualify_status;convert_campaign_response_qualify_lead_options;convert_campaign_response_options;connectionrole_category;confirminterest;channelaccessprofile_webaccess;channelaccessprofile_viewknowledgearticles;channelaccessprofile_viewarticlerating;channelaccessprofile_twitteraccess;channelaccessprofile_submitfeedback;channelaccessprofile_rateknowledgearticles;channelaccessprofile_phoneaccess;channelaccessprofile_isguestprofile;channelaccessprofile_haveprivilegeschanged;channelaccessprofile_facebookaccess;channelaccessprofile_emailaccess;cardtype_ispreviewcard;cardtype_isliveonly;cardtype_isenabled;cardtype_isbasecard;capability;budgetstatus;botsharingroletypes;botcomponentreusepolicy;bookableresourcecharacteristictype;allocationtype;adx_partnercontactroles;adx_partnerapplicationstatus;adx_likertscalesatisfaction;adx_likertscalequality;adx_likertscalelikelihood;adx_likertscaleimportance;adx_likertscalefrequency;adx_likertscaleagreement;admin_settings_feature;admin_settings;activityfileattachment_objectcode";

            string[] optionsetnames = optionsSetName.Split(';');
            Console.WriteLine("Totla Count " + optionsetnames.Length);
            int done = 0;
            int totla = optionsetnames.Length;
            int error = 0;
            foreach (string option in optionsetnames)
            {
                Console.WriteLine("Option Set Name " + option);
                try
                {
                    //Retrive all Options which want to remove
                    RetrieveOptionSetRequest retrieveOptionSetRequest = new RetrieveOptionSetRequest
                    {
                        Name = option
                    };
                    // Execute the request.
                    RetrieveOptionSetResponse retrieveOptionSetResponse = (RetrieveOptionSetResponse)serviceSer.Execute(retrieveOptionSetRequest);
                    try
                    {
                        OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;
                        Console.WriteLine("Optons Set retrived");
                        //remove options 
                        foreach (OptionMetadata i in retrievedOptionSetMetadata.Options)
                        {
                            try
                            {
                                DeleteOptionValueRequest deleteOptionValueRequest = new DeleteOptionValueRequest
                                {
                                    OptionSetName = option,
                                    Value = (int)i.Value
                                };
                                // Execute the request.
                                serviceSer.Execute(deleteOptionValueRequest);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("-------------------------Error in Removing Value of " + option + " of value " + i.Value + " Error is- " + ex.Message);
                            }
                        }

                        //Create new Options
                        RetrieveOptionSetRequest retrieveOptionSetRequestPrd = new RetrieveOptionSetRequest
                        {
                            Name = option
                        };
                        // Execute the request.
                        RetrieveOptionSetResponse retrieveOptionSetResponsePrd = (RetrieveOptionSetResponse)service.Execute(retrieveOptionSetRequestPrd);
                        OptionSetMetadata retrievedOptionSetMetadataPrd = (OptionSetMetadata)retrieveOptionSetResponsePrd.OptionSetMetadata;

                        foreach (OptionMetadata i in retrievedOptionSetMetadataPrd.Options)
                        {
                            try
                            {
                                InsertOptionValueRequest insertOptionValueRequest = new InsertOptionValueRequest
                                {
                                    OptionSetName = option,
                                    Description = i.Description,
                                    ExtensionData = i.ExtensionData,
                                    Label = i.Label,
                                    ParentValues = i.ParentValues,
                                    Value = i.Value
                                };
                                var _insertedOptionValue = ((InsertOptionValueResponse)serviceSer.Execute(insertOptionValueRequest)).NewOptionValue;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("-------------------------Error in Removing Value of " + option + " of value " + i.Value + " Error is- " + ex.Message);
                            }
                        }

                        PublishXmlRequest pxReq2 = new PublishXmlRequest
                        {
                            ParameterXml =
                            String.Format("<importexportxml><optionsets><optionset>{0}</optionset></optionsets></importexportxml>", option)
                        };
                        //serviceSer.Execute(pxReq2);
                        Console.WriteLine("Optons Set published");

                        done++;
                        Console.WriteLine(done + " / " + totla);
                    }
                    catch (Exception ex1)
                    {
                        if (ex1.Message == "Unable to cast object of type 'Microsoft.Xrm.Sdk.Metadata.BooleanOptionSetMetadata' to type 'Microsoft.Xrm.Sdk.Metadata.OptionSetMetadata'.")
                        {
                        }
                        else
                        {
                            error++;
                            Console.WriteLine("-------------------------Error " + error + " in Updation OptionSet with name " + option + " Error is :- " + ex1.Message);
                        }
                    }


                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("Could not find an optionset with name"))
                    {
                        try
                        {
                            RetrieveOptionSetRequest retrieveOptionSetRequest =
                                new RetrieveOptionSetRequest
                                {
                                    Name = option
                                };
                            // Execute the request.
                            RetrieveOptionSetResponse retrieveOptionSetResponse = (RetrieveOptionSetResponse)service.Execute(retrieveOptionSetRequest);
                            OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;
                            Console.WriteLine("Optons Set retrived");
                            CreateOptionSetRequest createOptionSetRequest = new CreateOptionSetRequest
                            {
                                OptionSet = retrievedOptionSetMetadata
                            };
                            CreateOptionSetResponse optionsResp = (CreateOptionSetResponse)serviceSer.Execute(createOptionSetRequest);
                            Console.WriteLine("Optons Set Created");
                            PublishXmlRequest pxReq2 = new PublishXmlRequest { ParameterXml = String.Format("<importexportxml><optionsets><optionset>{0}</optionset></optionsets></importexportxml>", option) };
                            serviceSer.Execute(pxReq2);
                            Console.WriteLine("Optons Set published");

                            Console.WriteLine(done + " / " + totla);
                            done++;
                        }
                        catch (Exception ex1)
                        {
                            error++;
                            Console.WriteLine("-------------------------Error " + error + " in Createing OptionSet with name " + option + " Error is :- " + ex1.Message);
                        }
                    }
                    else
                    {
                        error++;
                        Console.WriteLine("-------------------------Error " + error + " in Updation OptionSet with name " + option + " Error is :- " + ex.Message);
                    }
                }
            }
            Console.WriteLine("********************** UPDATE GLOBAL OPTIONSET METHOD ENDs **********************");
        }
        public static void CreateEmailTemplate(IOrganizationService service, IOrganizationService serviceSer)
        {
            Console.WriteLine("********************** EMAIL TEMPLATE MIGRATION IS STARTED **********************");
            var query = new QueryExpression("template");
            query.ColumnSet = new ColumnSet(true);
            var results = service.RetrieveMultiple(query);
            Console.WriteLine("totla Email Template Count " + results.Entities.Count);
            int totalCount = results.Entities.Count;
            int done = 1;
            int error = 0;
            foreach (Entity entity in results.Entities)
            {
                try
                {

                    entity["ownerid"] = Program.SystemAdmin;
                    entity["owninguser"] = Program.SystemAdmin;

                    serviceSer.Create(entity);

                    Console.WriteLine("EMAIL TEMPLATE with name " + entity.GetAttributeValue<string>("title") + " is created " + done + " / " + totalCount);
                    done++;
                }
                catch (Exception ex)
                {
                    error++;
                    Console.WriteLine("-------------------------Failuer " + error + " in EMAIL TEMPLATE creation with name " +
                        entity.GetAttributeValue<string>("title") + ". Error is:-  " + ex.Message);
                }
            }
            Console.WriteLine("********************** EMAIL TEMPLATE MIGRATION IS ENDED **********************");

        }
        public static void craeteworkflow(IOrganizationService service, IOrganizationService serviceSer)
        {
            var query = new QueryExpression("workflow");
            //query.Criteria.AddCondition("primaryentity", ConditionOperator.Equal, 10359);
            query.ColumnSet = new ColumnSet(true);
            EntityCollection entityCollection = service.RetrieveMultiple(query);
            foreach (Entity e in entityCollection.Entities)
            {
                var aa = e.GetAttributeValue<OptionSetValue>("type");
                //e["type"] = new OptionSetValue(1);
                e["statuscode"] = new OptionSetValue(1);
                EntityReference owner = SystemAdmin;
                Console.WriteLine("dd");
                e["ownerid"] = SystemAdmin;
                serviceSer.Create(e);
            }

        }
        public static void CreateWebResource(IOrganizationService service, IOrganizationService serviceSer)
        {
            Console.WriteLine("********************** WEBRESOURCE MIGRATION IS STARTED **********************");
            string web = "lat_/CRMRESTBuilder/scripts/Sdk.Soap.min.js";// "dyn_result_32;dyn_result_16;ispl_AutoNumberIcon_32;ispl_AutoNumberIcon_16;hil_WorkOrderNew.js;hil_WorkOrderProduct;hil_WorkOrder;hil_WOIncident;hil_WOClaims;hil_Webresource/JS/Case.js;hil_WebResource/JS/AdvisoryEnqueryLine.js;hil_WebResource/JS/AdvisoryEnquery.js;hil_WebResource/HTML/Attachment/WebPage.html;hil_WarrantlyDetail;hil_validateAMC;hil_UserSecurityRoleExtension;hil_TimeSlot;hil_TestQuickCreate;hil_Tendertechno1;hil_Tendertechno;hil_TenderProduct;hil_TenderPo1;hil_TenderPo;hil_TenderPaymentdetailsJs;hil_Tendergtpuploaded1;hil_Tendergtpuploaded;hil_Tenderfinaltcp1;hil_Tenderfinaltcp;hil_TenderFinalgtpuploaded1;hil_TenderFinalgtpuploaded;hil_TenderFinal1;hil_TenderFinal;hil_TenderDesigntaskcompleted1;hil_TenderDesigntaskcompleted;hil_TenderButton;hil_BankGuarnteeDetails;hil_tenderbankguarntee1;hil_TenderFormEvents;hil_tenderFieldEvents;hil_technicianJS;hil_TechnicalSpecification;hil_TechnicalLocker;hil_SubmittoDesignteam1;hil_SubmittoDesignteam;hil_style;hil_SolarServey;hil_SMSTemplate;hil_slick;hil_SetSubStatusCanceled_LockJobForm;hil_SetJobIDonload;hil_Scheme_validate;hil_sawActivityApproval;new_Reversegeocoding;hil_ReturnLine;hil_ReturnHeader;hil_WebResource/HTML/RemotePay.html;hil_Reject.svg;hil_qrcode.min.js;hil_qr-code.jpg;hil_purchaseorderimage;hil_ProjectUtils.js;hil_profileimageguidelines;hil_ProductionSupport;hil_ProductRequestsValidation;hil_ProductRequest;hil_PostOrderCheckList;hil_JSPOStatusTracker;hil_POStatusHistoryTracker;hil_PMSvalidate;hil_pmsconfiglines;hil_photo2;hil_PhoneCall;hil_PerformaInvoice.js;hil_part_validate;hil_OrderCheckListProductJS;hil_OrderCheckListJS;hil_OAProductJS;hil_OALineinspectionStatus;hil_OAHeaderAttachment;hil_oAHeader;hil_namejs;hil_MinimunStockLevel;new_map;hil_LocalPurchase;hil_lead;hil_labor_validate;hil_Jquery3.6.0;hil_jquery_1.9.1.min.js;hil_jquery.min.js;hil_JobEstimates;hil_JobCancelIMAGE;hil_JobServices;hil_JobProducts;hil_IncomingCall;hil_imageview;hil_yellowIcon.png;hil_Subgridformatting;hil_redIcon.png;hil_greenIcon.png;hil_HavellsJSLib.js;hil_Havells.WarrantyTemplate.js;hil_Havells.WarrantyTemplate.FieldEvents.js;hil_Havells.Utility.js;hil_GRNLine;new_GetAllEntityList;hil_GeoFencingMapView;hil_EnquiryStatus;hil_ems;hil_DigiLockerTest;hil_DigiLocker;hil_DesignDashboard;hil_DailyHealthIndicatorLine;hil_hil_CustomerAssetPortal;hil_Customerasset;hil_custom;hil_ContactNew.js;hil_Contact;hil_ClaimOverheadLineJs;hil_claimRemarks;hil_Characterstics;hil_care-360.jpg;hil_care-360;hil_CampaignMainLibrary;hil_Bootstrap_min_js_3.3.7;hil_Bootstrap_min_css_3.3.7;hil_bootstrap.min.js;hil_bootstrap.min;hil_BankGurabteeValidation;hil_WebResource/HTML/AuditLog/AuditLogTable.html;hil_Attachment;hil_ArchivedJobView;hil_ArchivedJobJS;hil_AquaParametersValidation;hil_Approve.svg;hil_Approve;hil_Approval;hil_address;hil_ACInstallationCheckList;hil_Account;hil_/css/alert.css;hil_communication;hil_EnquiryBomLine;hil_HealthIndicatorIcon;hil_HavellsLogo;hil_group-logo.jpg;";
            //"hil_validateAMC";// "adx_scripts /jquery1.9.1.min.js;adx_scripts/jquery1.9.1.min.js;msdyn_SDK.REST.js;mag_/js/process.js";
            string[] webs = web.Split(';');

            try
            {
                var query = new QueryExpression("webresource");
                query.Criteria.AddCondition("name", ConditionOperator.In, webs);
                query.ColumnSet = new ColumnSet(true);
                var results = service.RetrieveMultiple(query);
                int count = results.Entities.Count;
                int done = 1;
                int error = 1;
                foreach (Entity webresource in results.Entities)
                {
                    try
                    {
                        var query1 = new QueryExpression("webresource");
                        query1.Criteria.AddCondition("name", ConditionOperator.Equal, webresource["name"]);
                        query1.ColumnSet = new ColumnSet(true);
                        var results1 = serviceSer.RetrieveMultiple(query1);
                        if (results1.Entities.Count == 0)
                        {
                            serviceSer.Create(webresource);
                            Console.WriteLine("Done " + done + "/" + count + " Record Created. " + webresource["name"]);
                            done++;
                        }
                        else
                        {
                            results1[0]["content"] = webresource["content"];
                            //entity["webresourceidunique"] = results1[0].Id;
                            //entity.Id = results1[0].Id;
                            serviceSer.Update(results1[0]);
                            Console.WriteLine("Done " + done + "/" + count + " Record updated. " + webresource["name"]);
                            done++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error with webresource name " + webresource["name"] + " Error is " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error with webresource name Error is " + ex.Message);
            }

            Console.WriteLine("********************** WEBRESOURCE MIGRATION IS ENDED **********************");

        }
        public static void CreateGlobalOptionSet(IOrganizationService service, IOrganizationService serviceSer)
        {
            Console.WriteLine("********************** CREATE GLOBAL OPTIONSET METHOD **********************");
            string optionsSetName = "adx_partnercontactroles";// "hil_yesnona;hil_yesno;hil_warrantysubstatus;hil_voltage;hil_vaccinationstatus;hil_typeofproduct;hil_tenderstatus;hil_tenderstakeholder;hil_syncstatus;hil_sourceofcreation;hil_smstemplatetype;hil_slastatus;hil_serviceengineerstatus;hil_serialnumbercount;hil_sawcategoryentrymode;hil_sawapprovalstatus;hil_salutation;hil_returntype;hil_requesttype;hil_recordtype;hil_prtype;hil_performainvoicestatus;hil_paymentterm;hil_paymentstatus;hil_operator;hil_nomineerelationship;hil_maritalstatus;hil_level;hil_leadtype;hil_joberrorsstatus;hil_jobclass;hil_jobclaimstatus;hil_jobadditionalactions;hil_inventorytype;hil_interest;hil_incentivetype;hil_incentivecategory;hil_icutomainwirespecs;hil_hierarchylevel;hil_gdatatype;hil_franchiseecategory;hil_enquirytype;hil_disposition;hil_discounttype;hil_departmentenquiry;hil_customerfeedback;hil_countryclassification;hil_consumernonconsumer;hil_claimstatus;hil_chargetype;hil_category;hil_callcenter;hil_brand;hil_bloodgroup;hil_availabilitystatus;hil_approvalstatus;hil_approvallevel;hil_approvalentitystatus;hil_activitygstslab;hil_aboutsyncappdownload";
            //"new_customertype;hil_yesnona;hil_yesno;hil_warrantysubstatus;hil_voltage;hil_vaccinationstatus;hil_typeofproduct;hil_tenderstatus;hil_tenderstakeholder;hil_syncstatus;hil_sourceofcreation;hil_smstemplatetype;hil_slastatus;hil_serviceengineerstatus;hil_serialnumbercount;hil_sawcategoryentrymode;hil_sawapprovalstatus;hil_salutation;hil_returntype;hil_requesttype;hil_recordtype;hil_prtype;hil_performainvoicestatus;hil_paymentterm;hil_paymentstatus;hil_operator;hil_nomineerelationship;hil_maritalstatus;hil_level;hil_leadtype;hil_joberrorsstatus;hil_jobclass;hil_jobclaimstatus;hil_jobadditionalactions;hil_inventorytype;hil_interest;hil_incentivetype;hil_incentivecategory;hil_icutomainwirespecs;hil_hierarchylevel;hil_gdatatype;hil_franchiseecategory;hil_enquirytype;hil_disposition;hil_discounttype;hil_departmentenquiry;hil_customerfeedback;hil_countryclassification;hil_consumernonconsumer;hil_claimstatus;hil_chargetype;hil_category;hil_callcenter;hil_brand;hil_bloodgroup;hil_availabilitystatus;hil_approvalstatus;hil_approvallevel;hil_approvalentitystatus;hil_activitygstslab;hil_aboutsyncappdownload";
            string[] optionsetnames = optionsSetName.Split(';');
            Console.WriteLine("Totla Count " + optionsetnames.Length);
            int done = 0;
            int totla = optionsetnames.Length;
            int error = 0;

            foreach (string option in optionsetnames)
            {
                Console.WriteLine("Option Set Name " + option);
                try
                {
                    RetrieveOptionSetRequest retrieveOptionSetRequest =
                        new RetrieveOptionSetRequest
                        {
                            Name = option
                        };
                    // Execute the request.
                    RetrieveOptionSetResponse retrieveOptionSetResponse = (RetrieveOptionSetResponse)service.Execute(retrieveOptionSetRequest);
                    OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;
                    Console.WriteLine("Optons Set retrived");
                    CreateOptionSetRequest createOptionSetRequest = new CreateOptionSetRequest
                    {
                        OptionSet = retrievedOptionSetMetadata
                    };
                    CreateOptionSetResponse optionsResp = (CreateOptionSetResponse)serviceSer.Execute(createOptionSetRequest);
                    Console.WriteLine("Optons Set Created");
                    PublishXmlRequest pxReq2 = new PublishXmlRequest { ParameterXml = String.Format("<importexportxml><optionsets><optionset>{0}</optionset></optionsets></importexportxml>", option) };
                    serviceSer.Execute(pxReq2);
                    Console.WriteLine("Optons Set published");

                    Console.WriteLine(done + " / " + totla);
                    done++;
                }
                catch (Exception ex)
                {
                    error++;
                    Console.WriteLine("-------------------------Error " + error + " in Createing OptionSet with name " + option + " Error is :- " + ex.Message);
                }
            }
            Console.WriteLine("********************** CREATE GLOBAL OPTIONSET METHOD ENDs **********************");
        }
        public static void EntityMigration(IOrganizationService service, IOrganizationService serviceSer)
        {
            Console.WriteLine("********************** ENTITY MIGRATEION METHOD **********************");
            string EntityNames = "new_userlocations";//"new_integrationstaging;plt_idgenerator";
            //"hil_address;hil_advancefind;hil_advisormaster;hil_advisoryenquiry;hil_alerts;hil_alternatepart;hil_amcdiscountmatrix;hil_amcplan;hil_amcstaging;hil_approval;hil_approvalmatrix;hil_aquaparameters;hil_archivedactivity;hil_archvedjobdatasource;hil_area;hil_assignmentmatrix;hil_assignmentmatrixupload;hil_attachment;hil_attachmentdocumenttype;hil_auditlog;hil_autonumber;hil_bdtatdetails;hil_bdtatownership;hil_bdteam;hil_bdteammember;hil_bizgeomappingstagigng;hil_bomcategory;hil_bomsubcategory;hil_branch;hil_bulkjobsuploader;hil_businessmapping;hil_callsubtype;hil_calltype;hil_campaigndivisions;hil_campaignenquirytypes;hil_campaignwebsitesetup;hil_cancellationreason;hil_casecategory;hil_channelpartnercountryclassification;hil_city;hil_claimallocator;hil_claimcategory;hil_claimheader;hil_claimline;hil_claimlines;hil_claimoverheadline;hil_claimperiod;hil_claimpostingsetup;hil_claimsummary;hil_claimtype;hil_constructiontype;hil_consumerappbridge;hil_consumercategory;hil_consumernpssurvey;hil_consumertype;hil_country;hil_customerassetstagging;hil_customerfeedback;hil_customerwishlist;hil_dealerstockverificationheader;hil_deliveryschedule;hil_deliveryschedulemaster;hil_designteambranchmapping;hil_despatchteammaster;hil_discountmatrix;hil_distributionchannel;hil_district;hil_effeciency;hil_enclosuretype;hil_enquerysegment;hil_enquirybusinessprocessflow;hil_enquirydepartment;hil_enquirydocumenttype;hil_enquirylostreason;hil_enquiryproductsegment;hil_enquirysegmentdcmapping;hil_enquirytype;hil_entityfieldsmetadata;hil_errorcode;hil_escalation;hil_escalationmatrix;hil_estimate;hil_feedback;hil_fixedcompensation;hil_forgotpassword;hil_frequency;hil_grnline;hil_healthindicatorheader;hil_homeadvisoryline;hil_hsncode;hil_icauploader;hil_incentivetable;hil_industrysubtype;hil_industrytype;hil_inspectiontype;hil_installationchecklist;hil_integrationconfiguration;hil_integrationjob;hil_integrationjobrun;hil_integrationjobrundetail;hil_integrationtrace;hil_inventory;hil_inventoryjournal;hil_inventoryrequest;hil_invoice;hil_jobbasecharge;hil_jobcancellationrequest;hil_joberrorcode;hil_jobestimation;hil_jobreassignreason;hil_jobsextension;hil_jobtat;hil_jobtatcategory;hil_labor;hil_leadproduct;hil_materialgroup;hil_materialgroup2;hil_materialgroup3;hil_materialgroup4;hil_materialgroup5;hil_materialreturn;hil_minimumguarantee;hil_minimumstocklevel;hil_mobileappbanner;hil_natureofcomplaint;hil_npssetup;hil_oaheader;hil_oaproduct;hil_oareadinessdatestaging;hil_observation;hil_orderchecklist;hil_orderchecklistproduct;hil_part;hil_partnerdepartment;hil_partnerdivisionmapping;hil_paymentstatus;hil_pincode;hil_plantmaster;hil_plantordertypesetup;hil_pmsconfiguration;hil_pmsconfigurationlines;hil_pmsjobuploader;hil_pmsscheduleconfiguration;hil_politicalmapping;hil_postatustracker;hil_postatustrackerupload;hil_postorder;hil_preferredlanguageforcommunication;hil_productcatalog;hil_production;hil_productrequest;hil_productrequestheader;hil_productrequestheaderarchived;hil_productstaging;hil_productvideos;hil_propertytype;hil_refreshjobs;hil_region;hil_remarks;hil_returnheader;hil_returnline;hil_salesoffice;hil_salesofficebranchheadmapping;hil_salesofficeplantmapping;hil_salestatmaster;hil_sawactivity;hil_sawactivityapproval;hil_sawcategoryapprovals;hil_sbubranchmapping;hil_schemedistrictexclusion;hil_schemeincentive;hil_schemeline;hil_securityroleextension;hil_serialnumber;hil_serviceactionwork;hil_serviceactionworksetup;hil_servicebom;hil_servicecallrequest;hil_servicecallrequestdetail;hil_serviceengineergeocode;hil_smsconfiguration;hil_smstemplates;hil_solarservey;hil_specialincentive;hil_staging;hil_stagingdivisonmaterialgroupmapping;hil_stagingintegrationjson;hil_stagingpricingmapping;hil_stagingsbudivisionmapping;hil_startingmethod;hil_state;hil_statustransitionmatrix;hil_subterritory;hil_tatachievementslabmaster;hil_tatbreachpenalty;hil_tatincentive;hil_technicalspecfication;hil_technician;hil_technicianhealthindicator;hil_tender;hil_tenderattachmentdoctype;hil_tenderattachmentmanager;hil_tenderbankguarantee;hil_tenderbomlineitem;hil_tenderdocs;hil_tenderemailersetup;hil_tenderpaymentdetail;hil_tenderproduct;hil_tolerencesetup;hil_travelclosureincentive;hil_travelexpense;hil_typeofcustomer;hil_typeofproduct;hil_unitwarranty;hil_upcountrydataupdate;hil_upcountrytravelcharge;hil_usageindicator;hil_userbranchmapping;hil_usersecurityroleextension;hil_usersimeibinding;hil_voltage;hil_warrantyperiod;hil_warrantyscheme;hil_warrantytemplate;hil_warrantyvoidreason;hil_whatsappproductdivisionconfig;hil_workorderarch;hil_wrongclosurepenalty;hil_yammersettings;dyn_advancefindconfig;ispl_autonumberconfiguration";
            string[] entities = EntityNames.Split(';');
            Console.WriteLine("totla Entity Count " + entities.Length);
            int totalCount = entities.Length;
            int done = 1;
            int error = 0;
            //Console.WriteLine("Entity Creation Started");

            #region createEntity
            foreach (string entityName in entities)
            {
                try
                {
                    //if (entityName == "plt_idgenerator")
                    {
                        Console.WriteLine("Enity With Name " + entityName + " started");
                        //updateEntity(service, serviceSer, entityName);
                        createEntity(service, serviceSer, entityName);
                        Console.WriteLine("Entity Created " + done + " / " + totalCount);
                        done++;
                    }
                }
                catch (Exception ex)
                {
                    error++;
                    Console.WriteLine("-------------------------Failuer " + error + " in entity creation with name " + entityName + " Error is:-  " + ex.Message);
                }
            }
            #endregion

            #region CreateFieldsExceptLookUp
            Console.WriteLine("Fields Creation Started");
            foreach (string ent in entities)
            {
                try
                {
                    RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
                    {
                        EntityFilters = EntityFilters.All,
                        LogicalName = ent
                    };
                    RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
                    EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;
                    CreateFieldsExceptLookUp(StateEntity, serviceSer);
                    Console.WriteLine("Entity Created " + done + " / " + totalCount);
                    done++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("-------------------------Failuer in entity MetaData creation with name " + ent + " Error is:-  " + ex.Message);
                }
            }
            #endregion

            #region createNto1RelationShip
            Console.WriteLine("Relationship Creation Started");
            //create relationship
            foreach (string ent in entities)
            {
                try
                {
                    RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
                    {
                        EntityFilters = EntityFilters.All,
                        LogicalName = ent
                    };
                    RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
                    EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;
                    createNto1RelationShip(StateEntity, serviceSer);
                    //CreateSystemView(service, serviceSer, StateEntity.ObjectTypeCode);
                    //CreateUserView(service, serviceSer, StateEntity.ObjectTypeCode);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("-------------------------Failuer in entity MetaData creation with name " + ent + " Error is:-  " + ex.Message);
                }
            }
            #endregion

            #region CreateN2NRelationShip
            //  create Views
            Console.WriteLine("Relationship N2N Creation Started");
            foreach (string ent in entities)
            {
                try
                {
                    RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
                    {
                        EntityFilters = EntityFilters.All,
                        LogicalName = ent
                    };
                    RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
                    EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;
                    CreateN2NRelationShip(StateEntity, serviceSer);
                    Console.WriteLine("Entity Created " + done + " / " + totalCount);
                    done++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("-------------------------Failuer in entity MetaData creation with name " + ent + " Error is:-  " + ex.Message);
                }
            }
            #endregion

            #region CreateSystemView
            Console.WriteLine("View Creation Started");

            foreach (string ent in entities)
            {
                try
                {
                    RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
                    {
                        EntityFilters = EntityFilters.All,
                        LogicalName = ent
                    };
                    RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
                    EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;
                    CreateSystemView(service, serviceSer, StateEntity.ObjectTypeCode, 0);
                    CreateUserView(service, serviceSer, StateEntity.ObjectTypeCode);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("-------------------------Failuer in entity MetaData creation with name " + ent + " Error is:-  " + ex.Message);
                }
            }
            #endregion

            #region CreateForm
            Console.WriteLine("Forms Creation Started");
            foreach (string ent in entities)
            {
                RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.All,
                    LogicalName = ent
                };
                RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
                EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;

                RetrieveEntityRequest retrieveEntityRequestDemo = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.Entity,
                    LogicalName = ent
                };
                RetrieveEntityResponse retrieveAccountEntityResponseDemo = (RetrieveEntityResponse)serviceSer.Execute(retrieveEntityRequestDemo);
                EntityMetadata StateEntityDemo = retrieveAccountEntityResponseDemo.EntityMetadata;

                CreateForm(StateEntity, serviceSer, service, StateEntityDemo.ObjectTypeCode);
            }
            #endregion
        }
        public static void createEntity(IOrganizationService service, IOrganizationService serviceSer, string entityName)
        {
            RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.All,
                LogicalName = entityName
            };
            RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
            EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;

            Label displayName = new Label();
            Label discreption = new Label();
            foreach (AttributeMetadata attr in StateEntity.Attributes)
            {
                if (attr.LogicalName == StateEntity.PrimaryNameAttribute)
                {
                    displayName = attr.DisplayName;
                    discreption = attr.Description;
                    goto myCode;
                }
            }
        myCode:
            EntityMetadata entityMetadata = new EntityMetadata
            {
                SchemaName = StateEntity.SchemaName,
                DisplayName = StateEntity.DisplayName,//new Label("Bank Account", 1033),
                DisplayCollectionName = StateEntity.DisplayCollectionName,//new Label("Bank Accounts", 1033),
                Description = StateEntity.Description,//new Label("An entity to store information about customer bank accounts", 1033),
                OwnershipType = StateEntity.OwnershipType,//OwnershipTypes.UserOwned,
                IsActivity = StateEntity.IsActivity,
                IsAuditEnabled = StateEntity.IsAuditEnabled,//
                IsActivityParty = StateEntity.IsActivityParty,//
                IsBPFEntity = StateEntity.IsBPFEntity,//
                IsAvailableOffline = StateEntity.IsAvailableOffline,
                IsBusinessProcessEnabled = StateEntity.IsBusinessProcessEnabled,
                IsConnectionsEnabled = StateEntity.IsConnectionsEnabled,
                IsCustomizable = StateEntity.IsCustomizable,
                IsDocumentManagementEnabled = StateEntity.IsDocumentManagementEnabled,
                IsDocumentRecommendationsEnabled = StateEntity.IsDocumentRecommendationsEnabled,
                IsDuplicateDetectionEnabled = StateEntity.IsDuplicateDetectionEnabled,
                IsEnabledForExternalChannels = StateEntity.IsEnabledForExternalChannels,
                IsInteractionCentricEnabled = StateEntity.IsInteractionCentricEnabled,
                IsKnowledgeManagementEnabled = StateEntity.IsKnowledgeManagementEnabled,
                IsMailMergeEnabled = StateEntity.IsMailMergeEnabled,
                IsMappable = StateEntity.IsMappable,
                IsMSTeamsIntegrationEnabled = StateEntity.IsMSTeamsIntegrationEnabled,
                IsOfflineInMobileClient = StateEntity.IsOfflineInMobileClient,
                IsOneNoteIntegrationEnabled = StateEntity.IsOneNoteIntegrationEnabled,
                IsQuickCreateEnabled = StateEntity.IsQuickCreateEnabled,
                IsReadingPaneEnabled = StateEntity.IsReadingPaneEnabled,
                ActivityTypeMask = StateEntity.ActivityTypeMask,
                AutoCreateAccessTeams = StateEntity.AutoCreateAccessTeams,
                AutoRouteToOwnerQueue = StateEntity.AutoRouteToOwnerQueue,
                CanChangeHierarchicalRelationship = StateEntity.CanChangeHierarchicalRelationship,
                CanChangeTrackingBeEnabled = StateEntity.CanChangeTrackingBeEnabled,
                CanCreateAttributes = StateEntity.CanCreateAttributes,
                CanCreateCharts = StateEntity.CanCreateCharts,
                CanCreateForms = StateEntity.CanCreateForms,
                CanCreateViews = StateEntity.CanCreateViews,
                CanEnableSyncToExternalSearchIndex = StateEntity.CanEnableSyncToExternalSearchIndex,
                CanModifyAdditionalSettings = StateEntity.CanModifyAdditionalSettings,
                ChangeTrackingEnabled = StateEntity.ChangeTrackingEnabled,
                DataProviderId = StateEntity.DataProviderId,
                DataSourceId = StateEntity.DataSourceId,
                DaysSinceRecordLastModified = StateEntity.DaysSinceRecordLastModified,
                EntityColor = StateEntity.EntityColor,
                EntityHelpUrl = StateEntity.EntityHelpUrl,
                EntityHelpUrlEnabled = StateEntity.EntityHelpUrlEnabled,
                EntitySetName = StateEntity.EntitySetName,
                ExtensionData = StateEntity.ExtensionData,
                ExternalCollectionName = StateEntity.ExternalCollectionName,
                ExternalName = StateEntity.ExternalName,
                HasActivities = StateEntity.HasActivities,
                HasChanged = StateEntity.HasChanged,
                HasEmailAddresses = StateEntity.HasEmailAddresses,
                HasFeedback = StateEntity.HasFeedback,
                HasNotes = StateEntity.HasNotes,
                IconLargeName = StateEntity.IconLargeName,
                IconMediumName = StateEntity.IconMediumName,
                IconSmallName = StateEntity.IconSmallName,
                IconVectorName = StateEntity.IconVectorName,
                IsReadOnlyInMobileClient = StateEntity.IsReadOnlyInMobileClient,
                IsRenameable = StateEntity.IsRenameable,
                IsRetrieveAuditEnabled = StateEntity.IsRetrieveAuditEnabled,
                IsRetrieveMultipleAuditEnabled = StateEntity.IsRetrieveMultipleAuditEnabled,
                IsSLAEnabled = StateEntity.IsSLAEnabled,
                IsSolutionAware = StateEntity.IsSolutionAware,
                IsValidForQueue = StateEntity.IsValidForQueue,
                IsVisibleInMobile = StateEntity.IsVisibleInMobile,
                IsVisibleInMobileClient = StateEntity.IsVisibleInMobileClient,
                LogicalCollectionName = StateEntity.LogicalCollectionName,
                LogicalName = StateEntity.LogicalName,
                //MetadataId = StateEntity.MetadataId,
                MobileOfflineFilters = StateEntity.MobileOfflineFilters,
                OwnerId = StateEntity.OwnerId,
                OwnerIdType = StateEntity.OwnerIdType,
                OwningBusinessUnit = StateEntity.OwningBusinessUnit,
                SettingOf = StateEntity.SettingOf,
                Settings = StateEntity.Settings,
                SyncToExternalSearchIndex = StateEntity.SyncToExternalSearchIndex,
                UsesBusinessDataLabelTable = StateEntity.UsesBusinessDataLabelTable,
                MetadataId=StateEntity.MetadataId
            };
            if (StateEntity.IsActivity == true)
            {

                CreateEntityRequest createrequest = new CreateEntityRequest
                {
                    HasActivities = false,
                    HasNotes = true,
                    //Define the entity
                    Entity = entityMetadata,
                    PrimaryAttribute = new StringAttributeMetadata
                    {
                        SchemaName = "Subject",
                        RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.SystemRequired),
                        MaxLength = 100,
                        FormatName = StringFormatName.Text,
                        DisplayName = displayName,
                        Description = discreption
                    }

                };
                serviceSer.Execute(createrequest);
            }
            else
            {
                CreateEntityRequest createrequest = new CreateEntityRequest
                {
                    //Define the entity
                    HasActivities = (bool)StateEntity.HasActivities,
                    HasFeedback = (bool)StateEntity.HasFeedback,
                    HasNotes = (bool)StateEntity.HasNotes,
                    Entity = entityMetadata,
                    PrimaryAttribute = new StringAttributeMetadata
                    {
                        SchemaName = StateEntity.PrimaryNameAttribute,
                        RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.SystemRequired),
                        MaxLength = 100,
                        FormatName = StringFormatName.Text,
                        DisplayName = displayName,
                        Description = discreption
                    }
                };
                serviceSer.Execute(createrequest);
            }
        }
        public static void CreateFieldsExceptLookUp(EntityMetadata StateEntity, IOrganizationService serviceSer)
        {
            Console.WriteLine("********************** FIELD CREATION FOR ENTITY " + StateEntity.LogicalName.ToUpper() + " IS STARTED **********************");
            Console.WriteLine("totla Attribute Count " + StateEntity.Attributes.Length);
            int totalCount = StateEntity.Attributes.Length;
            int done = 1;
            int error = 0;
            if (!StateEntity.LogicalName.Contains("archived"))
                foreach (AttributeMetadata a in StateEntity.Attributes)
                {
                    //if (a.LogicalName.ToLower() != "ispl_preview".ToLower())
                    //    continue;
                    if (a.AttributeTypeName.Value == "MultiSelectPicklistType")
                    {
                        OptionSetMetadata optionSetMetadata = ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)a).OptionSet;
                        if (optionSetMetadata.IsGlobal == true)
                            optionSetMetadata.Options.Clear();
                        ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)a).OptionSet = optionSetMetadata;
                        CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                        {
                            EntityName = StateEntity.LogicalName,
                            Attribute = a
                        };
                        try
                        {
                            serviceSer.Execute(createAttributeRequest);
                            Console.WriteLine("attribute Created " + done + " / " + totalCount + " with name " + a.LogicalName);
                            done++;
                        }
                        catch (Exception ex)
                        {
                            error++;
                            Console.WriteLine("-------------------------Failuer " + error + " in attribute creation with name " + a.LogicalName + " Error is:-  " + ex.Message);
                        }

                    }
                    if (a.LogicalName.Contains("_"))
                    {

                        if (!a.LogicalName.Contains("_base") && a.AttributeType != AttributeTypeCode.Picklist && a.IsValidForForm == true
                            && a.LogicalName != StateEntity.PrimaryNameAttribute && a.AttributeType != AttributeTypeCode.Lookup
                            && a.AttributeType != AttributeTypeCode.Customer && a.AttributeType != AttributeTypeCode.Uniqueidentifier)
                        {
                            CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                            {
                                EntityName = StateEntity.LogicalName,
                                Attribute = a
                            };
                            try
                            {
                                serviceSer.Execute(createAttributeRequest);
                                Console.WriteLine("attribute Created " + done + " / " + totalCount + " with name " + a.LogicalName);
                                done++;
                            }
                            catch (Exception ex)
                            {
                                if (ex.Message.Contains("because there is already an Attribute"))
                                {
                                    Console.WriteLine(a.LogicalName + " is Already Exist ");
                                }
                                else
                                {
                                    error++;
                                    Console.WriteLine("-------------------------Failuer " + error + " in attribute creation with name " + a.LogicalName + " Error is:-  " + ex.Message);
                                }
                            }
                        }
                        else if (a.AttributeType == AttributeTypeCode.Picklist)
                        {


                            OptionSetMetadata optionSetMetadata = ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)a).OptionSet;
                            if (optionSetMetadata.IsGlobal == true)
                                optionSetMetadata.Options.Clear();
                            ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)a).OptionSet = optionSetMetadata;
                            CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                            {
                                EntityName = StateEntity.LogicalName,
                                Attribute = a
                            };
                            try
                            {
                                serviceSer.Execute(createAttributeRequest);
                                Console.WriteLine("attribute Created " + done + " / " + totalCount + " with name " + a.LogicalName);
                                done++;
                            }
                            catch (Exception ex)
                            {
                                error++;
                                Console.WriteLine("-------------------------Failuer " + error + " in attribute creation with name " + a.LogicalName + " Error is:-  " + ex.Message);
                            }
                        }

                        else
                        {
                            Console.WriteLine("attribute Skiped " + done + " / " + totalCount + " with name " + a.LogicalName); done++;
                        }
                    }
                    else
                    {
                        if (a.AttributeType == AttributeTypeCode.Status)
                        {

                            RetrieveEntityRequest retrieveEntityRequestSer = new RetrieveEntityRequest
                            {
                                EntityFilters = EntityFilters.All,
                                LogicalName = StateEntity.LogicalName
                            };
                            RetrieveEntityResponse retrieveAccountEntityResponseSer = (RetrieveEntityResponse)serviceSer.Execute(retrieveEntityRequestSer);
                            EntityMetadata StateEntitySer = retrieveAccountEntityResponseSer.EntityMetadata;

                            var aa = StateEntitySer.Attributes;

                            StatusAttributeMetadata lookField = (StatusAttributeMetadata)Array.Find(aa, ele => ele.SchemaName.ToLower() == a.SchemaName);
                            if (lookField != null)
                            {
                                OptionSetMetadata optionSetMetadata = ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)a).OptionSet;
                                lookField.OptionSet = optionSetMetadata;
                                UpdateAttributeRequest createAttributeRequest = new UpdateAttributeRequest
                                {
                                    EntityName = StateEntity.LogicalName,
                                    Attribute = lookField
                                };
                                try
                                {
                                    serviceSer.Execute(createAttributeRequest);
                                    Console.WriteLine("attribute Created " + done + " / " + totalCount + " with name " + a.LogicalName);
                                    done++;
                                }
                                catch (Exception ex)
                                {
                                    error++;
                                    Console.WriteLine("-------------------------Failuer " + error + " in attribute creation with name " + a.LogicalName + " Error is:-  " + ex.Message);
                                }
                            }
                            Console.Write("");
                        }
                        else
                            Console.WriteLine("attribute Skiped " + done + " / " + totalCount + " with name " + a.LogicalName); done++;
                    }
                }
            Console.WriteLine("********************** FIELD CREATION FOR ENTITY " + StateEntity.LogicalName.ToUpper() + " IS ENDED **********************");
        }
        public static void UpdateStatusField(EntityMetadata StateEntity, IOrganizationService serviceSer)
        {
            Console.WriteLine("********************** FIELD CREATION FOR ENTITY " + StateEntity.LogicalName.ToUpper() + " IS STARTED **********************");
            Console.WriteLine("totla Attribute Count " + StateEntity.Attributes.Length);
            int totalCount = StateEntity.Attributes.Length;
            int done = 1;
            int error = 0;
            if (!StateEntity.LogicalName.Contains("archived"))
                foreach (AttributeMetadata a in StateEntity.Attributes)
                {
                    if (a.SchemaName.ToLower() == "statuscode")
                    {
                        Console.WriteLine("********************** FIELD CREATION FOR ENTITY " + StateEntity.LogicalName.ToUpper() + " IS ENDED **********************");
                        if (a.AttributeType == AttributeTypeCode.Status)
                        {

                            RetrieveEntityRequest retrieveEntityRequestSer = new RetrieveEntityRequest
                            {
                                EntityFilters = EntityFilters.All,
                                LogicalName = StateEntity.LogicalName
                            };
                            RetrieveEntityResponse retrieveAccountEntityResponseSer = (RetrieveEntityResponse)serviceSer.Execute(retrieveEntityRequestSer);
                            EntityMetadata StateEntitySer = retrieveAccountEntityResponseSer.EntityMetadata;

                            var aa = StateEntitySer.Attributes;

                            StatusAttributeMetadata lookField = (StatusAttributeMetadata)Array.Find(aa, ele => ele.SchemaName.ToLower() == a.SchemaName);

                            OptionSetMetadata optionSetMetadata = ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)a).OptionSet;

                            a.MetadataId = lookField.MetadataId;

                            lookField.OptionSet = optionSetMetadata;


                            UpdateAttributeRequest createAttributeRequest = new UpdateAttributeRequest
                            {
                                EntityName = StateEntity.LogicalName,
                                Attribute = lookField
                            };
                            try
                            {
                                serviceSer.Execute(createAttributeRequest);
                                Console.WriteLine("attribute Created " + done + " / " + totalCount + " with name " + a.LogicalName);
                                done++;
                            }
                            catch (Exception ex)
                            {
                                error++;
                                Console.WriteLine("-------------------------Failuer " + error + " in attribute creation with name " + a.LogicalName + " Error is:-  " + ex.Message);
                            }
                        }
                    }
                    //if (a.LogicalName.Contains("_"))
                    //{

                    //    if (!a.LogicalName.Contains("_base") && a.AttributeType != AttributeTypeCode.Picklist && a.IsValidForForm == true && a.LogicalName != StateEntity.PrimaryNameAttribute && a.AttributeType != AttributeTypeCode.Lookup && a.AttributeType != AttributeTypeCode.Customer && a.AttributeType != AttributeTypeCode.Uniqueidentifier)
                    //    {
                    //        CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                    //        {
                    //            EntityName = StateEntity.LogicalName,
                    //            Attribute = a
                    //        };
                    //        try
                    //        {
                    //            serviceSer.Execute(createAttributeRequest);
                    //            Console.WriteLine("attribute Created " + done + " / " + totalCount + " with name " + a.LogicalName);
                    //            done++;
                    //        }
                    //        catch (Exception ex)
                    //        {
                    //            if (ex.Message.Contains("because there is already an Attribute"))
                    //            {
                    //                Console.WriteLine(a.LogicalName + " is Already Exist ");
                    //            }
                    //            else
                    //            {
                    //                error++;
                    //                Console.WriteLine("-------------------------Failuer " + error + " in attribute creation with name " + a.LogicalName + " Error is:-  " + ex.Message);
                    //            }
                    //        }
                    //    }
                    //    else if (a.AttributeType == AttributeTypeCode.Picklist)
                    //    {


                    //        OptionSetMetadata optionSetMetadata = ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)a).OptionSet;
                    //        if (optionSetMetadata.IsGlobal == true)
                    //            optionSetMetadata.Options.Clear();
                    //        ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)a).OptionSet = optionSetMetadata;
                    //        CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                    //        {
                    //            EntityName = StateEntity.LogicalName,
                    //            Attribute = a
                    //        };
                    //        try
                    //        {
                    //            serviceSer.Execute(createAttributeRequest);
                    //            Console.WriteLine("attribute Created " + done + " / " + totalCount + " with name " + a.LogicalName);
                    //            done++;
                    //        }
                    //        catch (Exception ex)
                    //        {
                    //            error++;
                    //            Console.WriteLine("-------------------------Failuer " + error + " in attribute creation with name " + a.LogicalName + " Error is:-  " + ex.Message);
                    //        }
                    //    }
                    //    else
                    //    {
                    //        Console.WriteLine("attribute Skiped " + done + " / " + totalCount + " with name " + a.LogicalName); done++;
                    //    }
                    //}
                    //else
                    //{

                    //    Console.WriteLine("attribute Skiped " + done + " / " + totalCount + " with name " + a.LogicalName); done++;
                    //}
                }
            Console.WriteLine("********************** FIELD CREATION FOR ENTITY " + StateEntity.LogicalName.ToUpper() + " IS ENDED **********************");
        }
        public static void createNto1RelationShip(EntityMetadata StateEntity, IOrganizationService serviceSer)
        {
            Console.WriteLine("********************** N to 1 RELATIONSHIP CREATION FOR ENTITY " + StateEntity.LogicalName.ToUpper() + " IS STARTED **********************");
            Console.WriteLine("totla Attribute Count " + StateEntity.ManyToOneRelationships.Length);
            int totalCount = StateEntity.ManyToOneRelationships.Length;
            int done = 1;
            int error = 0;
            foreach (object relationShip in StateEntity.ManyToOneRelationships)
            {
                OneToManyRelationshipMetadata relation = (OneToManyRelationshipMetadata)relationShip;
                try
                {
                    if (relation.ReferencingAttribute.Contains("_"))
                    {
                        //Console.WriteLine("RelatoinShip with ReferencingAttribu");
                        AttributeMetadata[] a = StateEntity.Attributes;
                        AttributeMetadata lookField = Array.Find(a, ele => ele.SchemaName.ToLower() == relation.ReferencingAttribute.ToLower());
                        relation.ReferencingAttribute = null;
                        if (lookField.AttributeType == AttributeTypeCode.Customer)
                            continue;
                        CreateOneToManyRequest createOneToManyRelationshipRequest = new CreateOneToManyRequest
                        {
                            OneToManyRelationship = relation,
                            Lookup = (LookupAttributeMetadata)lookField
                            //new LookupAttributeMetadata
                            //{
                            //    SchemaName = lookField.SchemaName.Contains("_") ? lookField.SchemaName : "hil_" + lookField.SchemaName,
                            //    LogicalName = lookField.LogicalName.Contains("_") ? lookField.LogicalName : "hil_" + lookField.LogicalName,
                            //    DisplayName = lookField.DisplayName,
                            //    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),

                            //}
                        };
                        CreateOneToManyResponse createOneToManyRelationshipResponse = (CreateOneToManyResponse)serviceSer.Execute(createOneToManyRelationshipRequest);
                        Console.WriteLine("attribute Created " + done + " / " + totalCount + " with Lookup name " + lookField.LogicalName);
                        done++;
                    }
                    else
                        Console.WriteLine("RelatoinShip with ReferencingAttribute name " + relation.ReferencingAttribute + " is skiped due to Default field");

                    //if (relation.ReferencingAttribute.Contains("_"))
                    //{
                    //    AttributeMetadata[] a = StateEntity.Attributes;
                    //    AttributeMetadata lookField = Array.Find(a, ele => ele.SchemaName.ToLower() == relation.ReferencingAttribute.ToLower());
                    //    relation.ReferencingAttribute = null;
                    //    CreateOneToManyRequest createOneToManyRelationshipRequest = new CreateOneToManyRequest
                    //    {
                    //        OneToManyRelationship = relation,
                    //        Lookup = (LookupAttributeMetadata)lookField
                    //        //new LookupAttributeMetadata
                    //        //{
                    //        //    SchemaName = lookField.SchemaName.Contains("_") ? lookField.SchemaName : "hil_" + lookField.SchemaName,
                    //        //    LogicalName = lookField.LogicalName.Contains("_") ? lookField.LogicalName : "hil_" + lookField.LogicalName,
                    //        //    DisplayName = lookField.DisplayName,
                    //        //    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),

                    //        //}
                    //    };
                    //    CreateOneToManyResponse createOneToManyRelationshipResponse = (CreateOneToManyResponse)serviceSer.Execute(createOneToManyRelationshipRequest);
                    //    Console.WriteLine("attribute Created " + done + " / " + totalCount + " with ReferencingAttribute name " + relation.ReferencingAttribute);
                    //    done++;
                    //}
                    //else
                    //    Console.WriteLine("RelatoinShip with ReferencingAttribute name " + relation.ReferencingAttribute + " is skiped due to Default field");

                }
                catch (Exception ex)
                {
                    if (!ex.Message.Contains("not unique within an entity"))
                    {
                        error++;
                        Console.WriteLine("-------------------------Failuer " + error + " in attribute creation with ReferencingAttribute name " + relation.ReferencingAttribute + " Error is:-  " + ex.Message);
                    }
                }

            }
            Console.WriteLine("********************** N to 1 RELATIONSHIP CREATION FOR ENTITY " + StateEntity.LogicalName.ToUpper() + " IS ENDED **********************");

        }
        public static void CreateForm(EntityMetadata StateEntity, IOrganizationService serviceSer, IOrganizationService service, int? ObjectTypeCodeDemo)
        {
            //Console.Clear();
            Console.WriteLine("********************** FORM CREATION FOR ENTITY " + StateEntity.LogicalName.ToUpper() + " IS STARTED **********************");
            var query = new QueryExpression("systemform");
            query.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, StateEntity.ObjectTypeCode);
            query.ColumnSet = new ColumnSet(true);

            var results = service.RetrieveMultiple(query);
            Console.WriteLine("totla Form Count " + results.Entities.Count);
            int totalCount = results.Entities.Count;
            int done = 1;
            int error = 0;
            foreach (Entity entity in results.Entities)
            {
                try
                {
                    var query1 = new QueryExpression("systemform");
                    query1.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, ObjectTypeCodeDemo);
                    if (entity.GetAttributeValue<string>("name") == "Job")
                    {
                        query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Work Order");// entity.GetAttributeValue<string>("name"));
                    }
                    else if (entity.GetAttributeValue<string>("name") == "Job Service - Mobile")
                    {
                        query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Work Order Service - Mobile");// entity.GetAttributeValue<string>("name"));
                    }
                    else if (entity.GetAttributeValue<string>("name") == "Consumer")
                    {
                        query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Contact");// entity.GetAttributeValue<string>("name"));
                    }
                    else
                        query1.Criteria.AddCondition("name", ConditionOperator.Equal, entity.GetAttributeValue<string>("name"));
                    query1.Criteria.AddCondition("type", ConditionOperator.Equal, entity.GetAttributeValue<OptionSetValue>("type").Value);
                    //entity.FormattedValues["type"]
                    query1.ColumnSet = new ColumnSet(true);
                    var results1 = serviceSer.RetrieveMultiple(query1);
                    if (results1.Entities.Count > 0)
                    {

                        if (results1[0].FormattedValues["type"] == "Main")
                        {
                            Console.WriteLine(results1[0].FormattedValues["type"]);
                        }

                        results1[0]["description"] = entity.GetAttributeValue<string>("description");
                        results1[0]["formactivationstate"] = entity.GetAttributeValue<OptionSetValue>("formactivationstate");
                        results1[0]["formjson"] = entity.GetAttributeValue<string>("formjson");
                        results1[0]["formpresentation"] = entity.GetAttributeValue<OptionSetValue>("formpresentation");
                        results1[0]["formxml"] = entity.GetAttributeValue<string>("formxml");
                        results1[0]["isdefault"] = entity.GetAttributeValue<bool>("isdefault");
                        results1[0]["isdesktopenabled"] = entity.GetAttributeValue<bool>("isdesktopenabled");
                        results1[0]["istabletenabled"] = entity.GetAttributeValue<bool>("istabletenabled");
                        results1[0]["type"] = entity.GetAttributeValue<OptionSetValue>("type");
                        results1[0]["version"] = entity.GetAttributeValue<int>("version");
                        results1[0]["name"] = entity.GetAttributeValue<string>("name");
                        // results1[0]["name"] = entity.GetAttributeValue<string>("name");

                        //entity.Id = results1[0].Id;
                        //entity["formid"] = results1[0].Id;
                        //entity["formactivationstate"] = null;
                        //entity["formpresentation"] = null;
                        //entity["componentstate"] = null;

                        //entity["introducedversion"] = null;
                        //entity["solutionid"] = null;
                        //entity["publishedon"] = null;
                        //entity["organizationid"] = null;
                        //entity["organizationid"] = null;

                        serviceSer.Update(results1[0]);
                        Console.WriteLine(results1[0].FormattedValues["type"]);
                        Console.WriteLine("From with name " + entity.GetAttributeValue<string>("name") + " is update " + done + " / " + totalCount);
                        done++;
                    }
                    else
                    {
                        serviceSer.Create(entity);
                        Console.WriteLine("From with name " + entity.GetAttributeValue<string>("name") + " is created " + done + " / " + totalCount);
                        done++;
                    }
                }
                catch (Exception ex)
                {
                    error++;
                    Console.WriteLine("-------------------------Failuer " + error + " in Form creation with name " + entity.GetAttributeValue<string>("name") +
                        ". Error is:-  " + ex.Message);
                }
            }
            Console.WriteLine("********************** FORM CREATION FOR ENTITY " + StateEntity.LogicalName.ToUpper() + " IS ENDED **********************");
        }
        public static void CreateSystemView(IOrganizationService service, IOrganizationService serviceSer, int? ObjectTypeCode, int? ObjectTypeCodeDemo)
        {
            Console.WriteLine("********************** VIEWs CREATION FOR ENTITY " + ObjectTypeCode.ToString().ToUpper() + " IS STARTED **********************");
            var query = new QueryExpression("savedquery");
            query.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, ObjectTypeCode);
            query.ColumnSet = new ColumnSet(true);
            RetrieveMultipleRequest retrieveSavedQueriesRequest = new RetrieveMultipleRequest { Query = query };
            RetrieveMultipleResponse retrieveSavedQueriesResponse = (RetrieveMultipleResponse)service.Execute(retrieveSavedQueriesRequest);
            DataCollection<Entity> savedQueries = retrieveSavedQueriesResponse.EntityCollection.Entities;
            //Display the Retrieved views
            Console.WriteLine("totla Views Count " + savedQueries.Count);
            int totalCount = savedQueries.Count;
            int done = 1;
            int error = 0;

            foreach (Entity ent in savedQueries)
            {
                try
                {

                    Console.WriteLine("View with name " + ent.GetAttributeValue<string>("name") + " is Started ");
                    //if (ent.GetAttributeValue<string>("name") == "Product Sub Category")
                    //{
                    //    serviceSer.Create(ent);
                    //    Console.WriteLine("View with name " + ent.GetAttributeValue<string>("name") + " is created " + done + " / " + totalCount);
                    //    done++;
                    //}
                    var query1 = new QueryExpression("savedquery");
                    query1.Criteria.AddCondition("savedqueryid", ConditionOperator.Equal, ent.Id);
                    query1.ColumnSet = new ColumnSet(true);
                    RetrieveMultipleRequest retrieveSavedQueriesRequest1 = new RetrieveMultipleRequest { Query = query1 };
                    RetrieveMultipleResponse retrieveSavedQueriesResponse1 = (RetrieveMultipleResponse)serviceSer.Execute(retrieveSavedQueriesRequest1);
                    DataCollection<Entity> savedQueries1 = retrieveSavedQueriesResponse1.EntityCollection.Entities;
                    if (savedQueries1.Count > 0)
                    {
                        Entity entity = new Entity(savedQueries1[0].LogicalName, savedQueries1[0].Id);
                        entity["columnsetxml"] = savedQueries1[0].GetAttributeValue<string>("columnsetxml");
                        entity["fetchxml"] = savedQueries1[0].GetAttributeValue<string>("fetchxml");
                        entity["layoutjson"] = savedQueries1[0].GetAttributeValue<string>("layoutjson");
                        entity["layoutxml"] = savedQueries1[0].GetAttributeValue<string>("layoutxml");
                        entity["name"] = savedQueries1[0].GetAttributeValue<string>("name");
                        entity["description"] = savedQueries1[0].GetAttributeValue<string>("description");
                        serviceSer.Update(entity);
                        Console.WriteLine("View with name " + ent.GetAttributeValue<string>("name") + " is update " + done + " / " + totalCount);
                        done++;
                    }
                    else
                    {
                        ent["isdefault"] = false;
                        ent["description"] = ent.GetAttributeValue<string>("description") + " newView";
                        try
                        {
                            serviceSer.Create(ent);
                            Console.WriteLine("View with name " + ent.GetAttributeValue<string>("name") + " is created " + done + " / " + totalCount);
                            done++;
                        }
                        catch (Exception ex)
                        {
                            //if (!ex.Message.Contains("Cannot insert duplicate key exception"))
                            {
                                error++;
                                Console.WriteLine("-------------------------Failuer " + error + " in View creation with name " + ent.GetAttributeValue<string>("name") + ". Error is:-  " + ex.Message);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    error++;
                    Console.WriteLine("-------------------------Failuer " + error + " in View Updation with name " + ent.GetAttributeValue<string>("name") + ". Error is:-  " + ex.Message);
                }

            }
            Console.WriteLine("********************** VIEWs CREATION FOR ENTITY " + ObjectTypeCode.ToString().ToUpper() + " IS ENDED **********************");
        }
        public static void UpdateCreateSystemView(IOrganizationService service, IOrganizationService serviceSer, int? ObjectTypeCode)
        {
            Console.WriteLine("********************** VIEWs CREATION FOR ENTITY " + ObjectTypeCode.ToString().ToUpper() + " IS STARTED **********************");
            var query = new QueryExpression("savedquery");
            query.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, ObjectTypeCode);
            query.ColumnSet = new ColumnSet(true);
            RetrieveMultipleRequest retrieveSavedQueriesRequest = new RetrieveMultipleRequest { Query = query };
            RetrieveMultipleResponse retrieveSavedQueriesResponse = (RetrieveMultipleResponse)service.Execute(retrieveSavedQueriesRequest);
            DataCollection<Entity> savedQueries = retrieveSavedQueriesResponse.EntityCollection.Entities;
            //Display the Retrieved views
            Console.WriteLine("totla Views Count " + savedQueries.Count);
            int totalCount = savedQueries.Count;
            int done = 1;
            int error = 0;
            foreach (Entity ent in savedQueries)
            {
                try
                {
                    //var query1 = new QueryExpression("savedquery");
                    //query1.Criteria.AddCondition("name", ConditionOperator.Equal, ent.GetAttributeValue<string>("name"));
                    //query1.ColumnSet = new ColumnSet(true);
                    //RetrieveMultipleRequest retrieveSavedQueriesRequest1 = new RetrieveMultipleRequest { Query = query1 };
                    //RetrieveMultipleResponse retrieveSavedQueriesResponse1 = (RetrieveMultipleResponse)serviceSer.Execute(retrieveSavedQueriesRequest1);
                    //DataCollection<Entity> savedQueries1 = retrieveSavedQueriesResponse1.EntityCollection.Entities;
                    //if (savedQueries1.Count > 0)
                    //{
                    //    Entity entity = new Entity(savedQueries1[0].LogicalName, savedQueries1[0].Id);
                    //    entity["columnsetxml"] = savedQueries1[0].GetAttributeValue<string>("columnsetxml");
                    //    entity["fetchxml"] = savedQueries1[0].GetAttributeValue<string>("fetchxml");
                    //    entity["layoutjson"] = savedQueries1[0].GetAttributeValue<string>("layoutjson");
                    //    entity["layoutxml"] = savedQueries1[0].GetAttributeValue<string>("layoutxml");
                    //    serviceSer.Update(entity);
                    //    Console.WriteLine("View with name " + ent.GetAttributeValue<string>("name") + " is update " + done + " / " + totalCount);
                    //    done++;
                    //}
                    //else
                    {
                        ent["isdefault"] = false;
                        ent["description"] = ent.GetAttributeValue<string>("description") + " newView";
                        try
                        {
                            serviceSer.Create(ent);
                            Console.WriteLine("View with name " + ent.GetAttributeValue<string>("name") + " is created " + done + " / " + totalCount);
                            done++;
                        }
                        catch (Exception ex)
                        {
                            error++;
                            Console.WriteLine("-------------------------Failuer " + error + " in View creation with name " + ent.GetAttributeValue<string>("name") + ". Error is:-  " + ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    error++;
                    Console.WriteLine("-------------------------Failuer " + error + " in View Updation with name " + ent.GetAttributeValue<string>("name") + ". Error is:-  " + ex.Message);
                }

            }
            Console.WriteLine("********************** VIEWs CREATION FOR ENTITY " + ObjectTypeCode.ToString().ToUpper() + " IS ENDED **********************");
        }
        public static void CreateUserView(IOrganizationService service, IOrganizationService serviceSer, int? ObjectTypeCode)
        {
            Console.WriteLine("********************** USER VIEWs CREATION FOR ENTITY " + ObjectTypeCode.ToString().ToUpper() + " IS STARTED **********************");
            var query = new QueryExpression("userquery");
            query.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, ObjectTypeCode);
            query.ColumnSet = new ColumnSet(true);
            RetrieveMultipleRequest retrieveSavedQueriesRequest = new RetrieveMultipleRequest { Query = query };
            RetrieveMultipleResponse retrieveSavedQueriesResponse = (RetrieveMultipleResponse)service.Execute(retrieveSavedQueriesRequest);
            DataCollection<Entity> savedQueries = retrieveSavedQueriesResponse.EntityCollection.Entities;
            //Display the Retrieved views
            Console.WriteLine("totla USER Views Count " + savedQueries.Count);
            int totalCount = savedQueries.Count;
            int done = 1;
            int error = 0;
            foreach (Entity ent in savedQueries)
            {
                try
                {
                    serviceSer.Create(ent);
                    Console.WriteLine("USER View with name " + ent.GetAttributeValue<string>("name") + " is created " + done + " / " + totalCount);
                }
                catch (Exception ex)
                {
                    error++;
                    Console.WriteLine("-------------------------Failuer " + error + " in USER View creation with name " + ent.GetAttributeValue<string>("name") + ". Error is:-  " + ex.Message);
                }
            }
            Console.WriteLine("********************** USER VIEWs CREATION FOR ENTITY " + ObjectTypeCode.ToString().ToUpper() + " IS ENDED **********************");
        }
        public static void UpdateEntityMetaData(IOrganizationService service, IOrganizationService serviceSer)
        {
            Console.WriteLine("********************** ENTITY Updation METHOD **********************");
            string EntityNames = "account;appointment;bookableresourcebooking;campaign;characteristic;contact;email;incident;lead;msdyn_customerasset;msdyn_incidenttype;msdyn_incidenttypeproduct;msdyn_resourcerequirement;msdyn_surveyresponse;msdyn_timeoffrequest;msdyn_workorder;msdyn_workorderincident;msdyn_workorderproduct;msdyn_workorderservice;msdyn_workordersubstatus;new_integrationstaging;phonecall;plt_idgenerator;product;systemuser;task";
            //  "task;systemuser;product;phonecall;msdyn_workordersubstatus;msdyn_workorderservice;msdyn_workorderproduct;msdyn_workorderincident;msdyn_workorder;msdyn_timeoffrequest;msdyn_surveyresponse;msdyn_resourcerequirement;msdyn_incidenttypeproduct;msdyn_incidenttype;msdyn_customerasset;lead;incident;email;contact;characteristic;campaign;bookableresourcebooking;appointment;account";
            string[] entities = EntityNames.Split(';');// { "hil_state", "hil_city", "hil_smstemplates", "hil_country" };//plt_idgenerator;new_integrationstaging
            Console.WriteLine("totla Entity Count " + entities.Length);
            int totalCount = entities.Length;
            int done = 1;
            int error = 0;
            Console.WriteLine("Entity Updation Started");
            //foreach (string entityName in entities)
            //{
            //    updateEntity(service, serviceSer, entityName);
            //}
            //foreach (string entityName in entities)
            //{
            //    UpdateAttribute(service, serviceSer, entityName);
            //}
            //foreach (string entityName in entities)
            //{
            //    //UpdateAttribute(service, serviceSer, entityName);
            //}
            //foreach (string entityName in entities)
            //{
            //    RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
            //    {
            //        EntityFilters = EntityFilters.All,
            //        LogicalName = entityName
            //    };
            //    RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
            //    EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;
            //    createNto1RelationShip(StateEntity, serviceSer);
            //    //Console.Clear();
            //    //UpdateCreateSystemView(service, serviceSer, StateEntity.ObjectTypeCode);
            //    //CreateUserView(service, serviceSer, StateEntity.ObjectTypeCode);
            //}
            Console.WriteLine("Relationship N2N Creation Started");
            foreach (string ent in entities)
            {
                try
                {
                    RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
                    {
                        EntityFilters = EntityFilters.All,
                        LogicalName = ent
                    };
                    RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
                    EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;
                    CreateN2NRelationShip(StateEntity, serviceSer);
                    Console.WriteLine("Entity Created " + done + " / " + totalCount);
                    done++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("-------------------------Failuer in entity MetaData creation with name " + ent + " Error is:-  " + ex.Message);
                }
            }
            //foreach (string entityName in entities)
            //{
            //    RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
            //    {
            //        EntityFilters = EntityFilters.All,
            //        LogicalName = entityName
            //    };
            //    RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
            //    EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;
            //    RetrieveEntityRequest retrieveEntityRequestDemo = new RetrieveEntityRequest
            //    {
            //        EntityFilters = EntityFilters.Entity,
            //        LogicalName = entityName
            //    };
            //    RetrieveEntityResponse retrieveAccountEntityResponseDemo = (RetrieveEntityResponse)serviceSer.Execute(retrieveEntityRequestDemo);
            //    EntityMetadata StateEntityDemo = retrieveAccountEntityResponseDemo.EntityMetadata;

            //    CreateForm(StateEntity, serviceSer, service, StateEntityDemo.ObjectTypeCode);
            //}
            Console.WriteLine("********************** ENTITY Updation METHOD END's**********************");
        }
        public static void updateEntity(IOrganizationService service, IOrganizationService serviceSer, string entityName)
        {
            Console.WriteLine("Entity name " + entityName);
            try
            {
                RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.All,
                    LogicalName = entityName
                };
                RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
                EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;

                RetrieveEntityRequest retrieveEntityRequestDEV = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.All,
                    LogicalName = entityName
                };
                RetrieveEntityResponse retrieveAccountEntityResponseDEV = (RetrieveEntityResponse)serviceSer.Execute(retrieveEntityRequestDEV);
                EntityMetadata StateEntityDEV = retrieveAccountEntityResponseDEV.EntityMetadata;

                EntityMetadata entityMetadata = new EntityMetadata
                {
                    SchemaName = StateEntity.SchemaName,
                    DisplayName = StateEntity.DisplayName,//new Label("Bank Account", 1033),
                    DisplayCollectionName = StateEntity.DisplayCollectionName,//new Label("Bank Accounts", 1033),
                    Description = StateEntity.Description,//new Label("An entity to store information about customer bank accounts", 1033),
                    OwnershipType = StateEntity.OwnershipType,//OwnershipTypes.UserOwned,
                    IsActivity = StateEntity.IsActivity,
                    IsAuditEnabled = StateEntity.IsAuditEnabled,//
                    IsActivityParty = StateEntity.IsActivityParty,//
                    IsBPFEntity = StateEntity.IsBPFEntity,//
                    IsAvailableOffline = StateEntity.IsAvailableOffline,
                    IsBusinessProcessEnabled = StateEntity.IsBusinessProcessEnabled,
                    IsConnectionsEnabled = StateEntity.IsConnectionsEnabled,
                    IsCustomizable = StateEntity.IsCustomizable,
                    IsDocumentManagementEnabled = StateEntity.IsDocumentManagementEnabled,
                    IsDocumentRecommendationsEnabled = StateEntity.IsDocumentRecommendationsEnabled,
                    IsDuplicateDetectionEnabled = StateEntity.IsDuplicateDetectionEnabled,
                    IsEnabledForExternalChannels = StateEntity.IsEnabledForExternalChannels,
                    IsInteractionCentricEnabled = StateEntity.IsInteractionCentricEnabled,
                    IsKnowledgeManagementEnabled = StateEntity.IsKnowledgeManagementEnabled,
                    IsMailMergeEnabled = StateEntity.IsMailMergeEnabled,
                    IsMappable = StateEntity.IsMappable,
                    IsMSTeamsIntegrationEnabled = StateEntity.IsMSTeamsIntegrationEnabled,
                    IsOfflineInMobileClient = StateEntity.IsOfflineInMobileClient,
                    IsOneNoteIntegrationEnabled = StateEntity.IsOneNoteIntegrationEnabled,
                    IsQuickCreateEnabled = StateEntity.IsQuickCreateEnabled,
                    IsReadingPaneEnabled = StateEntity.IsReadingPaneEnabled,
                    ActivityTypeMask = StateEntity.ActivityTypeMask,
                    AutoCreateAccessTeams = StateEntity.AutoCreateAccessTeams,
                    AutoRouteToOwnerQueue = StateEntity.AutoRouteToOwnerQueue,
                    CanChangeHierarchicalRelationship = StateEntity.CanChangeHierarchicalRelationship,
                    CanChangeTrackingBeEnabled = StateEntity.CanChangeTrackingBeEnabled,
                    CanCreateAttributes = StateEntity.CanCreateAttributes,
                    CanCreateCharts = StateEntity.CanCreateCharts,
                    CanCreateForms = StateEntity.CanCreateForms,
                    CanCreateViews = StateEntity.CanCreateViews,
                    CanEnableSyncToExternalSearchIndex = StateEntity.CanEnableSyncToExternalSearchIndex,
                    CanModifyAdditionalSettings = StateEntity.CanModifyAdditionalSettings,
                    ChangeTrackingEnabled = StateEntity.ChangeTrackingEnabled,
                    DataProviderId = StateEntity.DataProviderId,
                    DataSourceId = StateEntity.DataSourceId,
                    DaysSinceRecordLastModified = StateEntity.DaysSinceRecordLastModified,
                    EntityColor = StateEntity.EntityColor,
                    EntityHelpUrl = StateEntity.EntityHelpUrl,
                    EntityHelpUrlEnabled = StateEntity.EntityHelpUrlEnabled,
                    EntitySetName = StateEntity.EntitySetName,
                    ExtensionData = StateEntity.ExtensionData,
                    ExternalCollectionName = StateEntity.ExternalCollectionName,
                    ExternalName = StateEntity.ExternalName,
                    HasActivities = StateEntity.HasActivities,
                    HasChanged = StateEntity.HasChanged,
                    HasEmailAddresses = StateEntity.HasEmailAddresses,
                    HasFeedback = StateEntity.HasFeedback,
                    HasNotes = StateEntity.HasNotes,
                    IconLargeName = StateEntity.IconLargeName,
                    IconMediumName = StateEntity.IconMediumName,
                    IconSmallName = StateEntity.IconSmallName,
                    IconVectorName = StateEntity.IconVectorName,
                    IsReadOnlyInMobileClient = StateEntity.IsReadOnlyInMobileClient,
                    IsRenameable = StateEntity.IsRenameable,
                    IsRetrieveAuditEnabled = StateEntity.IsRetrieveAuditEnabled,
                    IsRetrieveMultipleAuditEnabled = StateEntity.IsRetrieveMultipleAuditEnabled,
                    IsSLAEnabled = StateEntity.IsSLAEnabled,
                    IsSolutionAware = StateEntity.IsSolutionAware,
                    IsValidForQueue = StateEntity.IsValidForQueue,
                    IsVisibleInMobile = StateEntity.IsVisibleInMobile,
                    IsVisibleInMobileClient = StateEntity.IsVisibleInMobileClient,
                    LogicalCollectionName = StateEntity.LogicalCollectionName,
                    LogicalName = StateEntity.LogicalName,
                    MetadataId = StateEntity.MetadataId,
                    MobileOfflineFilters = StateEntity.MobileOfflineFilters,
                    OwnerId = StateEntity.OwnerId,
                    OwnerIdType = StateEntity.OwnerIdType,
                    OwningBusinessUnit = StateEntity.OwningBusinessUnit,
                    SettingOf = StateEntity.SettingOf,
                    Settings = StateEntity.Settings,
                    SyncToExternalSearchIndex = StateEntity.SyncToExternalSearchIndex,
                    UsesBusinessDataLabelTable = StateEntity.UsesBusinessDataLabelTable
                };
                entityMetadata.MetadataId = StateEntityDEV.MetadataId;
                UpdateEntityRequest updatereq = new UpdateEntityRequest();
                updatereq.HasActivities = (bool)StateEntity.HasActivities;
                updatereq.HasFeedback = (bool)StateEntity.HasFeedback;
                updatereq.HasNotes = (bool)StateEntity.HasNotes;
                updatereq.Entity = entityMetadata;

                UpdateEntityResponse updateresp = (UpdateEntityResponse)serviceSer.Execute(updatereq);
                Console.WriteLine("Entity Updated ");

            }
            catch (Exception ex)
            {
                Console.WriteLine("------------------------- ENTITY UPDATION FAILDE WITH NAME " + entityName + " with ERROR " + ex.Message);
            }

        }
        public static void UpdateAttribute(IOrganizationService service, IOrganizationService serviceSer, string entityName)
        {
            Console.WriteLine("Entity name " + entityName);
            try
            {
                RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.All,
                    LogicalName = entityName
                };
                RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
                EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;

                RetrieveEntityRequest retrieveEntityRequestDEV = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.All,
                    LogicalName = entityName
                };
                RetrieveEntityResponse retrieveAccountEntityResponseDEV = (RetrieveEntityResponse)serviceSer.Execute(retrieveEntityRequestDEV);
                EntityMetadata StateEntityDEV = retrieveAccountEntityResponseDEV.EntityMetadata;

                foreach (AttributeMetadata attribute in StateEntity.Attributes)
                {
                    try
                    {
                        AttributeMetadata[] a = StateEntityDEV.Attributes;
                        if (attribute.AttributeType == AttributeTypeCode.Virtual || attribute.AttributeType == AttributeTypeCode.EntityName || (attribute.SchemaName == "VersionNumber" || attribute.SchemaName == "OwningBusinessUnitName"))
                            continue;
                        AttributeMetadata lookField = Array.Find(a, ele => ele.SchemaName.ToLower() == attribute.SchemaName.ToLower());
                        if (lookField != null)
                        {
                            //attribute.MetadataId = lookField.MetadataId;

                            //UpdateAttributeRequest updateAttributeRequest = new UpdateAttributeRequest
                            //{
                            //    EntityName = entityName,
                            //    Attribute = attribute
                            //};
                            //serviceSer.Execute(updateAttributeRequest);
                            //Console.WriteLine("Field Updated " + attribute.LogicalName);
                        }
                        else
                        {
                            if (!attribute.LogicalName.Contains("_base") && attribute.AttributeType != AttributeTypeCode.Picklist && attribute.IsValidForForm == true &&
                                attribute.LogicalName != StateEntity.PrimaryNameAttribute && attribute.AttributeType != AttributeTypeCode.Lookup &&
                                attribute.AttributeType != AttributeTypeCode.Customer && attribute.AttributeType != AttributeTypeCode.Virtual &&
                                attribute.AttributeType != AttributeTypeCode.Uniqueidentifier)
                            {
                                CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                                {
                                    EntityName = StateEntity.LogicalName,
                                    Attribute = attribute
                                };
                                try
                                {
                                    serviceSer.Execute(createAttributeRequest);
                                    Console.WriteLine("attribute Created with name " + attribute.LogicalName);
                                }
                                catch (Exception ex)
                                {
                                    if (ex.Message.Contains("because there is already an Attribute"))
                                    {
                                        Console.WriteLine(attribute.LogicalName + " is Already Exist ");
                                    }
                                    else
                                    {
                                        Console.WriteLine("-------------------------Failuer in attribute creation with name " + attribute.LogicalName + " Error is:-  " + ex.Message);
                                    }
                                }
                            }
                            else if (attribute.AttributeType == AttributeTypeCode.Picklist)
                            {


                                OptionSetMetadata optionSetMetadata = ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)attribute).OptionSet;
                                if (optionSetMetadata.IsGlobal == true)
                                    optionSetMetadata.Options.Clear();
                                ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)attribute).OptionSet = optionSetMetadata;
                                CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                                {
                                    EntityName = StateEntity.LogicalName,
                                    Attribute = attribute
                                };
                                try
                                {
                                    serviceSer.Execute(createAttributeRequest);
                                    Console.WriteLine("attribute Created with name " + attribute.LogicalName);
                                    //done++;
                                }
                                catch (Exception ex)
                                {

                                    Console.WriteLine("-------------------------Failuer in attribute creation with name " + attribute.LogicalName + " Error is:-  " + ex.Message);
                                }
                            }
                            //else if (attribute.AttributeType == AttributeTypeCode.Lookup)
                            //{
                            //    OneToManyRelationshipMetadata[] relationshipMetadatas = StateEntity.OneToManyRelationships;
                            //    OneToManyRelationshipMetadata relation = Array.Find(relationshipMetadatas, ele => ele.ReferencingAttribute.ToLower() == attribute.SchemaName.ToLower());
                            //    //relationShip;

                            //    foreach(OneToManyRelationshipMetadata ele in StateEntity.OneToManyRelationships)
                            //    {
                            //        if (ele.ReferencingAttribute.ToLower() == attribute.SchemaName.ToLower())
                            //        {
                            //          Console.WriteLine("dd");
                            //        }
                            //      Console.WriteLine(ele.ReferencingAttribute.ToLower() +" :: "+ attribute.SchemaName.ToLower());
                            //    }

                            //    relation.ReferencingAttribute = null;

                            //    CreateOneToManyRequest createOneToManyRelationshipRequest = new CreateOneToManyRequest
                            //    {
                            //        OneToManyRelationship = relation,
                            //        Lookup = (LookupAttributeMetadata)attribute
                            //        //new LookupAttributeMetadata
                            //        //{
                            //        //    SchemaName = lookField.SchemaName.Contains("_") ? lookField.SchemaName : "hil_" + lookField.SchemaName,
                            //        //    LogicalName = lookField.LogicalName.Contains("_") ? lookField.LogicalName : "hil_" + lookField.LogicalName,
                            //        //    DisplayName = lookField.DisplayName,
                            //        //    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),

                            //        //}
                            //    };
                            //    CreateOneToManyResponse createOneToManyRelationshipResponse = (CreateOneToManyResponse)serviceSer.Execute(createOneToManyRelationshipRequest);
                            //}
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("-------------------------Error in Updating Attribute with name " + attribute.SchemaName + " ERROR IS " + ex.Message);
                    }
                }


                Console.WriteLine("Entity Updated ");

            }
            catch (Exception ex)
            {
                Console.WriteLine(" -------------------------ENTITY UPDATION FAILDE WITH NAME " + entityName + " with ERROR " + ex.Message);
            }

        }
        private static OrganizationWebProxyClient GetCRMService()
        {

            // private const string connStr = "AuthType=ClientSecret;url={0};ClientId=41623af4-f2a7-400a-ad3a-ae87462ae44e;ClientSecret=r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

            var aadInstance = "https://login.microsoftonline.com/";
            var organizationUrl = "https://havellscrmdev.crm8.dynamics.com";

            var tenantId = "7b7dc2f5-4e6a-4004-96dd-6c7923625b25";

            var clientId = "41623af4-f2a7-400a-ad3a-ae87462ae44e";

            var clientkey = "r/qy6jgTisxeKZ7T3nYTVbhswXc5pYSVoSp5XFKJtQ8=";

            var clientcred = new ClientCredential(clientId, clientkey);

            var authenticationContext = new AuthenticationContext(aadInstance + tenantId);

            var authenticationResult = authenticationContext.AcquireTokenAsync(organizationUrl, clientcred);

            var requestedToken = authenticationResult.Result.AccessToken;

            var sdkService = new OrganizationWebProxyClient(new Uri(organizationUrl), false);
            sdkService.HeaderToken = requestedToken;
            return sdkService;
        }
        static IOrganizationService ConnectToCRM(string connectionString)
        {
            //  GetCRMService();
            IOrganizationService service = null;
            try
            {
                service = new CrmServiceClient(connectionString);

                //var sdkService = new OrganizationWebProxyClient(GetServiceUrl(organizationUrl), new TimeSpan(0, 10, 0), false);


                if (((CrmServiceClient)service).LastCrmException != null && (((CrmServiceClient)service).LastCrmException.Message == "OrganizationWebProxyClient is null" ||
                    ((CrmServiceClient)service).LastCrmException.Message == "Unable to Login to Dynamics CRM"))
                {
                    Console.WriteLine(((CrmServiceClient)service).LastCrmException.Message);
                    throw new Exception(((CrmServiceClient)service).LastCrmException.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("-------------------------Error while Creating Conn: " + ex.Message);
            }
            return service;

        }
        static void CreateField(IOrganizationService service, string _entityName, List<Models> _ExcelAttributeList)
        {

            // Create storage for new attributes being created
            List<AttributeMetadata> addedAttributes = new List<AttributeMetadata>();

            foreach (Models model in _ExcelAttributeList)
            {
                if (model.AttributeType.Trim() == "Two Option")
                {
                    BooleanOptionSetMetadata booleanOptionSet = new BooleanOptionSetMetadata();
                    BooleanOptionSetMetadata booleanOptionSetValue = new BooleanOptionSetMetadata();
                    if (model.Options.Contains("|"))
                    {
                        string[] options = model.Options.Split('|');
                        foreach (string lable in options)
                            booleanOptionSet.FalseOption.Label = new Label(lable, _languageCode);
                    }
                    if (model.OptionValue.Contains("|"))
                    {
                        string[] options = model.OptionValue.Split('|');
                        foreach (string lable in options)
                            booleanOptionSetValue.FalseOption.Label = new Label(lable, _languageCode);
                    }
                    AttributeRequiredLevel requirementLevel = AttributeRequiredLevel.None;
                    if (model.RequiredLevel.Trim() == "SystemRequired")
                    {
                        requirementLevel = AttributeRequiredLevel.SystemRequired;
                    }
                    else if (model.RequiredLevel.Trim() == "ApplicationRequired")
                    {
                        requirementLevel = AttributeRequiredLevel.ApplicationRequired;
                    }

                    // Create a boolean attribute
                    BooleanAttributeMetadata boolAttribute = new BooleanAttributeMetadata
                    {
                        // Set base properties
                        SchemaName = preFix + "_" + model.DisplayName.Replace(" ", ""),
                        LogicalName = preFix + "_" + model.DisplayName.Replace(" ", "").ToLower(),
                        DisplayName = new Label(model.DisplayName, _languageCode),
                        IsAuditEnabled = new BooleanManagedProperty(true),
                        RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel),
                        // Set extended properties
                        OptionSet = booleanOptionSet
                    };

                    // Add to list
                    addedAttributes.Add(boolAttribute);
                }
                // Create a date time attribute
                else if (model.AttributeType.Trim() == "Date and Time")
                {
                    bool isAuditEnable = false;
                    if (model.IsAuditEnabled.Trim() == "Yes")
                        isAuditEnable = true;
                    AttributeRequiredLevel requirementLevel = AttributeRequiredLevel.None;
                    if (model.RequiredLevel.Trim() == "SystemRequired")
                        requirementLevel = AttributeRequiredLevel.SystemRequired;
                    else if (model.RequiredLevel.Trim() == "ApplicationRequired")
                        requirementLevel = AttributeRequiredLevel.ApplicationRequired;
                    DateTimeFormat dateTimeFormat = DateTimeFormat.DateOnly;
                    if (model.DateFormat.Trim() == "DateAndTime")
                        dateTimeFormat = DateTimeFormat.DateAndTime;


                    DateTimeAttributeMetadata dtAttribute = new DateTimeAttributeMetadata
                    {
                        // Set base properties
                        SchemaName = preFix + "_" + model.DisplayName.Replace(" ", ""),
                        LogicalName = preFix + "_" + model.DisplayName.Replace(" ", "").ToLower(),
                        DisplayName = new Label(model.DisplayName, _languageCode),
                        IsAuditEnabled = new BooleanManagedProperty(isAuditEnable),
                        RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel),
                        Format = dateTimeFormat,
                        ImeMode = ImeMode.Disabled
                    };

                    // Add to list
                    addedAttributes.Add(dtAttribute);
                }
                else if (model.AttributeType.Trim() == "Decimal Number")
                {
                    bool isAuditEnable = false;
                    if (model.IsAuditEnabled.Trim() == "Yes")
                        isAuditEnable = true;
                    AttributeRequiredLevel requirementLevel = AttributeRequiredLevel.None;
                    if (model.RequiredLevel.Trim() == "SystemRequired")
                        requirementLevel = AttributeRequiredLevel.SystemRequired;
                    else if (model.RequiredLevel.Trim() == "ApplicationRequired")
                        requirementLevel = AttributeRequiredLevel.ApplicationRequired;

                    // Create a decimal attribute   
                    DecimalAttributeMetadata decimalAttribute = new DecimalAttributeMetadata
                    {
                        // Set base properties
                        SchemaName = preFix + "_" + model.DisplayName.Replace(" ", ""),
                        LogicalName = preFix + "_" + model.DisplayName.Replace(" ", "").ToLower(),
                        DisplayName = new Label(model.DisplayName, _languageCode),
                        IsAuditEnabled = new BooleanManagedProperty(isAuditEnable),
                        RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel),
                        // Set extended properties
                        MaxValue = int.Parse(model.MaxValue),
                        MinValue = int.Parse(model.MinValue),
                        Precision = model.Precision == "" ? 0 : int.Parse(model.Precision)
                    };

                    // Add to list
                    addedAttributes.Add(decimalAttribute);
                }
                else if (model.AttributeType.Trim() == "Two Option")
                {
                    bool isAuditEnable = false;
                    if (model.IsAuditEnabled.Trim() == "Yes")
                        isAuditEnable = true;
                    AttributeRequiredLevel requirementLevel = AttributeRequiredLevel.None;
                    if (model.RequiredLevel.Trim() == "SystemRequired")
                        requirementLevel = AttributeRequiredLevel.SystemRequired;
                    else if (model.RequiredLevel.Trim() == "ApplicationRequired")
                        requirementLevel = AttributeRequiredLevel.ApplicationRequired;
                    IntegerFormat Format = IntegerFormat.Locale;
                    if (model.wholeNumberFormat.Trim() == "Duration")
                        Format = IntegerFormat.Duration;
                    else if (model.wholeNumberFormat.Trim() == "Language")
                        Format = IntegerFormat.Language;
                    else if (model.wholeNumberFormat.Trim() == "Locale")
                        Format = IntegerFormat.Locale;
                    else if (model.wholeNumberFormat.Trim() == "None")
                        Format = IntegerFormat.None;
                    else if (model.wholeNumberFormat.Trim() == "TimeZone")
                        Format = IntegerFormat.TimeZone;


                    // Create a integer attribute   
                    IntegerAttributeMetadata integerAttribute = new IntegerAttributeMetadata
                    {
                        // Set base properties
                        SchemaName = preFix + "_" + model.DisplayName.Replace(" ", ""),
                        LogicalName = preFix + "_" + model.DisplayName.Replace(" ", "").ToLower(),
                        DisplayName = new Label(model.DisplayName, _languageCode),
                        IsAuditEnabled = new BooleanManagedProperty(isAuditEnable),
                        RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel),
                        // Set extended properties
                        Format = Format,
                        MaxValue = int.Parse(model.MaxValue),
                        MinValue = int.Parse(model.MinValue)
                    };

                    // Add to list
                    addedAttributes.Add(integerAttribute);
                }
                else if (model.AttributeType.Trim() == "Multi Line of Text")
                {
                    bool isAuditEnable = false;
                    if (model.IsAuditEnabled.Trim() == "Yes")
                        isAuditEnable = true;
                    AttributeRequiredLevel requirementLevel = AttributeRequiredLevel.None;
                    if (model.RequiredLevel.Trim() == "SystemRequired")
                        requirementLevel = AttributeRequiredLevel.SystemRequired;
                    else if (model.RequiredLevel.Trim() == "ApplicationRequired")
                        requirementLevel = AttributeRequiredLevel.ApplicationRequired;
                    StringFormat Format = StringFormat.TextArea;
                    if (model.StringFormat.Trim() == "TextArea")
                        Format = StringFormat.TextArea;
                    else if (model.StringFormat.Trim() == "Json")
                        Format = StringFormat.Json;
                    else if (model.StringFormat.Trim() == "Phone")
                        Format = StringFormat.Phone;
                    else if (model.StringFormat.Trim() == "PhoneticGuide")
                        Format = StringFormat.PhoneticGuide;
                    else if (model.StringFormat.Trim() == "RichText")
                        Format = StringFormat.RichText;
                    else if (model.StringFormat.Trim() == "Text")
                        Format = StringFormat.Text;
                    else if (model.StringFormat.Trim() == "TickerSymbol")
                        Format = StringFormat.TickerSymbol;
                    else if (model.StringFormat.Trim() == "Url")
                        Format = StringFormat.Url;
                    else if (model.StringFormat.Trim() == "VersionNumber")
                        Format = StringFormat.VersionNumber;

                    // Create a memo attribute 
                    MemoAttributeMetadata memoAttribute = new MemoAttributeMetadata
                    {
                        // Set base properties
                        SchemaName = preFix + "_" + model.DisplayName.Replace(" ", ""),
                        LogicalName = preFix + "_" + model.DisplayName.Replace(" ", "").ToLower(),
                        DisplayName = new Label(model.DisplayName, _languageCode),
                        IsAuditEnabled = new BooleanManagedProperty(isAuditEnable),
                        RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel),
                        // Set extended properties
                        Format = Format,
                        ImeMode = ImeMode.Disabled,
                        MaxLength = int.Parse(model.MaxLength)
                    };

                    // Add to list
                    addedAttributes.Add(memoAttribute);
                }
                else if (model.AttributeType.Trim() == "Currency")
                {

                    bool isAuditEnable = false;
                    if (model.IsAuditEnabled.Trim() == "Yes")
                        isAuditEnable = true;
                    AttributeRequiredLevel requirementLevel = AttributeRequiredLevel.None;
                    if (model.RequiredLevel.Trim() == "SystemRequired")
                        requirementLevel = AttributeRequiredLevel.SystemRequired;
                    else if (model.RequiredLevel.Trim() == "ApplicationRequired")
                        requirementLevel = AttributeRequiredLevel.ApplicationRequired;



                    // Create a money attribute 
                    MoneyAttributeMetadata moneyAttribute = new MoneyAttributeMetadata
                    {
                        // Set base properties
                        SchemaName = preFix + "_" + model.DisplayName.Replace(" ", ""),
                        LogicalName = preFix + "_" + model.DisplayName.Replace(" ", "").ToLower(),
                        DisplayName = new Label(model.DisplayName, _languageCode),
                        IsAuditEnabled = new BooleanManagedProperty(isAuditEnable),
                        RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel),
                        // Set extended properties
                        MaxValue = float.Parse(model.MaxValue),
                        MinValue = float.Parse(model.MinValue),
                        Precision = model.Precision == "" ? 0 : int.Parse(model.Precision),
                        PrecisionSource = 1,
                        ImeMode = ImeMode.Disabled
                    };

                    // Add to list
                    addedAttributes.Add(moneyAttribute);
                }
                else if (model.AttributeType.Trim() == "Optioin Set")
                {
                    OptionMetadata[] optionMetadataList = { };
                    OptionMetadata optionMetadata = new OptionMetadata();


                    OptionSetMetadata optionSetMetadata = new OptionSetMetadata
                    {
                        IsGlobal = true,
                        OptionSetType = OptionSetType.Picklist,
                        Name = preFix + "_" + model.DisplayName.Replace(" ", "").ToLower()
                    };
                    bool isAuditEnable = false;
                    if (model.IsAuditEnabled.Trim() == "Yes")
                        isAuditEnable = true;
                    AttributeRequiredLevel requirementLevel = AttributeRequiredLevel.None;
                    if (model.RequiredLevel.Trim() == "SystemRequired")
                        requirementLevel = AttributeRequiredLevel.SystemRequired;
                    else if (model.RequiredLevel.Trim() == "ApplicationRequired")
                        requirementLevel = AttributeRequiredLevel.ApplicationRequired;


                    // Create a picklist attribute  
                    PicklistAttributeMetadata pickListAttribute = new PicklistAttributeMetadata
                    {
                        // Set base properties
                        SchemaName = preFix + "_" + model.DisplayName.Replace(" ", ""),
                        LogicalName = preFix + "_" + model.DisplayName.Replace(" ", "").ToLower(),
                        DisplayName = new Label(model.DisplayName, _languageCode),
                        IsAuditEnabled = new BooleanManagedProperty(isAuditEnable),
                        RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel),
                        // Set extended properties
                        // Build local picklist options
                        OptionSet = optionSetMetadata
                    };

                    CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                    {
                        EntityName = _entityName,
                        Attribute = pickListAttribute
                    };

                    // Execute the request.
                    service.Execute(createAttributeRequest);

                    #region Create Options in GlobalOptionSet
                    RetrieveOptionSetRequest retrieveOptionSetRequest =
                    new RetrieveOptionSetRequest
                    {
                        Name = preFix + "_" + model.DisplayName.Replace(" ", "").ToLower()
                    };
                    // Execute the request.
                    RetrieveOptionSetResponse retrieveOptionSetResponse = (RetrieveOptionSetResponse)service.Execute(retrieveOptionSetRequest);
                    //Console.WriteLine("Retrieved {0}.", retrieveOptionSetRequest.Name);
                    // Access the retrieved OptionSetMetadata.
                    OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;
                    // Get the current options list for the retrieved attribute.
                    OptionMetadata[] optionList = retrievedOptionSetMetadata.Options.ToArray();
                    if (model.Options.Contains("|"))
                    {
                        string[] optionValue = { };
                        if (model.OptionValue.Contains("|"))
                        {
                            optionValue = model.OptionValue.Split('|');
                        }
                        string[] options = model.Options.Split('|');
                        int i = 0;
                        foreach (string lable in options)
                        {
                            i++;
                            if (Array.Find(optionList, o => o.Value == Convert.ToInt32(options[i])) == null)
                            {
                                InsertOptionValueRequest insertOptionValueRequest =
                                new InsertOptionValueRequest
                                {
                                    OptionSetName = preFix + "_" + model.DisplayName.Replace(" ", "").ToLower(),
                                    Label = new Label(lable, _languageCode),
                                    Value = Convert.ToInt32(Convert.ToInt32(options[i]))
                                };
                                int _insertedOptionValue = ((InsertOptionValueResponse)service.Execute(
                                insertOptionValueRequest)).NewOptionValue;
                            }
                            else
                            {
                                UpdateOptionValueRequest updateOptionValueRequest = new UpdateOptionValueRequest
                                {
                                    OptionSetName = preFix + "_" + model.DisplayName.Replace(" ", "").ToLower(),
                                    Label = new Label(lable, _languageCode),
                                    Value = Convert.ToInt32(Convert.ToInt32(options[i]))
                                };
                                service.Execute(updateOptionValueRequest);

                            }
                            //Publish the OptionSet
                            PublishXmlRequest pxReq2 = new PublishXmlRequest { ParameterXml = String.Format("<importexportxml><optionsets><optionset>{0}</optionset></optionsets></importexportxml>", preFix + "_" + model.DisplayName.Replace(" ", "").ToLower()) };
                            service.Execute(pxReq2);

                        }

                    }

                    #endregion
                }
                else if (model.AttributeType.Trim() == "Single Line of Text")
                {
                    bool isAuditEnable = false;
                    if (model.IsAuditEnabled.Trim() == "Yes")
                        isAuditEnable = true;
                    AttributeRequiredLevel requirementLevel = AttributeRequiredLevel.None;
                    if (model.RequiredLevel.Trim() == "SystemRequired")
                        requirementLevel = AttributeRequiredLevel.SystemRequired;
                    else if (model.RequiredLevel.Trim() == "ApplicationRequired")
                        requirementLevel = AttributeRequiredLevel.ApplicationRequired;


                    // Create a string attribute
                    StringAttributeMetadata stringAttribute = new StringAttributeMetadata
                    {
                        // Set base properties
                        SchemaName = preFix + "_" + model.DisplayName.Replace(" ", ""),
                        LogicalName = preFix + "_" + model.DisplayName.Replace(" ", "").ToLower(),
                        DisplayName = new Label(model.DisplayName, _languageCode),
                        IsAuditEnabled = new BooleanManagedProperty(isAuditEnable),
                        RequiredLevel = new AttributeRequiredLevelManagedProperty(requirementLevel),
                        // Set extended properties
                        MaxLength = int.Parse(model.MaxLength)
                    };

                    // Add to list
                    addedAttributes.Add(stringAttribute);
                }


                //bool eligibleCreateOneToManyRelationship = EligibleCreateOneToManyRelationship("account", "campaign");

                else if (model.AttributeType.Trim() == "Lookup")
                {
                    CreateOneToManyRequest createOneToManyRelationshipRequest = new CreateOneToManyRequest
                    {
                        OneToManyRelationship = new OneToManyRelationshipMetadata
                        {
                            ReferencedEntity = model.EntityName,
                            ReferencingEntity = _entityName,
                            SchemaName = preFix + "_" + model.EntityName + "_" + _entityName,
                            AssociatedMenuConfiguration = new AssociatedMenuConfiguration
                            {
                                Behavior = AssociatedMenuBehavior.UseCollectionName,
                                Group = AssociatedMenuGroup.Details,
                                Label = new Label(model.DisplayName, 1033),
                                Order = 10000
                            },
                            CascadeConfiguration = new CascadeConfiguration
                            {
                                Assign = CascadeType.NoCascade,
                                Delete = CascadeType.RemoveLink,
                                Merge = CascadeType.NoCascade,
                                Reparent = CascadeType.NoCascade,
                                Share = CascadeType.NoCascade,
                                Unshare = CascadeType.NoCascade
                            }
                        },
                        Lookup = new LookupAttributeMetadata
                        {
                            SchemaName = preFix + "_" + model.DisplayName.Replace(" ", ""),
                            DisplayName = new Label(model.DisplayName, _languageCode),
                            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),

                        }
                    };


                    CreateOneToManyResponse createOneToManyRelationshipResponse = (CreateOneToManyResponse)service.Execute(createOneToManyRelationshipRequest);

                    var _oneToManyRelationshipId = createOneToManyRelationshipResponse.RelationshipId;
                    var _oneToManyRelationshipName = createOneToManyRelationshipRequest.OneToManyRelationship.SchemaName;
                }

            }



            // NOTE: LookupAttributeMetadata cannot be created outside the context of a relationship.
            // Refer to the WorkWithRelationships.cs reference SDK sample for an example of this attribute type.

            // NOTE: StateAttributeMetadata and StatusAttributeMetadata cannot be created via the SDK.

            foreach (AttributeMetadata anAttribute in addedAttributes)
            {
                // Create the request.
                CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                {
                    EntityName = _entityName,
                    Attribute = anAttribute
                };

                // Execute the request.
                service.Execute(createAttributeRequest);

                Console.WriteLine(string.Format("Created the attribute {0}.", anAttribute.SchemaName));
            }
        }
        public static void getdataFromExcel(IOrganizationService service)
        {
            string excelPath = @"C:\Users\35405\OneDrive - Havells\Desktop\Havells\Fields\ListOfFields.xlsx";
            Application excelApp = new Application();
            if (excelApp != null)
            {
                Workbook excelWorkbook = excelApp.Workbooks.Open(excelPath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet excelWorksheet = (Worksheet)excelWorkbook.Sheets[1];
                string _entityName = excelWorksheet.Name;
                Range excelRange = excelWorksheet.UsedRange;
                int rowCount = excelRange.Rows.Count;
                int iLeft = excelRange.Rows.Count;
                int colCount = excelRange.Columns.Count;
                int iDone = 0;
                List<Models> _ExcelAttributeList = new List<Models>();
                Models models = new Models();
                for (int i = 2; i <= rowCount; i++)  //rowCount
                {
                    try
                    {
                        models = new Models();
                        models.DisplayName = (excelWorksheet.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value == null ? "" : (excelWorksheet.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                        models.AttributeType = (excelWorksheet.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value == null ? "" : (excelWorksheet.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                        models.MaxLength = (excelWorksheet.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Value == null ? "" : (excelWorksheet.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                        models.MinValue = (excelWorksheet.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Value == null ? "" : (excelWorksheet.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                        models.MaxValue = (excelWorksheet.Cells[i, 5] as Microsoft.Office.Interop.Excel.Range).Value == null ? "" : (excelWorksheet.Cells[i, 5] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                        models.Precision = (excelWorksheet.Cells[i, 6] as Microsoft.Office.Interop.Excel.Range).Value == null ? "" : (excelWorksheet.Cells[i, 6] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                        models.RequiredLevel = (excelWorksheet.Cells[i, 7] as Microsoft.Office.Interop.Excel.Range).Value == null ? "" : (excelWorksheet.Cells[i, 7] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                        models.IsAuditEnabled = (excelWorksheet.Cells[i, 8] as Microsoft.Office.Interop.Excel.Range).Value == null ? "" : (excelWorksheet.Cells[i, 8] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                        models.EntityName = (excelWorksheet.Cells[i, 9] as Microsoft.Office.Interop.Excel.Range).Value == null ? "" : (excelWorksheet.Cells[i, 9] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                        models.Options = (excelWorksheet.Cells[i, 10] as Microsoft.Office.Interop.Excel.Range).Value == null ? "" : (excelWorksheet.Cells[i, 10] as Microsoft.Office.Interop.Excel.Range).Value.ToString();

                        models.OptionValue = (excelWorksheet.Cells[i, 11] as Microsoft.Office.Interop.Excel.Range).Value == null ? "" : (excelWorksheet.Cells[i, 11] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                        models.DateFormat = (excelWorksheet.Cells[i, 12] as Microsoft.Office.Interop.Excel.Range).Value == null ? "" : (excelWorksheet.Cells[i, 12] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                        models.wholeNumberFormat = (excelWorksheet.Cells[i, 13] as Microsoft.Office.Interop.Excel.Range).Value == null ? "" : (excelWorksheet.Cells[i, 13] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
                        models.StringFormat = (excelWorksheet.Cells[i, 14] as Microsoft.Office.Interop.Excel.Range).Value == null ? "" : (excelWorksheet.Cells[i, 14] as Microsoft.Office.Interop.Excel.Range).Value.ToString();

                        _ExcelAttributeList.Add(models);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("-------------------------eee " + ex.Message);
                    }
                }
                excelWorkbook.Close();
                Console.WriteLine("Total row completed " + iDone);
                //Console.ReadLine();
                excelApp.Quit();
                CreateField(service, _entityName, _ExcelAttributeList);

            }
        }
        public static void retriveMetedata(IOrganizationService service)
        {
            List<AttributeMetadata> addedAttributes = new List<AttributeMetadata>();
            string entityname = "hil_performainvoicearchived";
            RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.All,
                LogicalName = "msdyn_workorder"
            };
            RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
            EntityMetadata AccountEntity = retrieveAccountEntityResponse.EntityMetadata;

            Console.WriteLine("Account entity metadata:");
            Console.WriteLine(AccountEntity.SchemaName);
            Console.WriteLine(AccountEntity.DisplayName.UserLocalizedLabel.Label);
            Console.WriteLine(AccountEntity.EntityColor);

            Console.WriteLine("Account entity attributes:");
            int i = 1;
            #region Relationship
            //foreach (object relationShip in AccountEntity.ManyToOneRelationships)
            //{
            //    try
            //    {
            //        OneToManyRelationshipMetadata relation = (OneToManyRelationshipMetadata)relationShip;
            //        relation.MetadataId = new Guid();
            //        relation.ReferencedEntity = relation.ReferencedEntity == "owner" ? "systemuser" : relation.ReferencedEntity;
            //        AttributeMetadata[] a = AccountEntity.Attributes;
            //        AttributeMetadata lookField = Array.Find(a, ele => ele.SchemaName.ToLower() == relation.ReferencingAttribute.ToLower());
            //      Console.WriteLine(AccountEntity.EntityColor);
            //        relation.SchemaName = relation.SchemaName.Replace("msdyn_workorder", entityname);
            //        CreateOneToManyRequest createOneToManyRelationshipRequest = new CreateOneToManyRequest
            //        {
            //            OneToManyRelationship = new OneToManyRelationshipMetadata
            //            {
            //                ReferencedEntity = relation.ReferencedEntity,
            //                ReferencingEntity = entityname,
            //                SchemaName = relation.SchemaName,
            //                AssociatedMenuConfiguration = relation.AssociatedMenuConfiguration,
            //                CascadeConfiguration = relation.CascadeConfiguration,
            //            },
            //            Lookup = new LookupAttributeMetadata
            //            {
            //                SchemaName = lookField.SchemaName.Contains("_") ? lookField.SchemaName : "hil_" + lookField.SchemaName,
            //                LogicalName = lookField.LogicalName.Contains("_") ? lookField.LogicalName : "hil_" + lookField.LogicalName,
            //                DisplayName = lookField.DisplayName,
            //                RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),

            //            }
            //        };


            //        CreateOneToManyResponse createOneToManyRelationshipResponse = (CreateOneToManyResponse)service.Execute(createOneToManyRelationshipRequest);

            //        var _oneToManyRelationshipId = createOneToManyRelationshipResponse.RelationshipId;
            //        var _oneToManyRelationshipName = createOneToManyRelationshipRequest.OneToManyRelationship.SchemaName;

            //      Console.WriteLine(
            //            "The one-to-many relationship has been created between {0} and {1}.",
            //            "account", "campaign");
            //    }
            //    catch (Exception ex)
            //    {
            //        if(!ex.Message.Contains("type EntityRelationship is not unique."))
            //        {
            //          Console.WriteLine("Ex " + ex.Message);
            //        }
            //      Console.WriteLine("Ex " + ex.Message);
            //    }

            //}
            #endregion
            Console.WriteLine("Entity Name/firld Name");
            foreach (AttributeMetadata attr in AccountEntity.Attributes)
            {
                //AttributeMetadata a = (AttributeMetadata)attr;
                if (attr.AttributeType == AttributeTypeCode.Lookup)
                {
                    //Console.WriteLine("lookup");
                    Console.WriteLine(((Microsoft.Xrm.Sdk.Metadata.LookupAttributeMetadata)attr).Targets[0] + "/" + attr.LogicalName);
                }
            }
            Console.WriteLine("Count ");
            Parallel.ForEach(AccountEntity.Attributes, attribute =>
            {
                try
                {
                    AttributeMetadata a = (AttributeMetadata)attribute;
                    Console.WriteLine("Count " + i + " || " + a.LogicalName);
                    if (a.AttributeType != AttributeTypeCode.Virtual)
                    {
                        if (a.AttributeType == AttributeTypeCode.String)
                        {
                            if (a.DisplayName.LocalizedLabels.Count != 0)
                            {

                                // Create a string attribute
                                StringAttributeMetadata stringAttribute = new StringAttributeMetadata
                                {
                                    // Set base properties
                                    SchemaName = a.SchemaName.Contains("_") ? a.SchemaName : "hil_" + a.SchemaName,
                                    LogicalName = a.LogicalName.Contains("_") ? a.LogicalName : "hil_" + a.LogicalName,
                                    DisplayName = a.DisplayName,//new Label(model.DisplayName, _languageCode),
                                                                // IsAuditEnabled = new BooleanManagedProperty(a.IsAuditEnabled.Value),
                                                                // Set extended properties
                                    MaxLength = ((StringAttributeMetadata)a).MaxLength
                                };

                                CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                                {
                                    EntityName = entityname,
                                    Attribute = stringAttribute
                                };

                                // Execute the request.
                                service.Execute(createAttributeRequest);
                                // Add to list
                                addedAttributes.Add(stringAttribute);
                            }

                        }
                        else if (a.AttributeType == AttributeTypeCode.Money)
                        {
                            // Create a decimal attribute   
                            DoubleAttributeMetadata decimalAttribute = new DoubleAttributeMetadata
                            {
                                // Set base properties
                                SchemaName = a.SchemaName,
                                LogicalName = a.LogicalName,
                                DisplayName = a.DisplayName,
                                //MaxValue = ((MoneyAttributeMetadata)a).MaxValue,
                                //MinValue = ((MoneyAttributeMetadata)a).MinValue,
                                //Precision = ((MoneyAttributeMetadata)a).Precision
                            };

                            CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                            {
                                EntityName = entityname,
                                Attribute = decimalAttribute
                            };

                            // Execute the request.
                            service.Execute(createAttributeRequest);


                        }
                        else if (a.AttributeType == AttributeTypeCode.Double)
                        {
                            // Create a decimal attribute   
                            DoubleAttributeMetadata decimalAttribute = new DoubleAttributeMetadata
                            {
                                // Set base properties
                                SchemaName = a.SchemaName,
                                LogicalName = a.LogicalName,
                                DisplayName = a.DisplayName,
                                MaxValue = ((DoubleAttributeMetadata)a).MaxValue,
                                MinValue = ((DoubleAttributeMetadata)a).MinValue,
                                Precision = ((DoubleAttributeMetadata)a).Precision
                            };

                            CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                            {
                                EntityName = entityname,
                                Attribute = decimalAttribute
                            };

                            // Execute the request.
                            service.Execute(createAttributeRequest);

                        }
                        else if (a.AttributeType == AttributeTypeCode.Decimal)
                        {
                            // Create a decimal attribute   
                            DecimalAttributeMetadata decimalAttribute = new DecimalAttributeMetadata
                            {
                                // Set base properties
                                SchemaName = a.SchemaName.Contains("_") ? a.SchemaName : "hil_" + a.SchemaName,
                                LogicalName = a.LogicalName.Contains("_") ? a.LogicalName : "hil_" + a.LogicalName,
                                DisplayName = a.DisplayName,
                                MaxValue = ((DecimalAttributeMetadata)a).MaxValue,
                                MinValue = ((DecimalAttributeMetadata)a).MinValue
                            };

                            CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                            {
                                EntityName = entityname,
                                Attribute = decimalAttribute
                            };

                            // Execute the request.
                            service.Execute(createAttributeRequest);
                        }
                        else if (a.AttributeType == AttributeTypeCode.Integer)
                        {
                            // Create a decimal attribute   
                            IntegerAttributeMetadata decimalAttribute = new IntegerAttributeMetadata
                            {
                                // Set base properties
                                SchemaName = a.SchemaName.Contains("_") ? a.SchemaName : "hil_" + a.SchemaName,
                                LogicalName = a.LogicalName.Contains("_") ? a.LogicalName : "hil_" + a.LogicalName,
                                DisplayName = a.DisplayName,
                                MaxValue = ((IntegerAttributeMetadata)a).MaxValue,
                                MinValue = ((IntegerAttributeMetadata)a).MinValue
                            };

                            CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                            {
                                EntityName = entityname,
                                Attribute = decimalAttribute
                            };

                            // Execute the request.
                            service.Execute(createAttributeRequest);
                        }
                        else if (a.AttributeType == AttributeTypeCode.BigInt)
                        {
                            // Create a decimal attribute   
                            BigIntAttributeMetadata decimalAttribute = new BigIntAttributeMetadata
                            {
                                // Set base properties
                                SchemaName = a.SchemaName.Contains("_") ? a.SchemaName : "hil_" + a.SchemaName,
                                LogicalName = a.LogicalName.Contains("_") ? a.LogicalName : "hil_" + a.LogicalName,
                                DisplayName = a.DisplayName
                            };

                            CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                            {
                                EntityName = entityname,
                                Attribute = decimalAttribute
                            };
                            if (a.SchemaName != "EntityImage_Timestamp")
                                // Execute the request.
                                service.Execute(createAttributeRequest);
                        }
                        else if (a.AttributeType == AttributeTypeCode.Picklist)
                        {
                            OptionMetadata[] optionMetadataList = { };
                            OptionMetadata optionMetadata = new OptionMetadata();


                            OptionSetMetadata optionSetMetadata = ((Microsoft.Xrm.Sdk.Metadata.EnumAttributeMetadata)a).OptionSet;
                            // Create a picklist attribute  
                            PicklistAttributeMetadata pickListAttribute = new PicklistAttributeMetadata
                            {
                                // Set base properties
                                SchemaName = a.SchemaName.Contains("_") ? a.SchemaName : "hil_" + a.SchemaName,
                                LogicalName = a.LogicalName.Contains("_") ? a.LogicalName : "hil_" + a.LogicalName,
                                DisplayName = a.DisplayName,
                                // Build local picklist options
                                OptionSet = optionSetMetadata
                            };
                            //EnumAttributeMetadata optionSetMetadata = (EnumAttributeMetadata)a;
                            //optionSetMetadata.MetadataId = new Guid();    /* This is to avoid GUID collision */

                            if (optionSetMetadata.IsGlobal == false) /* If it is NOT Global Option Set */
                            {
                                optionSetMetadata.MetadataId = new Guid();   /* This is to avoid GUID collision */
                                optionSetMetadata.Name = entityname + optionSetMetadata.Name.Substring("msdyn_workorder".Length); /* Replace Parent LogicalName with Child LogicalName */
                            }
                            else
                            {
                                Console.WriteLine("Global Option Set Name ---> " + optionSetMetadata.Name);
                                optionSetMetadata.Options.Clear();
                            }
                            //optionSetMetadata= false;
                            CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                            {
                                EntityName = entityname,
                                Attribute = pickListAttribute
                            };

                            // Execute the request.
                            service.Execute(createAttributeRequest);
                        }
                        else if (a.AttributeType == AttributeTypeCode.Boolean)
                        {
                            BooleanOptionSetMetadata booleanOptionSet = new BooleanOptionSetMetadata();
                            BooleanOptionSetMetadata booleanOptionSetValue = new BooleanOptionSetMetadata();


                            // Create a boolean attribute
                            BooleanAttributeMetadata boolAttribute = new BooleanAttributeMetadata
                            {
                                SchemaName = a.SchemaName,
                                LogicalName = a.LogicalName,
                                DisplayName = a.DisplayName,
                                // Set extended properties
                                OptionSet = ((Microsoft.Xrm.Sdk.Metadata.BooleanAttributeMetadata)a).OptionSet
                            };

                            if (boolAttribute.OptionSet.IsGlobal == false) /* If it is NOT Global Option Set */
                            {
                                boolAttribute.OptionSet.MetadataId = new Guid();   /* This is to avoid GUID collision */
                                boolAttribute.OptionSet.Name = entityname + boolAttribute.OptionSet.Name.Substring("msdyn_workorder".Length); /* Replace Parent LogicalName with Child LogicalName */
                            }
                            else
                            {
                                Console.WriteLine("Global Option Set Name ---> " + boolAttribute.OptionSet.Name);

                            }


                            //BooleanAttributeMetadata optionSetMetadata = (BooleanAttributeMetadata)a;
                            //optionSetMetadata.MetadataId = new Guid();    /* This is to avoid GUID collision */

                            //if (optionSetMetadata.OptionSet.IsGlobal == false) /* If it is NOT Global Option Set */
                            //{
                            //    optionSetMetadata.OptionSet.MetadataId = new Guid();   /* This is to avoid GUID collision */
                            //    optionSetMetadata.OptionSet.Name = entityname + optionSetMetadata.OptionSet.Name.Substring("msdyn_workorder".Length); /* Replace Parent LogicalName with Child LogicalName */
                            //}
                            //else
                            //{
                            //  Console.WriteLine("Global Option Set Name ---> " + optionSetMetadata.OptionSet.Name);

                            //}
                            boolAttribute.IsSecured = false;
                            CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                            {
                                EntityName = entityname,
                                Attribute = boolAttribute
                            };

                            // Execute the request.
                            service.Execute(createAttributeRequest);
                            // Add to list
                            addedAttributes.Add(boolAttribute);
                        }
                        else if (a.AttributeType == AttributeTypeCode.Memo)
                        {
                            MemoAttributeMetadata memoAttribute = new MemoAttributeMetadata
                            {
                                // Set base properties
                                SchemaName = a.SchemaName,
                                LogicalName = a.LogicalName,
                                DisplayName = a.DisplayName,
                                // Set extended properties
                                Format = ((Microsoft.Xrm.Sdk.Metadata.MemoAttributeMetadata)a).Format,
                                ImeMode = ImeMode.Disabled,
                                MaxLength = ((Microsoft.Xrm.Sdk.Metadata.MemoAttributeMetadata)a).MaxLength
                            };
                            CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                            {
                                EntityName = entityname,
                                Attribute = memoAttribute
                            };

                            // Execute the request.
                            service.Execute(createAttributeRequest);
                            // Add to list
                            addedAttributes.Add(memoAttribute);
                        }
                        else if (a.AttributeType == AttributeTypeCode.DateTime)
                        {
                            DateTimeAttributeMetadata dtAttribute = new DateTimeAttributeMetadata
                            {
                                // Set base properties
                                SchemaName = a.SchemaName.Contains("_") ? a.SchemaName : "hil_" + a.SchemaName,
                                LogicalName = a.LogicalName.Contains("_") ? a.LogicalName : "hil_" + a.LogicalName,
                                DisplayName = a.DisplayName,
                                Format = ((Microsoft.Xrm.Sdk.Metadata.DateTimeAttributeMetadata)a).Format,
                                ImeMode = ImeMode.Disabled
                            };
                            CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
                            {
                                EntityName = entityname,
                                Attribute = dtAttribute
                            };

                            // Execute the request.
                            service.Execute(createAttributeRequest);
                            // Add to list
                            addedAttributes.Add(dtAttribute);
                        }
                        else if (a.AttributeType == AttributeTypeCode.Uniqueidentifier)
                        {

                        }
                        else if (a.AttributeType == AttributeTypeCode.Customer)
                        {
                            Console.WriteLine("customer");
                        }
                        else if (a.AttributeType == AttributeTypeCode.Lookup)
                        {
                            Console.WriteLine("lookup");
                        }
                        else if (a.AttributeType == AttributeTypeCode.Owner)
                        {
                            Console.WriteLine("owner");
                        }
                        else
                        {
                            Console.WriteLine("dd");
                        }
                    }

                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("already exists for entity"))
                    {
                        Console.WriteLine("Error " + ex.Message);
                    }
                    else
                    {
                        Console.WriteLine("-------------------------Error " + ex.Message);
                    }
                }
            });

            Console.WriteLine("Count ");
            #region
            //foreach (object relationShip in AccountEntity.ManyToOneRelationships)
            //{
            //    OneToManyRelationshipMetadata relation = (OneToManyRelationshipMetadata)relationShip;
            //    relation.MetadataId = new Guid();

            //    AttributeMetadata[] a = AccountEntity.Attributes;
            //    AttributeMetadata lookField = Array.Find(a, ele => ele.SchemaName.ToLower() == relation.ReferencingAttribute.ToLower());
            //  Console.WriteLine(AccountEntity.EntityColor);
            //    CreateOneToManyRequest createOneToManyRelationshipRequest = new CreateOneToManyRequest
            //    {
            //        OneToManyRelationship = new OneToManyRelationshipMetadata
            //        {
            //            ReferencedEntity = relation.ReferencedEntity,
            //            ReferencingEntity = entityname,
            //            SchemaName = relation.SchemaName,
            //            AssociatedMenuConfiguration = new AssociatedMenuConfiguration
            //            {
            //                Behavior = AssociatedMenuBehavior.UseCollectionName,
            //                Group = AssociatedMenuGroup.Details,
            //                // Label = new Label(relation.DisplayName, 1033),
            //                Order = 10000
            //            },
            //            CascadeConfiguration = new CascadeConfiguration
            //            {
            //                Assign = CascadeType.NoCascade,
            //                Delete = CascadeType.RemoveLink,
            //                Merge = CascadeType.NoCascade,
            //                Reparent = CascadeType.NoCascade,
            //                Share = CascadeType.NoCascade,
            //                Unshare = CascadeType.NoCascade
            //            }
            //        },
            //        Lookup = new LookupAttributeMetadata
            //        {
            //            SchemaName = lookField.SchemaName.Contains("_") ? lookField.SchemaName : "hil_" + lookField.SchemaName,
            //            LogicalName = lookField.LogicalName.Contains("_") ? lookField.LogicalName : "hil_" + lookField.LogicalName,
            //            DisplayName = lookField.DisplayName,
            //            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),

            //        }
            //    };


            //    CreateOneToManyResponse createOneToManyRelationshipResponse = (CreateOneToManyResponse)service.Execute(createOneToManyRelationshipRequest);

            //    var _oneToManyRelationshipId = createOneToManyRelationshipResponse.RelationshipId;
            //    var _oneToManyRelationshipName = createOneToManyRelationshipRequest.OneToManyRelationship.SchemaName;

            //  Console.WriteLine(
            //        "The one-to-many relationship has been created between {0} and {1}.",
            //        "account", "campaign");
            //}
            #endregion
            Console.WriteLine("****************************************************************************************");
            //foreach (AttributeMetadata anAttribute in addedAttributes)
            //{
            //    // Create the request.
            //    CreateAttributeRequest createAttributeRequest = new CreateAttributeRequest
            //    {
            //        EntityName = entityname,
            //        Attribute = anAttribute
            //    };

            //    // Execute the request.
            //    service.Execute(createAttributeRequest);

            //  Console.WriteLine("Created the attribute {0}.", anAttribute.SchemaName);
            //}
        }
    }
}
