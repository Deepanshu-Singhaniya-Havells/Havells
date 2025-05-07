using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CreateFieldsFromExcelToD365
{
    public class MigrateKeys
    {
       public static void mainFunction(IOrganizationService service, IOrganizationService serviceSer)
        {
            string EntityLists = "account;appointment;bookableresourcebooking;campaign;characteristic;contact;email;incident;lead;msdyn_customerasset;msdyn_incidenttype;msdyn_incidenttypeproduct;msdyn_resourcerequirement;msdyn_timeoffrequest;msdyn_workorder;msdyn_workorderincident;msdyn_workorderproduct;msdyn_workorderservice;msdyn_workordersubstatus;phonecall;product;systemuser;task;hil_address;hil_advancefind;dyn_advancefindconfig;hil_advisormaster;hil_advisoryenquiry;hil_homeadvisoryline;hil_alerts;hil_alternatepart;hil_amcdiscountmatrix;hil_amcplan;hil_amcstaging;hil_approvalmatrix;hil_approval;hil_aquaparameters;hil_area;hil_assignmentmatrix;hil_assignmentmatrixupload;hil_attachment;hil_tenderattachmentdoctype;hil_attachmentdocumenttype;hil_tenderattachmentmanager;hil_auditlog;ispl_autonumberconfiguration;hil_autonumber;hil_tenderbankguarantee;hil_jobbasecharge;hil_bdtatdetails;hil_bdtatownership;hil_bdteam;hil_bdteammember;hil_bizgeomappingstagigng;hil_bomcategory;hil_bomsubcategory;hil_branch;hil_bulkjobsuploader;hil_businessmapping;hil_callsubtype;hil_calltype;hil_campaignwebsitesetup;hil_campaigndivisions;hil_campaignenquirytypes;hil_cancellationreason;hil_casecategory;hil_channelpartnercountryclassification;hil_city;hil_claimallocator;hil_claimcategory;hil_claimline;hil_claimlines;hil_claimoverheadline;hil_claimperiod;hil_claimpostingsetup;hil_claimsummary;hil_claimtype;hil_constructiontype;hil_consumerappbridge;hil_consumercategory;hil_consumernpssurvey;hil_consumertype;hil_country;hil_customerassetstagging;hil_remarks;hil_customerfeedback;hil_customerwishlist;hil_healthindicatorheader;hil_technicianhealthindicator;hil_dealerstockverificationheader;hil_materialreturn;hil_deliveryschedule;hil_deliveryschedulemaster;hil_despatchteammaster;hil_discountmatrix;hil_distributionchannel;hil_district;hil_effeciency;hil_enclosuretype;hil_enquirybusinessprocessflow;hil_enquirydepartment;hil_enquirydocumenttype;hil_enquirylostreason;hil_enquiryproductsegment;hil_enquerysegment;hil_enquirysegmentdcmapping;hil_enquirytype;hil_entityfieldsmetadata;hil_errorcode;hil_escalation;hil_escalationmatrix;hil_estimate;hil_feedback;hil_fixedcompensation;hil_forgotpassword;hil_frequency;hil_grnline;hil_hsncode;hil_icauploader;plt_idgenerator;hil_incentivetable;hil_industrysubtype;hil_industrytype;hil_inspectiontype;hil_installationchecklist;hil_integrationconfiguration;hil_integrationjob;hil_integrationjobrun;hil_integrationjobrundetail;new_integrationstaging;hil_staging;hil_integrationtrace;hil_inventory;hil_inventoryjournal;hil_inventoryrequest;hil_invoice;hil_jobcancellationrequest;hil_joberrorcode;hil_jobestimation;hil_jobreassignreason;hil_jobtat;hil_jobtatcategory;hil_jobsextension;hil_travelclosureincentive;hil_labor;hil_leadproduct;hil_materialgroup;hil_materialgroup2;hil_materialgroup3;hil_materialgroup4;hil_materialgroup5;hil_minimumguarantee;hil_minimumstocklevel;hil_mobileappbanner;hil_natureofcomplaint;hil_npssetup;hil_oaheader;hil_oaproduct;hil_oareadinessdatestaging;hil_observation;hil_orderchecklist;hil_orderchecklistproduct;hil_part;hil_partnerdepartment;hil_partnerdivisionmapping;hil_paymentstatus;hil_claimheader;hil_pincode;hil_plantmaster;hil_plantordertypesetup;hil_pmsconfiguration;hil_pmsconfigurationlines;hil_pmsjobuploader;hil_pmsscheduleconfiguration;hil_postatustracker;hil_postatustrackerupload;hil_politicalmapping;hil_postorder;hil_preferredlanguageforcommunication;hil_productcatalog;hil_productrequest;hil_productrequestheader;hil_productstaging;hil_stagingdivisonmaterialgroupmapping;hil_productvideos;hil_production;hil_propertytype;hil_refreshjobs;hil_region;hil_returnheader;hil_returnline;hil_salesoffice;hil_salesofficebranchheadmapping;hil_salesofficeplantmapping;hil_salestatmaster;hil_sawactivity;hil_sawactivityapproval;hil_serviceactionwork;hil_sawcategoryapprovals;hil_serviceactionworksetup;hil_sbubranchmapping;hil_schemedistrictexclusion;hil_schemeincentive;hil_schemeline;hil_securityroleextension;hil_serialnumber;hil_servicebom;hil_servicecallrequest;hil_servicecallrequestdetail;hil_serviceengineergeocode;hil_smsconfiguration;hil_smstemplates;hil_solarservey;hil_specialincentive;hil_stagingintegrationjson;hil_stagingpricingmapping;hil_stagingsbudivisionmapping;hil_startingmethod;hil_state;hil_statustransitionmatrix;hil_subterritory;hil_tatachievementslabmaster;hil_tatbreachpenalty;hil_tatincentive;hil_technicalspecfication;hil_technician;hil_tender;hil_tenderbomlineitem;hil_designteambranchmapping;hil_tenderdocs;hil_tenderemailersetup;hil_tenderpaymentdetail;hil_tenderproduct;hil_userbranchmapping;hil_tolerencesetup;hil_travelexpense;hil_typeofcustomer;hil_typeofproduct;hil_unitwarranty;hil_upcountrydataupdate;hil_upcountrytravelcharge;hil_usageindicator;hil_usersecurityroleextension;hil_usersimeibinding;hil_voltage;hil_warrantyperiod;hil_warrantyscheme;hil_warrantytemplate;hil_warrantyvoidreason;hil_whatsappproductdivisionconfig;hil_workorderarch;hil_wrongclosurepenalty;hil_yammersettings;hil_archivedactivity;hil_archvedjobdatasource;hil_productrequestheaderarchived;new_userlocations\r\n";

            string[] entities = EntityLists.Split(';');
            foreach (string entityName in entities)
            {
                try
                {
                    RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
                    {
                        EntityFilters = EntityFilters.All,
                        LogicalName = entityName
                    };
                    RetrieveEntityResponse retrieveAccountEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveEntityRequest);
                    EntityMetadata StateEntity = retrieveAccountEntityResponse.EntityMetadata;

                    RetrieveEntityRequest retrieveEntityRequestDemo = new RetrieveEntityRequest
                    {
                        EntityFilters = EntityFilters.Entity,
                        LogicalName = entityName
                    };
                    RetrieveEntityResponse retrieveAccountEntityResponseDemo = (RetrieveEntityResponse)serviceSer.Execute(retrieveEntityRequestDemo);
                    EntityMetadata StateEntityDemo = retrieveAccountEntityResponseDemo.EntityMetadata;
                    Console.WriteLine("Key Count" + StateEntity.Keys.Length);
                    foreach (dynamic key in StateEntity.Keys)
                    {
                        EntityKeyMetadata entityKeyMetadata = key;
                        CreateEntityKeyRequest createEntityKey = new CreateEntityKeyRequest() 
                        {
                            EntityName = entityName,
                            EntityKey = entityKeyMetadata,

                        };
                        try
                        {

                            serviceSer.Execute(createEntityKey);
                            Console.WriteLine("Done " + entityName);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error in " + entityName + " || " + ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error in "+ entityName+" || "+ex.Message);
                }

            }
        }
    }
}
