using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateFieldsFromExcelToD365
{
    class ViewsModificationNew
    {
        public static void mainFunction(IOrganizationService service, IOrganizationService serviceSer)
        {
            string entitylist = "account;appointment;bookableresourcebooking;campaign;characteristic;contact;email;incident;lead;msdyn_customerasset;msdyn_incidenttype;msdyn_incidenttypeproduct;msdyn_resourcerequirement;msdyn_timeoffrequest;msdyn_workorder;msdyn_workorderincident;msdyn_workorderproduct;msdyn_workorderservice;msdyn_workordersubstatus;phonecall;product;systemuser;task;hil_address;dyn_advancefindconfig;hil_advancefind;hil_advisormaster;hil_advisoryenquiry;hil_alerts;hil_alternatepart;hil_amcdiscountmatrix;hil_amcplan;hil_amcstaging;hil_approval;hil_approvalmatrix;hil_aquaparameters;hil_archvedjobdatasource;hil_area;hil_assignmentmatrix;hil_assignmentmatrixupload;hil_attachment;hil_attachmentdocumenttype;hil_auditlog;hil_autonumber;hil_bdtatdetails;hil_bdtatownership;hil_bdteam;hil_bdteammember;hil_bizgeomappingstagigng;hil_bomcategory;hil_bomsubcategory;hil_branch;hil_bulkjobsuploader;hil_businessmapping;hil_callsubtype;hil_calltype;hil_campaigndivisions;hil_campaignenquirytypes;hil_campaignwebsitesetup;hil_cancellationreason;hil_casecategory;hil_channelpartnercountryclassification;hil_city;hil_claimallocator;hil_claimcategory;hil_claimheader;hil_claimline;hil_claimlines;hil_claimoverheadline;hil_claimperiod;hil_claimpostingsetup;hil_claimsummary;hil_claimtype;hil_constructiontype;hil_consumerappbridge;hil_consumercategory;hil_consumernpssurvey;hil_consumertype;hil_country;hil_customerassetstagging;hil_customerfeedback;hil_customerwishlist;hil_dealerstockverificationheader;hil_deliveryschedule;hil_deliveryschedulemaster;hil_designteambranchmapping;hil_despatchteammaster;hil_discountmatrix;hil_distributionchannel;hil_district;hil_effeciency;hil_enclosuretype;hil_enquerysegment;hil_enquirybusinessprocessflow;hil_enquirydepartment;hil_enquirydocumenttype;hil_enquirylostreason;hil_enquiryproductsegment;hil_enquirysegmentdcmapping;hil_enquirytype;hil_entityfieldsmetadata;hil_errorcode;hil_escalation;hil_escalationmatrix;hil_estimate;hil_feedback;hil_fixedcompensation;hil_forgotpassword;hil_frequency;hil_grnline;hil_healthindicatorheader;hil_homeadvisoryline;hil_hsncode;hil_icauploader;hil_incentivetable;hil_industrysubtype;hil_industrytype;hil_inspectiontype;hil_installationchecklist;hil_integrationconfiguration;hil_integrationjob;hil_integrationjobrun;hil_integrationjobrundetail;hil_integrationtrace;hil_inventory;hil_inventoryjournal;hil_inventoryrequest;hil_invoice;hil_jobbasecharge;hil_jobcancellationrequest;hil_joberrorcode;hil_jobestimation;hil_jobreassignreason;hil_jobsextension;hil_jobtat;hil_jobtatcategory;hil_labor;hil_leadproduct;hil_materialgroup;hil_materialgroup2;hil_materialgroup3;hil_materialgroup4;hil_materialgroup5;hil_materialreturn;hil_minimumguarantee;hil_minimumstocklevel;hil_mobileappbanner;hil_natureofcomplaint;hil_npssetup;hil_oaheader;hil_oaproduct;hil_oareadinessdatestaging;hil_observation;hil_orderchecklist;hil_orderchecklistproduct;hil_part;hil_partnerdepartment;hil_partnerdivisionmapping;hil_paymentstatus;hil_pincode;hil_plantmaster;hil_plantordertypesetup;hil_pmsconfiguration;hil_pmsconfigurationlines;hil_pmsjobuploader;hil_pmsscheduleconfiguration;hil_politicalmapping;hil_postatustracker;hil_postatustrackerupload;hil_postorder;hil_preferredlanguageforcommunication;hil_productcatalog;hil_production;hil_productrequest;hil_productrequestheader;hil_productstaging;hil_productvideos;hil_propertytype;hil_refreshjobs;hil_region;hil_remarks;hil_returnheader;hil_returnline;hil_salesoffice;hil_salesofficebranchheadmapping;hil_salesofficeplantmapping;hil_salestatmaster;hil_sawactivity;hil_sawactivityapproval;hil_sawcategoryapprovals;hil_sbubranchmapping;hil_schemedistrictexclusion;hil_schemeincentive;hil_schemeline;hil_securityroleextension;hil_serialnumber;hil_serviceactionwork;hil_serviceactionworksetup;hil_servicebom;hil_servicecallrequest;hil_servicecallrequestdetail;hil_serviceengineergeocode;hil_smsconfiguration;hil_smstemplates;hil_solarservey;hil_specialincentive;hil_staging;hil_stagingdivisonmaterialgroupmapping;hil_stagingintegrationjson;hil_stagingpricingmapping;hil_stagingsbudivisionmapping;hil_startingmethod;hil_state;hil_statustransitionmatrix;hil_subterritory;hil_tatachievementslabmaster;hil_tatbreachpenalty;hil_tatincentive;hil_technicalspecfication;hil_technician;hil_technicianhealthindicator;hil_tender;hil_tenderattachmentdoctype;hil_tenderattachmentmanager;hil_tenderbankguarantee;hil_tenderbomlineitem;hil_tenderdocs;hil_tenderemailersetup;hil_tenderpaymentdetail;hil_tenderproduct;hil_tolerencesetup;hil_travelclosureincentive;hil_travelexpense;hil_typeofcustomer;hil_typeofproduct;hil_unitwarranty;hil_upcountrydataupdate;hil_upcountrytravelcharge;hil_usageindicator;hil_userbranchmapping;hil_usersecurityroleextension;hil_usersimeibinding;hil_voltage;hil_warrantyperiod;hil_warrantyscheme;hil_warrantytemplate;hil_warrantyvoidreason;hil_whatsappproductdivisionconfig;hil_workorderarch;hil_wrongclosurepenalty;hil_yammersettings;ispl_autonumberconfiguration;new_integrationstaging;new_userlocations;plt_idgenerator";
            string[] entitys = entitylist.Split(';');

            foreach (string ent in entitys)
            {
                RetrieveEntityRequest retrieveEntityRequest = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.Entity,
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
                ViewsModificationNew.UpdateCreateView(service, serviceSer, StateEntity.ObjectTypeCode, StateEntityDemo.ObjectTypeCode);
                Console.WriteLine(ent);
            }
        }
        public static void UpdateCreateView(IOrganizationService service, IOrganizationService serviceSer, int? ObjectTypeCodeQA, int? ObjectTypeCodeDemo)
        {
            Console.WriteLine("********************** VIEWs CREATION FOR ENTITY " + ObjectTypeCodeQA.ToString().ToUpper() + " IS STARTED **********************");
            var query = new QueryExpression("savedquery");
            // query.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, ObjectTypeCodeQA);
            query.Criteria.AddCondition("name", ConditionOperator.Equal, "Address* Lookup View");

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
                    string name = ent.GetAttributeValue<string>("name");
                    if (name.Contains("Advanced Find View") || name.Contains("Quick Find"))
                    {
                        continue;
                    }
                    EntityCollection savedQueriesDemo = getDemoViewId(ent.Id.ToString(), service, serviceSer);
                    if (savedQueriesDemo.Entities.Count > 1)
                        continue;
                    Entity entity = new Entity(savedQueriesDemo[0].LogicalName, savedQueriesDemo[0].Id);
                    entity["columnsetxml"] = ent.GetAttributeValue<string>("columnsetxml");
                    entity["conditionalformatting"] = ent.GetAttributeValue<string>("conditionalformatting");
                    entity["fetchxml"] = ent.GetAttributeValue<string>("fetchxml");
                    entity["layoutjson"] = ent.GetAttributeValue<string>("layoutjson");
                    entity["layoutxml"] = ent.GetAttributeValue<string>("layoutxml");
                    entity["name"] = ent.GetAttributeValue<string>("name");
                    entity["offlinesqlquery"] = ent.GetAttributeValue<string>("offlinesqlquery");
                    entity["queryappusage"] = ent.GetAttributeValue<string>("queryappusage");
                    entity["description"] = ent.GetAttributeValue<string>("description");
                    serviceSer.Update(entity);
                    Console.WriteLine("View Update with Name " + ent.GetAttributeValue<string>("name"));
                }
                catch (Exception ex)
                {
                    error++;
                    Console.WriteLine("-------------------------Failuer " + error + " in View Updation with name " +
                        ent.GetAttributeValue<string>("name") + ". Error is:-  " + ex.Message);
                }
            }
            Console.WriteLine("********************** VIEWs CREATION FOR ENTITY " + ObjectTypeCodeQA.ToString().ToUpper() + " IS ENDED **********************");
        }
        private static EntityCollection getDemoViewId(string ViewIdPrd, IOrganizationService service, IOrganizationService serviceSer)
        {
            EntityCollection entityCollection = new EntityCollection();
            string newViewId = null;
            try
            {
                var query = new QueryExpression("savedquery");
                query.Criteria.AddCondition("savedqueryid", ConditionOperator.Equal, ViewIdPrd);
                query.ColumnSet = new ColumnSet(true);
                RetrieveMultipleRequest retrieveSavedQueriesRequest = new RetrieveMultipleRequest { Query = query };
                RetrieveMultipleResponse retrieveSavedQueriesResponse = (RetrieveMultipleResponse)service.Execute(retrieveSavedQueriesRequest);
                DataCollection<Entity> savedQueries = retrieveSavedQueriesResponse.EntityCollection.Entities;

                var query1 = new QueryExpression("savedquery");
                //query1.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, savedQueries[0].GetAttributeValue<string>("returnedtypecode"));
                if (savedQueries[0].GetAttributeValue<string>("name") == "Channel Partner Associated View")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Account Associated View");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Consumers Lookup View")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Contacts Lookup View");

                else if (savedQueries[0].GetAttributeValue<string>("name") == "All Channel Partners")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "All Accounts");

                else if (savedQueries[0].GetAttributeValue<string>("name") == "Channel Partner Lookup View")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Account Lookup View");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Jobs Lookup View")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Work Order Lookup View");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Allow On Job Lookup")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Allow On Work Order Lookup");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Cause Lookup View")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Incident Type Lookup View");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Custom Job Incident Products View")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Custom Work Order Products View");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Active Causes")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Active Incident Types");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Active Consumers")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Active Contacts");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "My Active Consumers")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "My Active Contacts");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Active Channel Partners")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Active Accounts");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "My Active Channel Partners")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "My Active Accounts");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Active Jobs")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Active Work Orders");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Active Cause Spare Parts")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Active Incident Type Products");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Cause Spare Parts")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Incident Type Products");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Cause Spare Parts")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Incident Type Products");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Jobs Without Cancelled View")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Work Order Without Cancelled View");
                else
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, savedQueries[0].GetAttributeValue<string>("name"));
                //query1.Criteria.AddCondition("iscustom", ConditionOperator.Equal, false);
                query1.ColumnSet = new ColumnSet(true);
                RetrieveMultipleRequest retrieveSavedQueriesRequest1 = new RetrieveMultipleRequest { Query = query1 };
                RetrieveMultipleResponse retrieveSavedQueriesResponse1 = (RetrieveMultipleResponse)serviceSer.Execute(retrieveSavedQueriesRequest1);
                DataCollection<Entity> savedQueries1 = retrieveSavedQueriesResponse1.EntityCollection.Entities;
                entityCollection = retrieveSavedQueriesResponse1.EntityCollection;
                //foreach (Entity ent in savedQueries1)
                //{
                //    if (ent.GetAttributeValue<string>("returnedtypecode") == savedQueries[0].GetAttributeValue<string>("returnedtypecode"))
                //    {
                //        newViewId = ent.Id.ToString();
                //    }
                //    else
                //    {
                //        //  Console.WriteLine("dd");
                //    }
                //}
                if (savedQueries1.Count == 0)
                {
                    string nana = savedQueries[0].GetAttributeValue<string>("name");
                    Entity ddd = serviceSer.Retrieve(savedQueries[0].LogicalName, savedQueries[0].Id, new ColumnSet(true));
                    Console.WriteLine(nana);
                    entityCollection.Entities.Add(ddd);
                    //serviceSer.Create(savedQueries[0]);
                }
                else if (savedQueries1.Count > 1)
                {
                    string nana = savedQueries[0].GetAttributeValue<string>("name");
                    Console.WriteLine(nana);
                    //serviceSer.Create(savedQueries[0]);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return entityCollection;
        }
    }
    class ViewsModification
    {
        public static void UpdateCreateView(IOrganizationService service, IOrganizationService serviceSer, int? ObjectTypeCodeQA, int? ObjectTypeCodeDemo)
        {
            Console.WriteLine("********************** VIEWs CREATION FOR ENTITY " + ObjectTypeCodeQA.ToString().ToUpper() + " IS STARTED **********************");
            var query = new QueryExpression("savedquery");
            query.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, ObjectTypeCodeQA);
            query.ColumnSet = new ColumnSet(true);
            RetrieveMultipleRequest retrieveSavedQueriesRequest = new RetrieveMultipleRequest { Query = query };
            RetrieveMultipleResponse retrieveSavedQueriesResponse = (RetrieveMultipleResponse)service.Execute(retrieveSavedQueriesRequest);
            DataCollection<Entity> savedQueries = retrieveSavedQueriesResponse.EntityCollection.Entities;
            //Display the Retrieved views
            Console.WriteLine("totla Views Count " + savedQueries.Count);
            int totalCount = savedQueries.Count;
            int done = 1;
            //Console.Clear();
            int error = 0;
            foreach (Entity ent in savedQueries)
            {
                try
                {

                    //var queryDemo = new QueryExpression("savedquery");
                    //queryDemo.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, ObjectTypeCodeDemo);
                    //queryDemo.Criteria.AddCondition("querytype", ConditionOperator.Equal, ent.GetAttributeValue<int>("querytype"));
                    //queryDemo.Criteria.AddCondition("name", ConditionOperator.Equal, ent.GetAttributeValue<string>("name"));
                    //queryDemo.ColumnSet = new ColumnSet(true);
                    //RetrieveMultipleRequest retrieveSavedQueriesRequestDemo = new RetrieveMultipleRequest { Query = queryDemo };
                    //RetrieveMultipleResponse retrieveSavedQueriesResponseDemo = (RetrieveMultipleResponse)serviceSer.Execute(retrieveSavedQueriesRequestDemo);
                    //DataCollection<Entity> savedQueriesDemo = retrieveSavedQueriesResponseDemo.EntityCollection.Entities;
                    //if (savedQueriesDemo.Count != 1)
                    //    Console.WriteLine("more than 1 rec");
                    string name = ent.GetAttributeValue<string>("name");
                    EntityCollection savedQueriesDemo = getDemoViewId(ent.Id.ToString(), service, serviceSer);

                    Entity entity = new Entity(savedQueriesDemo[0].LogicalName, savedQueriesDemo[0].Id);
                    //ent.Id = savedQueriesDemo[0].Id;
                    //ent["savedqueryid"] = savedQueriesDemo[0].Id;
                    //ent[""] = false;

                    //savedQueriesDemo[0]["columnsetxml"] = ent.GetAttributeValue<string>("columnsetxml");

                    entity["columnsetxml"] = ent.GetAttributeValue<string>("columnsetxml");
                    entity["fetchxml"] = ent.GetAttributeValue<string>("fetchxml");
                    entity["layoutjson"] = ent.GetAttributeValue<string>("layoutjson");
                    entity["layoutxml"] = ent.GetAttributeValue<string>("layoutxml");
                    entity["name"] = ent.GetAttributeValue<string>("name");
                    entity["description"] = ent.GetAttributeValue<string>("description");

                    //entity["fetchxml"] = ent.GetAttributeValue<string>("fetchxml");
                    //entity["layoutjson"] = ent.GetAttributeValue<string>("layoutjson");
                    //entity["layoutxml"] = ent.GetAttributeValue<string>("layoutxml");
                    //entity["name"] = ent.GetAttributeValue<string>("name");
                    serviceSer.Update(entity);
                    Console.WriteLine("View Update with Name " + ent.GetAttributeValue<string>("name"));
                    //savedQueriesDemo[0].GetAttributeValue<string>("name");

                }
                catch (Exception ex)
                {
                    error++;
                    Console.WriteLine("-------------------------Failuer " + error + " in View Updation with name " +
                        ent.GetAttributeValue<string>("name") + ". Error is:-  " + ex.Message);
                }

            }
            Console.WriteLine("********************** VIEWs CREATION FOR ENTITY " + ObjectTypeCodeQA.ToString().ToUpper() + " IS ENDED **********************");
        }
        public static EntityCollection getDemoViewId(string ViewIdPrd, IOrganizationService service, IOrganizationService serviceSer)
        {
            EntityCollection entityCollection = new EntityCollection();
            string newViewId = null;
            try
            {
                var query = new QueryExpression("savedquery");
                query.Criteria.AddCondition("savedqueryid", ConditionOperator.Equal, ViewIdPrd);
                query.ColumnSet = new ColumnSet(true);
                RetrieveMultipleRequest retrieveSavedQueriesRequest = new RetrieveMultipleRequest { Query = query };
                RetrieveMultipleResponse retrieveSavedQueriesResponse = (RetrieveMultipleResponse)service.Execute(retrieveSavedQueriesRequest);
                DataCollection<Entity> savedQueries = retrieveSavedQueriesResponse.EntityCollection.Entities;

                var query1 = new QueryExpression("savedquery");
                //query1.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, savedQueries[0].GetAttributeValue<string>("returnedtypecode"));
                if (savedQueries[0].GetAttributeValue<string>("name") == "Consumers Lookup View")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Contacts Lookup View");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Channel Partner Lookup View")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Account Lookup View");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Jobs Lookup View")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Work Order Lookup View");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Allow On Job Lookup")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Allow On Work Order Lookup");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Cause Lookup View")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Incident Type Lookup View");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Custom Job Incident Products View")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Custom Work Order Products View");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Active Causes")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Active Incident Types");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Active Consumers")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Active Contacts");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "My Active Consumers")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "My Active Contacts");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Active Channel Partners")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Active Accounts");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "My Active Channel Partners")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "My Active Accounts");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Active Jobs")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Active Work Orders");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Active Cause Spare Parts")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Active Incident Type Products");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Cause Spare Parts")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Incident Type Products");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Cause Spare Parts")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Incident Type Products");
                else if (savedQueries[0].GetAttributeValue<string>("name") == "Jobs Without Cancelled View")
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Work Order Without Cancelled View");
                else
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, savedQueries[0].GetAttributeValue<string>("name"));
                //query1.Criteria.AddCondition("iscustom", ConditionOperator.Equal, false);
                query1.ColumnSet = new ColumnSet(true);
                RetrieveMultipleRequest retrieveSavedQueriesRequest1 = new RetrieveMultipleRequest { Query = query1 };
                RetrieveMultipleResponse retrieveSavedQueriesResponse1 = (RetrieveMultipleResponse)serviceSer.Execute(retrieveSavedQueriesRequest1);
                DataCollection<Entity> savedQueries1 = retrieveSavedQueriesResponse1.EntityCollection.Entities;
                entityCollection = retrieveSavedQueriesResponse1.EntityCollection;
                //foreach (Entity ent in savedQueries1)
                //{
                //    if (ent.GetAttributeValue<string>("returnedtypecode") == savedQueries[0].GetAttributeValue<string>("returnedtypecode"))
                //    {
                //        newViewId = ent.Id.ToString();
                //    }
                //    else
                //    {
                //        //  Console.WriteLine("dd");
                //    }
                //}
                if (savedQueries1.Count == 0)
                {
                    string nana = savedQueries[0].GetAttributeValue<string>("name");
                    Console.WriteLine(nana);
                    serviceSer.Create(savedQueries[0]);
                }
                else if (savedQueries1.Count > 1)
                {
                    string nana = savedQueries[0].GetAttributeValue<string>("name");
                    Console.WriteLine(nana);
                    //serviceSer.Create(savedQueries[0]);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return entityCollection;
        }
    }
    class FormModfication
    {
        static Dictionary<string, string> viewIdsDict = new Dictionary<string, string>();
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
                //if (entity.GetAttributeValue<string>("name") != "Job")
                //    continue;
                try
                {

                    Entity entity1 = entity;// service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet(true));
                    var query1 = new QueryExpression("systemform");
                    query1.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, ObjectTypeCodeDemo);

                    if (entity1.GetAttributeValue<string>("name") == "Channel Partner")
                    {
                        query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Account");// entity.GetAttributeValue<string>("name"));
                    }
                    else if (entity1.GetAttributeValue<string>("name") == "Consumer")
                    {
                        query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Contact");// entity.GetAttributeValue<string>("name"));
                    }
                    else
                        query1.Criteria.AddCondition("name", ConditionOperator.Equal, entity1.GetAttributeValue<string>("name"));
                    query1.Criteria.AddCondition("type", ConditionOperator.Equal, entity1.GetAttributeValue<OptionSetValue>("type").Value);
                    //entity.FormattedValues["type"]
                    query1.ColumnSet = new ColumnSet(true);
                    var results1 = serviceSer.RetrieveMultiple(query1);
                    if (results1.Entities.Count > 0)
                    {

                        if (results1[0].FormattedValues["type"] == "Main")
                        {
                            Console.WriteLine(results1[0].FormattedValues["type"]);
                        }

                        string formXML = entity1.GetAttributeValue<string>("formxml");
                        string formJSON = entity1.GetAttributeValue<string>("formjson");
                        viewIdsDict = new Dictionary<string, string>();
                        ChaneFormXml(formXML, service, serviceSer);

                        foreach (string key in viewIdsDict.Keys)
                        {
                            //Console.WriteLine(key + " : " + viewIdsDict[key]);
                            formXML = formXML.Replace(key.ToUpper(), viewIdsDict[key].ToUpper());
                            formJSON = formJSON.Replace(key.ToUpper(), viewIdsDict[key].ToUpper());
                        }

                        results1[0]["description"] = entity1.GetAttributeValue<string>("description");
                        results1[0]["formactivationstate"] = entity1.GetAttributeValue<OptionSetValue>("formactivationstate");
                        results1[0]["formjson"] = formJSON;// entity.GetAttributeValue<string>("formjson");
                        results1[0]["formpresentation"] = entity1.GetAttributeValue<OptionSetValue>("formpresentation");
                        results1[0]["formxml"] = formXML; //entity.GetAttributeValue<string>("formxml");
                        results1[0]["isdefault"] = entity1.GetAttributeValue<bool>("isdefault");
                        results1[0]["isdesktopenabled"] = entity1.GetAttributeValue<bool>("isdesktopenabled");
                        results1[0]["istabletenabled"] = entity1.GetAttributeValue<bool>("istabletenabled");
                        results1[0]["type"] = entity1.GetAttributeValue<OptionSetValue>("type");
                        results1[0]["version"] = entity1.GetAttributeValue<int>("version");
                        results1[0]["name"] = entity1.GetAttributeValue<string>("name");


                        serviceSer.Update(results1[0]);
                        Console.WriteLine(results1[0].FormattedValues["type"]);
                        Console.WriteLine("From with name " + entity1.GetAttributeValue<string>("name") + " is update " + done + " / " + totalCount);
                        done++;
                    }
                    else
                    {
                        string formXML = entity1.GetAttributeValue<string>("formxml");
                        string formJSON = entity1.GetAttributeValue<string>("formjson");
                        viewIdsDict = new Dictionary<string, string>();
                        ChaneFormXml(formXML, service, serviceSer);

                        foreach (string key in viewIdsDict.Keys)
                        {
                            //Console.WriteLine(key + " : " + viewIdsDict[key]);
                            formXML = formXML.Replace(key.ToUpper(), viewIdsDict[key].ToUpper());
                            formJSON = formJSON.Replace(key.ToUpper(), viewIdsDict[key].ToUpper());
                        }
                        entity1["formjson"] = formJSON;
                        entity1["formxml"] = formXML;
                        serviceSer.Create(entity1);
                        Console.WriteLine("From with name " + entity1.GetAttributeValue<string>("name") + " is created " + done + " / " + totalCount);
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
        public static void ChaneFormXml(string xml, IOrganizationService service, IOrganizationService serviceSer)
        {
            try
            {
                xml = xml.Replace("<ViewIds>", "|");
                xml = xml.Replace("</ViewIds>", "^");
                xml = xml.Replace("<AvailableViewIds>", "|");
                xml = xml.Replace("</AvailableViewIds>", "^");
                xml = xml.Replace("<DefaultViewId>", "|");
                xml = xml.Replace("</DefaultViewId>", "^");
                string[] allViews = xml.Split('|');
                foreach (string a in allViews)
                {
                    string[] views = a.Split('^');
                    if (views.Length == 1)
                        continue;
                    string[] vewsIds = views[0].Split(',');
                    foreach (string viewid in vewsIds)
                    {
                        var dictval = from x in viewIdsDict
                                      where x.Key.Contains(viewid)
                                      select x;
                        if (dictval.ToList().Count == 0)
                        {
                            string newViewId = getDemoViewId(viewid, service, serviceSer);
                            newViewId = newViewId == null ? "00000000-0000-0000-0000-000000000000" : newViewId;
                            viewIdsDict.Add(viewid, "{" + newViewId + "}");
                            //Console.WriteLine(viewid);
                        }
                        //else
                        //{
                        //    Console.WriteLine(dictval.First().Value);
                        //}
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


        }
        public static string getDemoViewId(string ViewIdPrd, IOrganizationService service, IOrganizationService serviceSer)
        {

            string newViewId = null;
            try
            {
                var query = new QueryExpression("savedquery");
                query.Criteria.AddCondition("savedqueryid", ConditionOperator.Equal, ViewIdPrd);
                query.ColumnSet = new ColumnSet(true);
                RetrieveMultipleRequest retrieveSavedQueriesRequest = new RetrieveMultipleRequest { Query = query };
                RetrieveMultipleResponse retrieveSavedQueriesResponse = (RetrieveMultipleResponse)service.Execute(retrieveSavedQueriesRequest);
                DataCollection<Entity> savedQueries = retrieveSavedQueriesResponse.EntityCollection.Entities;

                var query1 = new QueryExpression("savedquery");
                //if (savedQueries[0].GetAttributeValue<string>("name") == "Jobs Advanced Find View")
                //     query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Work Order Advanced Find View");
                // //else if (savedQueries[0].GetAttributeValue<string>("name") == "Allow On Job Lookup")
                // //    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Allow On Work Order Lookup");
                // //else if (savedQueries[0].GetAttributeValue<string>("name") == "Cause Lookup View")
                // //    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Incident Type Lookup View");
                // //else if (savedQueries[0].GetAttributeValue<string>("name") == "Custom Job Incident Products View")
                // //    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Custom Work Order Products View");
                // //else if (savedQueries[0].GetAttributeValue<string>("name") == "Active Causes")
                // //    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Active Incident Types");
                // //else if (savedQueries[0].GetAttributeValue<string>("name") == "Active Consumers")
                // //    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Active Contacts");
                // //else if (savedQueries[0].GetAttributeValue<string>("name") == "My Active Consumers")
                // //    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "My Active Contacts");
                // //else if (savedQueries[0].GetAttributeValue<string>("name") == "Active Channel Partners")
                // //    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Active Accounts");
                // //else if (savedQueries[0].GetAttributeValue<string>("name") == "My Active Channel Partners")
                // //    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "My Active Accounts");
                // //else if (savedQueries[0].GetAttributeValue<string>("name") == "Active Jobs")
                // //    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Active Work Orders");
                // //else if (savedQueries[0].GetAttributeValue<string>("name") == "Active Cause Spare Parts")
                // //    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Active Incident Type Products");
                // //else if (savedQueries[0].GetAttributeValue<string>("name") == "Cause Spare Parts")
                // //    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Incident Type Products");
                // //else if (savedQueries[0].GetAttributeValue<string>("name") == "Cause Spare Parts")
                // //    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Incident Type Products");
                // //else if (savedQueries[0].GetAttributeValue<string>("name") == "Jobs Without Cancelled View")
                // //    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Work Order Without Cancelled View");
                // else
                query1.Criteria.AddCondition("name", ConditionOperator.Equal, savedQueries[0].GetAttributeValue<string>("name"));
                //query1.Criteria.AddCondition("iscustom", ConditionOperator.Equal, false);
                query1.ColumnSet = new ColumnSet(true);
                RetrieveMultipleRequest retrieveSavedQueriesRequest1 = new RetrieveMultipleRequest { Query = query1 };
                RetrieveMultipleResponse retrieveSavedQueriesResponse1 = (RetrieveMultipleResponse)serviceSer.Execute(retrieveSavedQueriesRequest1);
                DataCollection<Entity> savedQueries1 = retrieveSavedQueriesResponse1.EntityCollection.Entities;

                foreach (Entity ent in savedQueries1)
                {
                    if (ent.GetAttributeValue<string>("returnedtypecode") == savedQueries[0].GetAttributeValue<string>("returnedtypecode"))
                    {
                        //ent["name"] = savedQueries[0].GetAttributeValue<string>("name");
                        //serviceSer.Update(ent);
                        newViewId = ent.Id.ToString();
                    }
                    //else
                    //{
                    //    //  Console.WriteLine("dd");
                    //}
                }
                if (savedQueries1.Count == 0)
                {
                    string nana = savedQueries[0].GetAttributeValue<string>("name");
                    Console.WriteLine(nana);
                    //serviceSer.Create(savedQueries[0]);
                }

            }
            catch (Exception ex)
            {
                if (!ex.Message.Contains("The entity with a name ="))
                    Console.WriteLine(ex.Message);
            }
            if (newViewId == null)
                Console.WriteLine("dd");
            return newViewId;
        }
    }
}
