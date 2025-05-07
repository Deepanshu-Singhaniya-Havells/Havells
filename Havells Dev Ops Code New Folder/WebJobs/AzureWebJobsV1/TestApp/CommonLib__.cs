using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.IO;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;
using WorkOrderVirtualEntity.Models;

namespace TestApp
{
    public class CommonLib
    {
        public static EntityCollection getData(IOrganizationService service, ITracingService tracingService, string eurl, String entName)
        {
            EntityCollection entcol = new EntityCollection();

            tracingService.Trace("GetTestData Function Strated");
            String response = string.Empty;

            IntegrationConfiguration inconfig = GetIntegrationConfiguration(service);
            String sUrl = inconfig.url + eurl;
            String sUserName = inconfig.userName;
            string sPassword = inconfig.password;
            tracingService.Trace("sUrl " + sUrl);

            tracingService.Trace("sUserName " + sUserName);

            tracingService.Trace("sPassword " + sPassword);

            WebClient webClient = new WebClient();
            string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
            webClient.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
            var jsonData = webClient.DownloadData(sUrl);
            tracingService.Trace("Data ******************************************");
            tracingService.Trace(jsonData.ToString());
            tracingService.Trace("******************************************");

            if (entName == "hil_jobincidentarchived")
            {
                tracingService.Trace("hil_jobincidentarchived");
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(IncidentResponse));
                IncidentResponse rootObject = (IncidentResponse)ser.ReadObject(new MemoryStream(jsonData));
                tracingService.Trace("rootObject.success " + rootObject.success);
                if (rootObject.success)
                {
                    tracingService.Trace("rootObject.workOrders.Count:- " + rootObject.incident.Count);
                    foreach (dtoIncidents _incident in rootObject.incident)
                    {
                        Entity job = DataCollectionsforIncident(service, tracingService, _incident);
                        entcol.Entities.AddRange(job);
                        tracingService.Trace("entcol.Entities.Count :-   " + entcol.Entities.Count);
                    }
                }
            }
            else if (entName == "hil_jobsarchived")
            {
                tracingService.Trace("hil_jobsarchived");
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(ArchivedJobsResponse));
                ArchivedJobsResponse rootObject = (ArchivedJobsResponse)ser.ReadObject(new MemoryStream(jsonData));
                tracingService.Trace("rootObject.success " + rootObject.success);
                if (rootObject.success)
                {
                    tracingService.Trace("rootObject.workOrders.Count:- " + rootObject.workOrders.Count);
                    foreach (dtoWorkOrders _job in rootObject.workOrders)
                    {
                        Entity job = DataCollectionsforJob(service, tracingService, _job);
                        entcol.Entities.AddRange(job);
                    }
                }
            }
            else if (entName == "hil_jobproductarchived")
            {
                //DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(WorkOrderProductResponse));
                //WorkOrderProductResponse rootObject = (WorkOrderProductResponse)ser.ReadObject(new MemoryStream(jsonData));
                //tracingService.Trace("rootObject.success " + rootObject.success);
                //if (rootObject.success)
                //{
                //    tracingService.Trace("rootObject.workOrders.Count:- " + rootObject.orderProduct.Count);
                //    foreach (dtoWorkorderProduct _serviceEnt in rootObject.orderProduct)
                //    {
                //        Entity job = DataCollectionsforService(service, tracingService, _serviceEnt);
                //        entcol.Entities.AddRange(job);
                //    }
                //}
            }
            else if (entName == "hil_jobservicearchived")
            {
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(WorkOrderServiceResponse));
                WorkOrderServiceResponse rootObject = (WorkOrderServiceResponse)ser.ReadObject(new MemoryStream(jsonData));
                tracingService.Trace("rootObject.success " + rootObject.success);
                if (rootObject.success)
                {
                    tracingService.Trace("rootObject.workOrders.Count:- " + rootObject.orderService.Count);
                    foreach (dtoWorkOrderService _serviceEnt in rootObject.orderService)
                    {
                        Entity job = DataCollectionsforService(service, tracingService, _serviceEnt);
                        entcol.Entities.AddRange(job);
                    }
                }

            }

            tracingService.Trace("entcol.Entities.Count :-   " + entcol.Entities.Count);
            return entcol;
        }
        private static Entity DataCollectionsforJob(IOrganizationService service, ITracingService tracingService, dtoWorkOrders _job)
        {
            Entity entity = new Entity("hil_jobsarchived");
            tracingService.Trace("DataCollections method  started");

            if (_job.createdby != null)
            {
                entity["hil_createdby"] = createEntityRef("systemuser", new Guid(_job.createdby), service);
                tracingService.Trace("createdby");
            }
            if (_job.createdon != null)
            {
                tracingService.Trace("createdon");
                entity["hil_createdon"] = dateFormater(_job.createdon, tracingService);
            }
            if (_job.createdonbehalfby != null)
            {
                tracingService.Trace("hil_createdonbehalfby");
                entity["hil_createdonbehalfby"] = createEntityRef("systemuser", new Guid(_job.createdonbehalfby), service);
            }
            if (_job.exchangerate != null)
            {
                tracingService.Trace("exchangerate");
                entity["hil_exchangerate"] = Convert.ToDecimal(_job.exchangerate);
            }
            if (_job.hil_actualcharges != null)
            {
                entity["hil_actualcharges"] = Convert.ToDecimal(_job.hil_actualcharges);
                tracingService.Trace("hil_actualcharges");
            }

            if (_job.hil_address != null)
            {
                entity["hil_address"] = createEntityRef("hil_address", new Guid(_job.hil_address), service);
                tracingService.Trace("hil_address");
            }
            if (_job.msdyn_name != null)
            {
                tracingService.Trace("msdyn_name");
                entity["hil_name"] = _job.msdyn_name;
            }
            if (_job.msdyn_substatus != null)
            {
                tracingService.Trace("msdyn_substatus");
                entity["hil_substatus"] = createEntityRef("msdyn_workordersubstatus", new Guid(_job.msdyn_substatus), service);
            }
            if (_job.hil_productsubcategory != null)
            {
                tracingService.Trace("hil_productsubcategory");
                entity["hil_productsubcategory"] = createEntityRef("product", new Guid(_job.hil_productsubcategory), service);
            }
            if (_job.hil_callsubtype != null)
            {
                tracingService.Trace("hil_callsubtype");
                entity["hil_callsubtype"] = createEntityRef("hil_callsubtype", new Guid(_job.hil_callsubtype), service);
            }
            if (_job.hil_natureofcomplaint != null)
            {
                tracingService.Trace("hil_natureofcomplaint");
                entity["hil_natureofcomplaint"] = createEntityRef("hil_natureofcomplaint", new Guid(_job.hil_natureofcomplaint), service);
            }
            if (_job.ownerid != null)
            {
                tracingService.Trace("hil_ownerid");
                entity["hil_ownerid"] = createEntityRef("systemuser", new Guid(_job.ownerid), service);
            }
            if (_job.hil_branch != null)
            {
                tracingService.Trace("hil_branch");
                entity["hil_branch"] = createEntityRef("hil_branch", new Guid(_job.hil_branch), service);
            }
            if (_job.hil_district != null)
            {
                tracingService.Trace("hil_district");
                entity["hil_district"] = createEntityRef("hil_district", new Guid(_job.hil_district), service);
            }
            if (_job.hil_mobilenumber != null)
            {
                tracingService.Trace("hil_mobilenumber");
                entity["hil_mobilenumber"] = _job.hil_mobilenumber;
            }
            if (_job.hil_customerref != null)
            {
                tracingService.Trace("hil_customerref");
                entity["hil_customerref"] = createEntityRef("contact", new Guid(_job.hil_customerref), service);
            }
            if (_job.hil_owneraccount != null)
            {
                tracingService.Trace("hil_owneraccount");
                entity["hil_owneraccount"] = createEntityRef("account", new Guid(_job.hil_owneraccount), service);
            }
            if (_job.hil_escallationcountinteger != null)
            {
                tracingService.Trace("hil_escallationcountinteger");
                entity["hil_escallationcountinteger"] = int.Parse(_job.hil_escallationcountinteger);
            }
            if (_job.hil_remaindercountinteger != null)
            {
                tracingService.Trace("hil_remaindercountinteger");
                entity["hil_remaindercountinteger"] = int.Parse(_job.hil_remaindercountinteger);
            }
            if (_job.hil_appointmenttime != null)
            {
                tracingService.Trace("hil_appointmenttime");
                entity["hil_appointmenttime"] = dateFormater(_job.hil_appointmenttime, tracingService);
            }
            if (_job.hil_fulladdress != null)
            {
                tracingService.Trace("hil_fulladdress");
                entity["hil_fulladdress"] = _job.hil_fulladdress;
            }
            if (_job.hil_pincode != null)
            {
                tracingService.Trace("hil_pincode");
                entity["hil_pincode"] = createEntityRef("hil_pincode", new Guid(_job.hil_pincode), service);
            }
            if (_job.hil_warrantystatus != null)
            {
                tracingService.Trace("hil_warrantystatus");
                entity["hil_warrantystatus"] = new OptionSetValue(Convert.ToInt32(_job.hil_warrantystatus));
            }
            if (_job.hil_productcategory != null)
            {
                tracingService.Trace("hil_productcategory");
                entity["hil_productcategory"] = createEntityRef("product", new Guid(_job.hil_productcategory), service);
            }
            if (_job.hil_sourceofjob != null)
            {
                tracingService.Trace("hil_sourceofjob");
                entity["hil_sourceofjob"] = new OptionSetValue(Convert.ToInt32(_job.hil_sourceofjob));
            }
            if (_job.msdyn_parentworkorder != null)
            {
                tracingService.Trace("msdyn_parentworkorder");
                entity["hil_parentworkorder"] = new EntityReference("hil_jobsarchived", new Guid(_job.msdyn_parentworkorder));
            }
            if (_job.hil_pmscount != null)
            {
                tracingService.Trace("hil_pmscount");
                entity["hil_pmscount"] = int.Parse(_job.hil_pmscount);
            }
            if (_job.msdyn_workorderid != null)
            {
                tracingService.Trace("msdyn_workorderid");
                entity["hil_jobsarchivedid"] = new Guid(_job.msdyn_workorderid);
            }
            tracingService.Trace("DataCollections method  end");
            return entity;
        }
        private static Entity DataCollectionsforIncident(IOrganizationService service, ITracingService tracingService, dtoIncidents _incident)
        {
            Entity entity = new Entity("hil_jobincidentarchived");
            tracingService.Trace("DataCollectionsforIncident method  started");

            if (_incident.msdyn_customerasset != null)
            {
                entity["hil_customerasset"] = createEntityRef("msdyn_customerasset", new Guid(_incident.msdyn_customerasset), service);
                tracingService.Trace("msdyn_customerasset");
            }
            if (_incident.hil_natureofcomplaint != null)
            {
                entity["hil_natureofcomplaint"] = createEntityRef("hil_natureofcomplaint", new Guid(_incident.hil_natureofcomplaint), service);
                tracingService.Trace("hil_natureofcomplaint");
                //
            }
            if (_incident.hil_observation != null)
            {
                entity["hil_observation"] = createEntityRef("hil_observation", new Guid(_incident.hil_observation), service);
                tracingService.Trace("hil_observation");
                //
            }
            if (_incident.msdyn_incidenttype != null)
            {
                entity["hil_cause"] = createEntityRef("msdyn_incidenttype", new Guid(_incident.msdyn_incidenttype), service);
                tracingService.Trace("msdyn_incidenttype");
                //
            }
            if (_incident.hil_quantity != null)
            {
                entity["hil_quantity"] = int.Parse(_incident.hil_quantity.ToString());
                tracingService.Trace("hil_quantity");
                //
            }
            if (_incident.new_warrantyenddate != null)
            {
                entity["hil_warrantyenddate"] = dateFormater(_incident.new_warrantyenddate.ToString(), tracingService);
                tracingService.Trace("new_warrantyenddate");
                //
            }
            if (_incident.hil_warrantystatus != null)
            {
                entity["hil_warrantystatus"] = int.Parse(_incident.hil_warrantystatus.ToString()) == 1 ? true : false;
                tracingService.Trace("hil_warrantystatus");
            }
            if (_incident.hil_customer != null)
            {
                entity["hil_customer"] = createEntityRef("contact", new Guid(_incident.hil_customer), service);
                tracingService.Trace("hil_customer");
                //  
            }
            if (_incident.msdyn_workorder != null)
            {
                entity["hil_jobs"] = createEntityRef("hil_jobsarchived", new Guid(_incident.msdyn_workorder), service);
                tracingService.Trace("msdyn_workorder");
                //  
            }
            if (_incident.ownerid != null)
            {
                tracingService.Trace("hil_owner");
                entity["hil_owner"] = createEntityRef("systemuser", new Guid(_incident.ownerid), service);
            }
            if (_incident.hil_productcategory != null)
            {
                tracingService.Trace("hil_productsubcategory");
                entity["hil_productcategory"] = createEntityRef("product", new Guid(_incident.hil_productcategory), service);
            }
            if (_incident.hil_productsubcategory != null)
            {
                tracingService.Trace("hil_productsubcategory");
                entity["hil_productsubcategory"] = createEntityRef("product", new Guid(_incident.hil_productsubcategory), service);
            }
            if (_incident.hil_modelcode != null)
            {
                tracingService.Trace("hil_modelcode");
                entity["hil_modelcode"] = createEntityRef("product", new Guid(_incident.hil_modelcode), service);
            }
            if (_incident.hil_productreplacement != null)
            {
                entity["hil_productreplacement"] = new OptionSetValue(int.Parse(_incident.hil_productreplacement.ToString()));
            }
            if (_incident.msdyn_name != null)
            {
                entity["hil_name"] = _incident.msdyn_name;
            }
            if (_incident.hil_inputwatertds != null)
            {
                entity["hil_inputwatertds"] = _incident.hil_inputwatertds;
            }
            if (_incident.hil_rejectwatertds != null)
            {
                entity["hil_rejectwatertds"] = _incident.hil_rejectwatertds;
            }
            if (_incident.hil_purewatertds != null)
            {
                entity["hil_purewatertds"] = _incident.hil_purewatertds;
            }
            if (_incident.hil_ph != null)
            {
                entity["hil_ph"] = _incident.hil_ph;
            }
            if (_incident.hil_noofpeopleinhome != null)
            {
                entity["hil_noofpeopleinhome"] = _incident.hil_noofpeopleinhome;
            }
            if (_incident.hil_purewatertds != null)
            {
                entity["hil_purewatertds"] = _incident.hil_purewatertds;
            }
            if (_incident.hil_waterstoragetype != null)
            {
                entity["hil_waterstoragetype"] = new OptionSetValue(int.Parse(_incident.hil_waterstoragetype.ToString()));
            }
            if (_incident.hil_watersource != null)
            {
                entity["hil_watersource"] = new OptionSetValue(int.Parse(_incident.hil_watersource.ToString()));
            }
            if (_incident.msdyn_workorderincidentid != null)
            {
                tracingService.Trace("msdyn_workorderid" + (_incident.msdyn_workorderincidentid));
                entity["hil_jobincidentarchivedid"] = new Guid(_incident.msdyn_workorderincidentid);
            }
            //if (_incident.hil_itemspopulated != null)
            //{
            //    entity["hil_watersource"] = new OptionSetValue(int.Parse(_incident.hil_watersource.ToString()));
            //}
            tracingService.Trace("DataCollections method  end");
            return entity;
        }
        private static Entity DataCollectionsforService(IOrganizationService service, ITracingService tracingService, dtoWorkOrderService _incident)
        {
            Entity entity = new Entity("hil_jobservicearchived");
            tracingService.Trace("DataCollectionsforIncident method  started");

            if (_incident.msdyn_customerasset != null)
            {
                entity["hil_customerasset"] = createEntityRef("msdyn_customerasset", new Guid(_incident.msdyn_customerasset), service);
                tracingService.Trace("msdyn_customerasset");
            }
            if (_incident.hil_markused != null)
            {
            //    entity["hil_markused"] = new OptionSetValue(int.Parse(_incident.hil_markused.ToString()));
///tracingService.Trace("hil_markused");
              //  //
            }
            if (_incident.msdyn_workorder != null)
            {
                entity["hil_jobs"] = new EntityReference("msdyn_workorder", new Guid(_incident.msdyn_workorder));
                tracingService.Trace("job");
                //
            }
            if (_incident.msdyn_product != null)
            {
                entity["hil_service"] = createEntityRef("product", new Guid(_incident.msdyn_product), service);
                tracingService.Trace("job");
                //
            }
            if (_incident.msdyn_workorderserviceid != null)
            {
                entity["msdyn_workorderserviceid"] =new Guid(_incident.msdyn_workorderserviceid);
                tracingService.Trace("job");
                //
            }
            if (_incident.hil_warrantystatus != null)
            {
                //entity["hil_warrantystatus"] = _incident.hil_warrantystatus;
              //  tracingService.Trace("job");
                //
            }
            if (_incident.msdyn_totalamount != null)
            {
                entity["hil_charge"] = decimal.Parse(_incident.hil_warrantystatus);
                tracingService.Trace("hil_charge");
                //
            }
            if (_incident.createdon != null)
            {
                entity["hil_createdon"] = dateFormater(_incident.createdon.ToString(), tracingService);
                tracingService.Trace("createdon");
                //
            }
            if (_incident.ownerid != null)
            {
                entity["hil_owner"] = createEntityRef("systemuser", new Guid(_incident.ownerid), service);
                tracingService.Trace("createdon");
                //
            }

            tracingService.Trace("DataCollections method  end");
            return entity;
        }

        private static Entity DataCollectionsforProduct(IOrganizationService service, ITracingService tracingService, dtoWorkorderProduct _product)
        {
            Entity entity = new Entity("hil_jobproductarchived");
            tracingService.Trace("DataCollectionsforIncident method  started");

            if (_product.msdyn_customerasset != null)
            {
                entity["hil_customerasset"] = createEntityRef("msdyn_customerasset", new Guid(_product.msdyn_customerasset), service);
                tracingService.Trace("hil_customerasset");
            }
            if (_product.msdyn_allocated != null)
            {
                entity["hil_allocated"] = _product.msdyn_allocated;
                tracingService.Trace("hil_allocated");
            }
            if (_product.hil_availabilitystatus != null)
            {
                entity["hil_availabilitystatus"] = new OptionSetValue(int.Parse(_product.hil_availabilitystatus.ToString())); ;
                tracingService.Trace("hil_availabilitystatus");
            }
            if (_product.createdby != null)
            {
                entity["hil_createdby"] = createEntityRef("systemuser", new Guid(_product.createdby), service);
                tracingService.Trace("hil_service");
            }
            if (_product.hil_defectiveserialnumber != null)
            {
                entity["hil_defectiveserialnumber"] = _product.hil_defectiveserialnumber;
                tracingService.Trace("hil_defectiveserialnumber");
            }
            if (_product.msdyn_description != null)
            {
                entity["hil_description"] = _product.msdyn_description;
                tracingService.Trace("hil_description");
            }
            if (_product.hil_isserialized != null)
            {
                entity["hil_isserialized"] = new OptionSetValue(int.Parse(_product.hil_isserialized.ToString())); ;
                tracingService.Trace("hil_isserialized");
            }
            if (_product.hil_iswarrantyvoid != null)
            {
                entity["hil_iswarrantyvoid"] = _product.hil_iswarrantyvoid;
                tracingService.Trace("hil_iswarrantyvoid");
            }
            if (_product.msdyn_workorderproductid != null)
            {
                entity["hil_jobproductarchivedid"] = new Guid(_product.msdyn_workorderproductid);
                tracingService.Trace("hil_jobproductarchivedid");
            }
            if (_product.hil_markused != null)
            {
                entity["hil_markused"] = _product.hil_markused;
                tracingService.Trace("hil_markused");
            }
            if (_product.hil_maxquantity != null)
            {
                entity["hil_maxquantity"] = decimal.Parse(_product.hil_maxquantity.ToString());
                tracingService.Trace("hil_maxquantity");
            }
            if (_product.msdyn_product != null)
            {
                entity["hil_msdyn_product"] = createEntityRef("product", new Guid(_product.msdyn_product), service);
                tracingService.Trace("hil_msdyn_product");
            }
            if (_product.msdyn_name != null)
            {
                entity["hil_name"] = _product.msdyn_name;
                tracingService.Trace("hil_name");
            }
            if (_product.hil_part != null)
            {
                entity["hil_part"] = _product.hil_part;
                tracingService.Trace("hil_part");
            }
            if (_product.hil_partamount != null)
            {
                entity["hil_partamount"] = decimal.Parse(_product.hil_partamount.ToString());
                tracingService.Trace("hil_partamount");
            }
            if (_product.hil_linestatus != null)
            {
                entity["hil_partstatus"] = new OptionSetValue(int.Parse(_product.hil_linestatus.ToString()));
                tracingService.Trace("hil_partstatus");
            }
            if (_product.hil_pendingquantity != null)
            {
                entity["hil_pendingquantity"] = int.Parse(_product.hil_pendingquantity.ToString());
                tracingService.Trace("hil_pendingquantity");
            }
            if (_product.hil_priority != null)
            {
                entity["hil_priority"] = _product.hil_priority;
                tracingService.Trace("hil_priority");
            }
            if (_product.hil_quantity != null)
            {
                entity["hil_quantity"] = decimal.Parse(_product.hil_quantity.ToString());
                tracingService.Trace("hil_quantity");
            }
            if (_product.hil_replacedpart != null)
            {
                entity["hil_replacedpart"] = createEntityRef("product", new Guid(_product.hil_replacedpart), service);
                tracingService.Trace("hil_replacedpart");
            }
            if (_product.hil_replacedserialnumber != null)
            {
                entity["hil_replacedserialnumber"] = _product.hil_replacedserialnumber;
                tracingService.Trace("hil_replacedserialnumber");
            }
            if (_product.hil_replacesamepart != null)
            {
                entity["hil_replacesamepart"] = _product.hil_replacesamepart;
                tracingService.Trace("hil_replacesamepart");
            }
            if (_product.hil_totalamount != null)
            {
                entity["hil_totalamount"] = decimal.Parse(_product.hil_totalamount.ToString());
                tracingService.Trace("hil_totalamount");
            }
            if (_product.hil_warrantystatus != null)
            {
                entity["hil_warrantystatus"] = new OptionSetValue(int.Parse(_product.hil_warrantystatus.ToString()));
                tracingService.Trace("hil_warrantystatus");
            }
            if (_product.msdyn_workorder != null)
            {
                entity["hil_jobs"] = new EntityReference("hil_jobsarchived", new Guid(_product.msdyn_workorder));
                tracingService.Trace("job");
            }
            if (_product.msdyn_workorderincident != null)
            {
                entity["hil_jobincident"] = new EntityReference("hil_jobincidentarchived", new Guid(_product.msdyn_workorderincident));
                tracingService.Trace("hil_jobincident");
            }
            if (_product.createdon != null)
            {
                entity["hil_createdon"] = dateFormater(_product.createdon.ToString(), tracingService);
                tracingService.Trace("createdon");
            }
            //if (_product.ownerid != null)
            //{
            //    entity["hil_owner"] = createEntityRef("systemuser", new Guid(_incident.ownerid), service);
            //    tracingService.Trace("createdon");
            //}

            tracingService.Trace("DataCollections method  end");
            return entity;
        }
        private static IntegrationConfiguration GetIntegrationConfiguration(IOrganizationService _service)
        {
            try
            {
                IntegrationConfiguration inconfig = new IntegrationConfiguration();
                QueryExpression qsCType = new QueryExpression("hil_integrationconfiguration");
                qsCType.ColumnSet = new ColumnSet("hil_url", "hil_username", "hil_password");
                qsCType.NoLock = true;
                qsCType.Criteria = new FilterExpression(LogicalOperator.And);
                qsCType.Criteria.AddCondition("hil_name", ConditionOperator.Equal, "ArchivedJobAPI");
                Entity integrationConfiguration = _service.RetrieveMultiple(qsCType)[0];
                inconfig.url = integrationConfiguration.GetAttributeValue<string>("hil_url");
                inconfig.userName = integrationConfiguration.GetAttributeValue<string>("hil_username");
                inconfig.password = integrationConfiguration.GetAttributeValue<string>("hil_password");
                return inconfig;
            }
            catch (Exception ex)
            {
                throw new Exception("Error : " + ex.Message);
            }
        }
        private static DateTime dateFormater(string date, ITracingService tracingService)
        {
            //2021-02-11T16:36:52
            tracingService.Trace("date " + date);
            string[] dateTime = date.Split('T');
            string[] dateOnly = dateTime[0].Split('-');
            string[] timeOnly = dateTime[1].Split(':');
            DateTime myDate = DateTime.ParseExact(dateOnly[0] + "-" +
                dateOnly[1] + "-" +
                dateOnly[2] + " " +
                timeOnly[0] + ":" +
                timeOnly[1] + ":" +
                timeOnly[2], "yyyy-MM-dd HH:mm:ss",
                System.Globalization.CultureInfo.InvariantCulture);
            tracingService.Trace("myDate " + myDate);
            return myDate;
        }
        private static EntityReference createEntityRef(String entName, Guid guid, IOrganizationService _service)
        {
            try
            {
                Entity entity = _service.Retrieve(entName, guid, new ColumnSet(false));
                return entity.ToEntityReference();
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("Error createEntityRef ; " + ex.Message);
            }
        }
    }
}
