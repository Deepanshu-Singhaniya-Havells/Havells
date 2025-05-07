using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.IO;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;
using WorkOrderVirtualEntity.Models;

namespace WorkOrderVirtualEntity
{
    public class CommonLib
    {
        public static EntityCollection getData(IOrganizationService service, string eurl, String entName)
        {
            EntityCollection entcol = new EntityCollection();

            Console.WriteLine("GetTestData Function Strated");
            String response = string.Empty;
            IntegrationConfiguration inconfig = GetIntegrationConfiguration(service);
            String sUrl = inconfig.url + eurl;
            String sUserName = inconfig.userName;
            string sPassword = inconfig.password;
            Console.WriteLine("sUrl " + sUrl);

            Console.WriteLine("sUserName " + sUserName);

            Console.WriteLine("sPassword " + sPassword);

            WebClient webClient = new WebClient();
            string credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes(sUserName + ":" + sPassword));
            webClient.Headers[HttpRequestHeader.Authorization] = "Basic " + credentials;
            var jsonData = webClient.DownloadData(sUrl);
            Console.WriteLine("Data ******************************************");
            Console.WriteLine(jsonData.ToString());
            Console.WriteLine("******************************************");

            if (entName == "hil_jobincidentarchived")
            {
                Console.WriteLine("hil_jobincidentarchived");
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(IncidentResponse));
                IncidentResponse rootObject = (IncidentResponse)ser.ReadObject(new MemoryStream(jsonData));
                Console.WriteLine("rootObject.success " + rootObject.success);
                if (rootObject.success)
                {
                    Console.WriteLine("rootObject.workOrders.Count:- " + rootObject.incident.Count);
                    foreach (dtoIncidents _incident in rootObject.incident)
                    {
                        Entity job = DataCollectionsforIncident(service, tracingService, _incident);
                        entcol.Entities.AddRange(job);
                        Console.WriteLine("entcol.Entities.Count :-   " + entcol.Entities.Count);
                    }
                }
            }
            else if (entName == "hil_jobsarchived")
            {
                Console.WriteLine("hil_jobsarchived");
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(ArchivedJobsResponse));
                ArchivedJobsResponse rootObject = (ArchivedJobsResponse)ser.ReadObject(new MemoryStream(jsonData));
                Console.WriteLine("rootObject.success " + rootObject.success);
                if (rootObject.success)
                {
                    Console.WriteLine("rootObject.workOrders.Count:- " + rootObject.workOrders.Count);
                    foreach (dtoWorkOrders _job in rootObject.workOrders)
                    {
                        Entity job = DataCollectionsforJob(service, tracingService, _job);
                        entcol.Entities.AddRange(job);
                    }
                }
            }
            else if (entName == "hil_jobproductarchived")
            {
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(WorkOrderProductResponse));
                WorkOrderProductResponse rootObject = (WorkOrderProductResponse)ser.ReadObject(new MemoryStream(jsonData));
                Console.WriteLine("rootObject.success " + rootObject.success);
                if (rootObject.success)
                {
                    Console.WriteLine("rootObject.orderService.Count:- " + rootObject.incident.Count);
                    foreach (dtoWorkorderProduct _serviceEnt in rootObject.incident)
                    {
                        Entity job = DataCollectionsforProduct(service, tracingService, _serviceEnt);
                        entcol.Entities.AddRange(job);
                    }
                }
            }
            else if (entName == "hil_jobservicearchived")
            {
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(WorkOrderServiceResponse));
                WorkOrderServiceResponse rootObject = (WorkOrderServiceResponse)ser.ReadObject(new MemoryStream(jsonData));
                Console.WriteLine("rootObject.success " + rootObject.success);
                if (rootObject.success)
                {
                    Console.WriteLine("rootObject.orderService.Count:- " + rootObject.orderService.Count);
                    foreach (dtoWorkOrderService _serviceEnt in rootObject.orderService)
                    {
                        Entity job = DataCollectionsforService(service, tracingService, _serviceEnt);
                        entcol.Entities.AddRange(job);
                    }
                }

            }

            Console.WriteLine("entcol.Entities.Count :-   " + entcol.Entities.Count);
            return entcol;
        }
        private static Entity DataCollectionsforJob(IOrganizationService service, dtoWorkOrders _job)
        {
            Entity entity = new Entity("hil_jobsarchived");
            Console.WriteLine("DataCollections method  started");
            if (_job.hil_appointmentmandatory != null)
            {
                entity["hil_appointmentmandatory"] = _job.hil_appointmentmandatory == true ? true : false;
                Console.WriteLine("hil_appointmentmandatory");
            }
            //if (_job.hil_archivelog != null)
            //{
            //    entity["hil_archivelog"] = _job.hil_archivelog == true ? true : false;
            //    Console.WriteLine("hil_archivelog");
            //}
            if (_job.hil_assigntome != null)
            {
                entity["hil_assigntome"] = _job.hil_assigntome == true ? true : false;
                Console.WriteLine("hil_assigntome");
            }
            if (_job.hil_estimatecharges != null)
            {
                entity["hil_estimatecharges"] = _job.hil_estimatecharges == true ? true : false;
                Console.WriteLine("hil_estimatecharges");
            }
            //if (_job.hil_cancelticket != null)
            //{
            //    entity["hil_cancelticket"] = _job.hil_cancelticket == true ? true : false;
            //    Console.WriteLine("hil_cancelticket");
            //}
            //if (_job.hil_generateclaim != null)
            //{
            //    entity["hil_generateclaim"] = _job.hil_generateclaim == true ? true : false;
            //        Console.WriteLine("hil_generateclaim");
            //}
            if (_job.hil_closeticket != null)
            {
                entity["hil_closeticket"] = _job.hil_closeticket == true ? true : false;
                Console.WriteLine("hil_closeticket");
            }
            //if (_job.hil_createchildjob != null)
            //{
            //    entity["hil_createchildjob"] = _job.hil_createchildjob == true ? true : false;
            //    Console.WriteLine("hil_createchildjob");
            //}
            if (_job.hil_delinkpo != null)
            {
                entity["hil_delinkpo"] = _job.hil_delinkpo == true ? true : false;
                Console.WriteLine("hil_delinkpo");
            }
            if (_job.hil_flagpo != null)
            {
                entity["hil_flagpo"] = _job.hil_flagpo == true ? true : false;
                Console.WriteLine("hil_flagpo");
            }
            if (_job.msdyn_followuprequired != null)
            {
                entity["msdyn_followuprequired"] = _job.msdyn_followuprequired == true ? true : false;
                Console.WriteLine("msdyn_followuprequired");
            }
            if (_job.hil_ifclosedjob != null)
            {
                entity["hil_ifclosedjob"] = _job.hil_ifclosedjob == true ? true : false;
                Console.WriteLine("hil_ifclosedjob");
            }
            if (_job.hil_ifparametersadded != null)
            {
                entity["hil_ifparametersadded"] = _job.hil_ifparametersadded == true ? true : false;
                Console.WriteLine("hil_ifparametersadded");
            }
            ////if (_job.hil_isauditlogmigrated != null)
            ////{
            ////    entity["hil_isauditlogmigrated"] = _job.hil_isauditlogmigrated == true ? true : false;
            ////    Console.WriteLine("hil_isauditlogmigrated");
            ////}
            if (_job.hil_ischargable != null)
            {
                entity["hil_ischargable"] = _job.hil_ischargable == true ? true : false;
                Console.WriteLine("hil_ischargable");
            }
            if (_job.msdyn_isfollowup != null)
            {
                entity["msdyn_isfollowup"] = _job.msdyn_isfollowup == true ? true : false;
                Console.WriteLine("msdyn_isfollowup");
            }
            //if (_job.hil_isgascharged != null)
            //{
            //    entity["hil_isgascharged"] = _job.hil_isgascharged == true ? true : false;
            //    Console.WriteLine("hil_isgascharged");
            //}
            if (_job.msdyn_ismobile != null)
            {
                entity["msdyn_ismobile"] = _job.msdyn_ismobile == true ? true : false;
                Console.WriteLine("msdyn_ismobile");
            }
            if (_job.hil_isocr != null)
            {
                entity["hil_isocr"] = _job.hil_isocr == true ? true : false;
                Console.WriteLine("hil_isocr");
            }
            if (_job.hil_jobclosemobile != null)
            {
                entity["hil_jobclosemobile"] = _job.hil_jobclosemobile == true ? true : false;
                Console.WriteLine("hil_jobclosemobile");
            }
            if (_job.hil_calculatecharges != null)
            {
                entity["hil_calculatecharges"] = _job.hil_calculatecharges == true ? true : false;
                Console.WriteLine("hil_calculatecharges");
            }
            if (_job.hil_kkgprovided != null)
            {
                entity["hil_kkgprovided"] = _job.hil_kkgprovided == true ? true : false;
                Console.WriteLine("hil_kkgprovided");
            }
            //if (_job.hil_laborinwarranty != null)
            //{
            //    entity["hil_laborinwarranty"] = _job.hil_laborinwarranty == true ? true : false;
            //    Console.WriteLine("hil_laborinwarranty");
            //}
            //if (_job.hil_paymentlinksent != null)
            //{
            //    entity["hil_paymentlinksent"] = _job.hil_paymentlinksent == true ? true : false;
            //    Console.WriteLine("hil_paymentlinksent");
            //}
            if (_job.hil_resendkkg != null)
            {
                entity["hil_resendkkg"] = _job.hil_resendkkg == true ? true : false;
                Console.WriteLine("hil_resendkkg");
            }
            //if (_job.hil_resendpaymentlink != null)
            //{
            //    entity["hil_resendpaymentlink"] = _job.hil_resendpaymentlink == true ? true : false;
            //    Console.WriteLine("hil_resendpaymentlink");
            //}
            //if (_job.hil_reviewforcountryclassification != null)
            //{
            //    entity["hil_reviewforcountryclassification"] = _job.hil_reviewforcountryclassification == true ? true : false;
            //        Console.WriteLine("hil_reviewforcountryclassification");
            //}
            if (_job.hil_sendestimate != null)
            {
                entity["hil_sendestimate"] = _job.hil_sendestimate == true ? true : false;
                Console.WriteLine("hil_sendestimate");
            }
            if (_job.hil_sendtcr != null)
            {
                entity["hil_sendtcr"] = _job.hil_sendtcr == true ? true : false;
                Console.WriteLine("hil_sendtcr");
            }
            //if (_job.hil_sparepartuse != null)
            //{
            //    entity["hil_sparepartuse"] = _job.hil_sparepartuse == true ? true : false;
            //    Console.WriteLine("hil_sparepartuse");
            //}
            if (_job.msdyn_taxable != null)
            {
                entity["msdyn_taxable"] = _job.msdyn_taxable == true ? true : false;
                Console.WriteLine("msdyn_taxable");
            }
            if (_job.hil_workstarted != null)
            {
                entity["hil_workstarted"] = _job.hil_workstarted == true ? true : false;
                Console.WriteLine("hil_workstarted");
            }
            if (_job.hil_appointmentcount_date != null)
            {
                entity["hil_appointmentcount_date"] = dateFormater(_job.hil_appointmentcount_date, tracingService);
                Console.WriteLine("hil_appointmentcount_date");
            }
            if (_job.hil_appointmentseton != null)
            {
                entity["hil_appointmentseton"] = dateFormater(_job.hil_appointmentseton, tracingService);
                Console.WriteLine("hil_appointmentseton");
            }
            if (_job.hil_appointmenttime != null)
            {
                entity["hil_appointmenttime"] = dateFormater(_job.hil_appointmenttime, tracingService);
                Console.WriteLine("hil_appointmenttime");
            }
            if (_job.hil_assignedon != null)
            {
                entity["hil_assignedon"] = dateFormater(_job.hil_assignedon, tracingService);
                Console.WriteLine("hil_assignedon");
            }
            if (_job.hil_assignmentdate != null)
            {
                entity["hil_assignmentdate"] = dateFormater(_job.hil_assignmentdate, tracingService);
                Console.WriteLine("hil_assignmentdate");
            }
            if (_job.hil_chequedate != null)
            {
                entity["hil_chequedate"] = dateFormater(_job.hil_chequedate, tracingService);
                Console.WriteLine("hil_chequedate");
            }
            if (_job.msdyn_timeclosed != null)
            {
                Console.WriteLine("hil_closeticket");
                entity["hil_closedon"] = dateFormater(_job.msdyn_timeclosed, tracingService);
            }
            if (_job.msdyn_timeclosed != null)
            {
                entity["msdyn_timeclosed"] = dateFormater(_job.msdyn_timeclosed, tracingService);
                Console.WriteLine("msdyn_timeclosed");
            }
            //if (_job.msdyn_completedon != null)
            //{
            //    entity["msdyn_completedon"] = dateFormater(_job.msdyn_completedon, tracingService);
            //    Console.WriteLine("msdyn_completedon");
            //}
            if (_job.createdon != null)
            {
                entity["hil_createdon"] = dateFormater(_job.createdon, tracingService);
                Console.WriteLine("hil_createdon");
            }
            if (_job.msdyn_datewindowend != null)
            {
                entity["msdyn_datewindowend"] = dateFormater(_job.msdyn_datewindowend, tracingService);
                Console.WriteLine("msdyn_datewindowend");
            }
            if (_job.msdyn_datewindowstart != null)
            {
                entity["msdyn_datewindowstart"] = dateFormater(_job.msdyn_datewindowstart, tracingService);
                Console.WriteLine("msdyn_datewindowstart");
            }
            if (_job.hil_escalationcount_date != null)
            {
                entity["hil_escalationcount_date"] = dateFormater(_job.hil_escalationcount_date, tracingService);
                Console.WriteLine("hil_escalationcount_date");
            }
            //if (_job.hil_expecteddeliverydate != null)
            //{
            //    entity["hil_expecteddeliverydate"] = dateFormater(_job.hil_expecteddeliverydate, tracingService);
            //    Console.WriteLine("hil_expecteddeliverydate");
            //}
            //if (_job.msdyn_firstarrivedon != null)
            //{
            //    entity["msdyn_firstarrivedon"] = dateFormater(_job.msdyn_firstarrivedon, tracingService);
            //    Console.WriteLine("msdyn_firstarrivedon");
            //}
            if (_job.hil_firstresponseon != null)
            {
                entity["hil_firstresponseon"] = dateFormater(_job.hil_firstresponseon, tracingService);
                Console.WriteLine("hil_firstresponseon");
            }
            if (_job.msdyn_timefrompromised != null)
            {
                entity["msdyn_timefrompromised"] = dateFormater(_job.msdyn_timefrompromised, tracingService);
                Console.WriteLine("msdyn_timefrompromised");
            }
            //if (_job.hil_lastonholdtime != null)
            //{
            //    entity["hil_lastonholdtime"] = dateFormater(_job.hil_lastonholdtime, tracingService);
            //    Console.WriteLine("hil_lastonholdtime");
            //}
            if (_job.hil_lastresponsetime != null)
            {
                entity["hil_lastresponsetime"] = dateFormater(_job.hil_lastresponsetime, tracingService);
                Console.WriteLine("hil_lastresponsetime");
            }
            if (_job.modifiedon != null)
            {
                entity["hil_modifiedon"] = dateFormater(_job.modifiedon, tracingService);
                Console.WriteLine("modifiedon");
            }
            if (_job.hil_pmsdate != null)
            {
                entity["hil_pmsdate"] = dateFormater(_job.hil_pmsdate, tracingService);
                Console.WriteLine("hil_pmsdate");
            }
            if (_job.hil_preferreddate != null)
            {
                entity["hil_preferreddate"] = dateFormater(_job.hil_preferreddate, tracingService);
                Console.WriteLine("hil_preferreddate");
            }
            if (_job.hil_preferredtimeofvisit != null)
            {
                entity["hil_preferredtimeofvisit"] = dateFormater(_job.hil_preferredtimeofvisit, tracingService);
                Console.WriteLine("hil_preferredtimeofvisit");
            }
            if (_job.hil_purchasedate != null)
            {
                entity["hil_purchasedate"] = dateFormater(_job.hil_purchasedate, tracingService);
                Console.WriteLine("hil_purchasedate");
            }
            //if (_job.createdby != null)
            //{
            //    entity["hil_overriddencreatedon"] = dateFormater(_job.hil_purchasedate, tracingService);
            //    Console.WriteLine("hil_purchasedate");
            //}
            if (_job.hil_remaindercount_date != null)
            {
                entity["hil_remaindercount_date"] = dateFormater(_job.hil_remaindercount_date, tracingService);
                Console.WriteLine("hil_remaindercount_date");
            }
            if (_job.hil_slastarton != null)
            {
                entity["hil_slastarton"] = dateFormater(_job.hil_slastarton, tracingService);
                Console.WriteLine("hil_slastarton");
            }
            if (_job.hil_jobclosureon != null)
            {
                entity["hil_jobclosureon"] = dateFormater(_job.hil_jobclosureon, tracingService);
                Console.WriteLine("hil_jobclosureon");
            }
            if (_job.msdyn_timetopromised != null)
            {
                entity["msdyn_timetopromised"] = dateFormater(_job.msdyn_timetopromised, tracingService);
                Console.WriteLine("msdyn_timetopromised");
            }
            if (_job.msdyn_timewindowend != null)
            {
                entity["msdyn_timewindowend"] = dateFormater(_job.msdyn_timewindowend, tracingService);
                Console.WriteLine("msdyn_timewindowend");
            }
            if (_job.msdyn_timewindowstart != null)
            {
                entity["msdyn_timewindowstart"] = dateFormater(_job.msdyn_timewindowstart, tracingService);
                Console.WriteLine("msdyn_timewindowstart");
            }
            //if (_job.createdby != null)
            //{
            //    entity["hil_warrantydate"] = dateFormater(_job.hil_warrantydate, tracingService);
            //    Console.WriteLine("hil_warrantydate");
            //}
            if (_job.hil_jobclosuredon != null)
            {
                entity["hil_jobclosuredon"] = dateFormater(_job.hil_jobclosuredon, tracingService);
                Console.WriteLine("hil_jobclosuredon");
            }
            if (_job.hil_jobclosuredon != null)
            {
                entity["hil_tattime"] = dateFormater(_job.hil_jobclosuredon, tracingService);
                Console.WriteLine("hil_jobclosuredon");
            }
            //if (_job.hil_balanceamount != null)
            //{
            //    entity["hil_balanceamount"] = Convert.ToDecimal(_job.hil_balanceamount);
            //    Console.WriteLine("hil_balanceamount");
            //}
            if (_job.exchangerate != null)
            {
                entity["hil_exchangerate"] = Convert.ToDecimal(_job.exchangerate);
                Console.WriteLine("exchangerate");
            }
            if (_job.hil_maxquantity != null)
            {
                entity["hil_maxquantity"] = Convert.ToDecimal(_job.hil_maxquantity);
                Console.WriteLine("hil_maxquantity");
            }
            if (_job.hil_minquantity != null)
            {
                entity["hil_minquantity"] = Convert.ToDecimal(_job.hil_minquantity);
                Console.WriteLine("hil_minquantity");
            }
            if (_job.hil_receiptamount != null)
            {
                entity["hil_receiptamount"] = Convert.ToDecimal(_job.hil_receiptamount);
                Console.WriteLine("hil_receiptamount");
            }
            //if (_job.hil_tattimecalculated != null)
            //{
            //    entity["hil_tattimecalculated"] = Convert.ToDecimal(_job.hil_tattimecalculated);
            //    Console.WriteLine("hil_tattimecalculated");
            //}
            if (_job.hil_timeinappointmentdate != null)
            {
                entity["hil_timeinappointmentdate"] = Convert.ToDecimal(_job.hil_timeinappointmentdate);
                Console.WriteLine("hil_timeinappointmentdate");
            }
            if (_job.hil_timeincurrentdate != null)
            {
                entity["hil_timeincurrentdate"] = Convert.ToDecimal(_job.hil_timeincurrentdate);
                Console.WriteLine("hil_timeincurrentdate");
            }
            if (_job.hil_timeinfirstresponse != null)
            {
                entity["hil_timeinfirstresponse"] = Convert.ToDecimal(_job.hil_timeinfirstresponse);
                Console.WriteLine("hil_timeinfirstresponse");
            }
            if (_job.hil_timeinjobclosure != null)
            {
                entity["hil_timeinjobclosure"] = Convert.ToDecimal(_job.hil_timeinjobclosure);
                Console.WriteLine("hil_timeinjobclosure");
            }
            if (_job.hil_estimatechargestotal != null)
            {
                entity["hil_estimatechargestotal"] = Convert.ToDecimal(_job.hil_estimatechargestotal);
                Console.WriteLine("hil_estimatechargestotal");
            }
            if (_job.hil_estimatedchargedecimal != null)
            {
                entity["hil_estimatedchargedecimal"] = Convert.ToDecimal(_job.hil_estimatedchargedecimal);
                Console.WriteLine("hil_estimatedchargedecimal");
            }
            if (_job.hil_payblechargedecimal != null)
            {
                entity["hil_payblechargedecimal"] = Convert.ToDecimal(_job.hil_payblechargedecimal);
                Console.WriteLine("hil_payblechargedecimal");
            }
            if (_job.createdby != null)
            {
                entity["hil_actualcharges"] = Convert.ToDecimal(_job.hil_actualcharges);
                Console.WriteLine("hil_actualcharges");
            }
            if (_job.hil_actualcharges_base != null)
            {
                entity["hil_actualcharges_base"] = Convert.ToDecimal(_job.hil_actualcharges_base);
                Console.WriteLine("hil_actualcharges_base");
            }
            //if (_job.hil_balanceamount_base != null)
            //{
            //    entity["hil_balanceamount_base"] = Convert.ToDecimal(_job.hil_balanceamount_base);
            //    Console.WriteLine("hil_balanceamount_base");
            //}
            //if (_job.hil_claimamount != null)
            //{
            //    entity["hil_claimamount"]  = Convert.ToDecimal(_job.hil_claimamount);
            //    Console.WriteLine("hil_claimamount");
            //}
            //           //if (_job.hil_claimamount_base != null)
            //           //{
            //           //    entity["hil_claimamount_base"] = Convert.ToDecimal(_job.hil_claimamount_base);
            //           //    Console.WriteLine("hil_claimamount_base");
            //           //}
            //           if (_job.createdby != null)
            //           {
            //               entity["hil_estimatechargestotal_base"] =
            //}
            if (_job.msdyn_estimatesubtotalamount != null)
            {
                entity["msdyn_estimatesubtotalamount"] = Convert.ToDecimal(_job.msdyn_estimatesubtotalamount);
                Console.WriteLine("msdyn_estimatesubtotalamount");
            }
            //           if (_job.createdby != null)
            //           {
            //               entity["msdyn_estimatesubtotalamount_base"] =
            //}
            if (_job.msdyn_latitude != null)
            {
                entity["msdyn_latitude"] = Convert.ToDecimal(_job.msdyn_latitude);
                Console.WriteLine("msdyn_latitude");
            }
            if (_job.msdyn_longitude != null)
            {
                entity["msdyn_longitude"] = Convert.ToDecimal(_job.msdyn_longitude);
                Console.WriteLine("msdyn_longitude");
            }
            //if (_job.createdby != null)
            //{
            //    entity["hil_receiptamount_base"] = Convert.ToDecimal(_job.msdyn_longitude);
            //    Console.WriteLine("msdyn_longitude");
            //}
            if (_job.msdyn_subtotalamount != null)
            {
                entity["msdyn_subtotalamount"] = Convert.ToDecimal(_job.msdyn_subtotalamount);
                Console.WriteLine("msdyn_subtotalamount");
            }
            //if (_job.createdby != null)
            //{
            //    entity["msdyn_subtotalamount_base"] = Convert.ToDecimal(_job.msdyn_subtotalamount);
            //    Console.WriteLine("msdyn_subtotalamount");
            //}
            if (_job.msdyn_totalamount != null)
            {
                entity["msdyn_totalamount"] = Convert.ToDecimal(_job.msdyn_totalamount);
                Console.WriteLine("msdyn_totalamount");
            }
            //           if (_job.createdby != null)
            //           {
            //               entity["msdyn_totalamount_base"] =
            //}
            if (_job.msdyn_totalsalestax != null)
            {
                entity["msdyn_totalsalestax"] = Convert.ToDecimal(_job.msdyn_totalsalestax);
                Console.WriteLine("msdyn_totalsalestax");
            }
            //if (_job.createdby != null)
            //{
            //    entity["msdyn_totalsalestax_base"] = Convert.ToDecimal(_job.msdyn_totalsalestax);
            //    Console.WriteLine("msdyn_totalsalestax");
            //}
            if (_job.hil_ageing != null)
            {
                entity["hil_ageing"] = Convert.ToInt16(_job.hil_ageing);
                Console.WriteLine("hil_ageing");
            }
            if (_job.hil_appointmentcount != null)
            {
                entity["hil_appointmentcount"] = Convert.ToInt16(_job.hil_appointmentcount);
                Console.WriteLine("hil_appointmentcount");
            }
            if (_job.hil_appointmentcount_state != null)
            {
                entity["hil_appointmentcount_state"] = Convert.ToInt16(_job.hil_appointmentcount_state);
                Console.WriteLine("hil_appointmentcount_state");
            }
            if (_job.hil_bucket_ageing != null)
            {
                entity["hil_bucket_ageing"] = Convert.ToInt16(_job.hil_bucket_ageing);
                Console.WriteLine("hil_bucket_ageing");
            }
            if (_job.msdyn_childindex != null)
            {
                entity["msdyn_childindex"] = Convert.ToInt16(_job.msdyn_childindex);
                Console.WriteLine("msdyn_childindex");
            }
            if (_job.hil_escalationcount != null)
            {
                entity["hil_escalationcount"] = Convert.ToInt16(_job.hil_escalationcount);
                Console.WriteLine("hil_escalationcount");
            }
            if (_job.hil_escalationcount_state != null)
            {
                entity["hil_escalationcount_state"] = Convert.ToInt16(_job.hil_escalationcount_state);
                Console.WriteLine("hil_escalationcount_state");
            }
            if (_job.hil_escallationcountinteger != null)
            {
                entity["hil_escallationcountinteger"] = Convert.ToInt16(_job.hil_escallationcountinteger);
                Console.WriteLine("hil_escallationcountinteger");
            }
            if (_job.importsequencenumber != null)
            {
                entity["hil_importsequencenumber"] = Convert.ToInt16(_job.importsequencenumber);
                Console.WriteLine("hil_escallationcountintehil_importsequencenumberger");
            }
            if (_job.hil_incidentquantity != null)
            {
                entity["hil_incidentquantity"] = Convert.ToInt16(_job.hil_incidentquantity);
                Console.WriteLine("hil_incidentquantity");
            }
            //if (_job.hil_jobclosurecounter != null)
            //{
            //    entity["hil_jobclosurecounter"] = Convert.ToInt16(_job.hil_jobclosurecounter);
            //    Console.WriteLine("hil_jobclosurecounter");
            //}
            if (_job.hil_jobpriority != null)
            {
                entity["hil_jobpriority"] = Convert.ToInt16(_job.hil_jobpriority);
                Console.WriteLine("hil_jobpriority");
            }
            if (_job.hil_jobincidentcount != null)
            {
                entity["hil_jobincidentcount"] = Convert.ToInt16(_job.hil_jobincidentcount);
                Console.WriteLine("hil_jobincidentcount");
            }
            //if (_job.hil_onholdtime != null)
            //{
            //    entity["hil_onholdtime"] = Convert.ToInt16(_job.hil_onholdtime);
            //    Console.WriteLine("hil_onholdtime");
            //}
            if (_job.hil_pendingquantity != null)
            {
                entity["hil_pendingquantity"] = Convert.ToInt16(_job.hil_pendingquantity);
                Console.WriteLine("hil_pendingquantity");
            }
            if (_job.hil_phonecallcount != null)
            {
                entity["hil_phonecallcount"] = Convert.ToInt16(_job.hil_phonecallcount);
                Console.WriteLine("hil_phonecallcount");
            }
            if (_job.createdby != null)
            {
                entity["hil_pmscount"] = Convert.ToInt16(_job.hil_phonecallcount);
                Console.WriteLine("hil_phonecallcount");
            }
            if (_job.msdyn_primaryincidentestimatedduration != null)
            {
                entity["msdyn_primaryincidentestimatedduration"] = Convert.ToInt16(_job.msdyn_primaryincidentestimatedduration);
                Console.WriteLine("msdyn_primaryincidentestimatedduration");
            }
            if (_job.hil_quantity != null)
            {
                entity["hil_quantity"] = Convert.ToInt16(_job.hil_quantity);
                Console.WriteLine("hil_quantity");
            }
            if (_job.hil_quantityofunits != null)
            {
                entity["hil_quantityofunits"] = Convert.ToInt16(_job.hil_quantityofunits);
                Console.WriteLine("hil_quantityofunits");
            }
            if (_job.hil_remaindercountinteger != null)
            {
                entity["hil_remaindercountinteger"] = Convert.ToInt16(_job.hil_remaindercountinteger);
                Console.WriteLine("hil_remaindercountinteger");
            }
            if (_job.hil_remaindercount != null)
            {
                entity["hil_remaindercount"] = Convert.ToInt16(_job.hil_remaindercount);
                Console.WriteLine("hil_remaindercount");
            }
            if (_job.hil_remaindercount_state != null)
            {
                entity["hil_remaindercount_state"] = Convert.ToInt16(_job.hil_remaindercount_state);
                Console.WriteLine("hil_remaindercount_state");
            }
            //if (_job.hil_timezoneruleversionnumber != null)
            //{
            //    entity["hil_timezoneruleversionnumber"] = Convert.ToInt16(_job.hil_timezoneruleversionnumber);
            //    Console.WriteLine("hil_timezoneruleversionnumber");
            //}
            //if (_job.msdyn_totalestimatedduration != null)
            //{
            //    entity["msdyn_totalestimatedduration"] = Convert.ToInt16(_job.msdyn_totalestimatedduration);
            //    Console.WriteLine("msdyn_totalestimatedduration");
            //}
            //if (_job.hil_utcconversiontimezonecode != null)
            //{
            //    entity["hil_utcconversiontimezonecode"] = Convert.ToInt16(_job.hil_utcconversiontimezonecode);
            //    Console.WriteLine("hil_utcconversiontimezonecode");
            //}
            if (_job.hil_addressdetails != null)
            {
                entity["hil_addressdetails"] = _job.hil_addressdetails;
                Console.WriteLine("hil_addressdetails");
            }
            if (_job.msdyn_bookingsummary != null)
            {
                entity["msdyn_bookingsummary"] = _job.msdyn_bookingsummary;
                Console.WriteLine("msdyn_bookingsummary");
            }
            if (_job.hil_closureremarks != null)
            {
                entity["hil_closureremarks"] = _job.hil_closureremarks;
                Console.WriteLine("hil_closureremarks");
            }
            if (_job.hil_customercomplaintdescription != null)
            {
                entity["hil_customercomplaintdescription"] = _job.hil_customercomplaintdescription;
                Console.WriteLine("hil_customercomplaintdescription");
            }
            if (_job.msdyn_primaryincidentdescription != null)
            {
                entity["msdyn_primaryincidentdescription"] = _job.msdyn_primaryincidentdescription;
                Console.WriteLine("msdyn_primaryincidentdescription");
            }
            if (_job.hil_emergencyremarks != null)
            {
                entity["hil_emergencyremarks"] = _job.hil_emergencyremarks;
                Console.WriteLine("hil_emergencyremarks");
            }
            if (_job.msdyn_followupnote != null)
            {
                entity["msdyn_followupnote"] = _job.msdyn_followupnote;
                Console.WriteLine("msdyn_followupnote");
            }
            if (_job.hil_fulladdress != null)
            {
                entity["hil_fulladdress"] = _job.hil_fulladdress;
                Console.WriteLine("hil_fulladdress");
            }
            if (_job.msdyn_instructions != null)
            {
                entity["msdyn_instructions"] = _job.msdyn_instructions;
                Console.WriteLine("msdyn_instructions");
            }
            if (_job.msdyn_internalflags != null)
            {
                entity["msdyn_internalflags"] = _job.msdyn_internalflags;
                Console.WriteLine("msdyn_internalflags");
            }
            if (_job.hil_reportbinary != null)
            {
                entity["hil_reportbinary"] = _job.hil_reportbinary;
                Console.WriteLine("hil_reportbinary");
            }
            if (_job.hil_reporttext != null)
            {
                entity["hil_reporttext"] = _job.hil_reporttext;
                Console.WriteLine("hil_reporttext");
            }
            if (_job.hil_systemremarks != null)
            {
                entity["hil_systemremarks"] = _job.hil_systemremarks;
                Console.WriteLine("hil_systemremarks");
            }
            //if (_job.hil_termsconditions != null)
            //{
            //    entity["hil_termsconditions"] = _job.hil_termsconditions;
            //    Console.WriteLine("hil_termsconditions");
            //}
            if (_job.msdyn_workordersummary != null)
            {
                entity["msdyn_workordersummary"] = _job.msdyn_workordersummary;
                Console.WriteLine("msdyn_workordersummary");
            }
            //if (_job.hil_additionalaction != null)
            //{
            //    entity["hil_additionalaction"] = new OptionSetValue(int.Parse(_job.hil_additionalaction));
            //    Console.WriteLine("hil_additionalaction");
            //}
            //if (_job.hil_amcsyncstatus != null)
            //{
            //    entity["hil_amcsyncstatus"] = new OptionSetValue(int.Parse(_job.hil_amcsyncstatus));
            //    Console.WriteLine("hil_amcsyncstatus");
            //}
            if (_job.hil_appointmentstatus != null)
            {
                entity["hil_appointmentstatus"] = new OptionSetValue(int.Parse(_job.hil_appointmentstatus.ToString()));
                Console.WriteLine("hil_appointmentstatus");
            }
            if (_job.hil_assigntobranchhead != null)
            {
                entity["hil_assigntobranchhead"] = new OptionSetValue(int.Parse(_job.hil_assigntobranchhead.ToString()));
                Console.WriteLine("hil_assigntobranchhead");
            }
            if (_job.hil_automaticassign != null)
            {
                entity["hil_automaticassign"] = new OptionSetValue(int.Parse(_job.hil_automaticassign.ToString()));
                Console.WriteLine("hil_automaticassign");
            }
            if (_job.hil_branchheadapproval != null)
            {
                entity["hil_branchheadapproval"] = new OptionSetValue(int.Parse(_job.hil_branchheadapproval.ToString()));
                Console.WriteLine("hil_branchheadapproval");
            }
            if (_job.hil_brand != null)
            {
                entity["hil_brand"] = new OptionSetValue(int.Parse(_job.hil_brand.ToString()));
                Console.WriteLine("hil_brand");
            }
            //if (_job.hil_callcenter != null)
            //{
            //    entity["hil_callcenter"] = new OptionSetValue(int.Parse(_job.hil_callcenter.ToString()));
            //    Console.WriteLine("hil_callcenter");
            //}
            if (_job.hil_callertype != null)
            {
                entity["hil_callertype"] = new OptionSetValue(int.Parse(_job.hil_callertype.ToString()));
                Console.WriteLine("hil_callertype");
            }
            //if (_job.hil_channelpartnercategory != null)
            //{
            //    entity["hil_channelpartnercategory"] = new OptionSetValue(int.Parse(_job.hil_channelpartnercategory.ToString()));
            //    Console.WriteLine("hil_channelpartnercategory");
            //}
            //if (_job.hil_claimstatus != null)
            //{
            //    entity["hil_claimstatus"] = new OptionSetValue(int.Parse(_job.hil_claimstatus.ToString()));
            //    Console.WriteLine("hil_claimstatus");
            //}
            if (_job.hil_countryclassification != null)
            {
                entity["hil_countryclassification"] = new OptionSetValue(int.Parse(_job.hil_countryclassification.ToString()));
                Console.WriteLine("hil_countryclassification");
            }
            //if (_job.hil_customercarenumber != null)
            //{
            //    entity["hil_customercarenumber"] = new OptionSetValue(int.Parse(_job.hil_customercarenumber.ToString()));
            //    Console.WriteLine("hil_customercarenumber");
            //}
            if (_job.hil_customerfeedback != null)
            {
                entity["hil_customerfeedback"] = new OptionSetValue(int.Parse(_job.hil_customerfeedback.ToString()));
                Console.WriteLine("hil_customerfeedback");
            }
            if (_job.hil_emergencycall != null)
            {
                entity["hil_emergencycall"] = new OptionSetValue(int.Parse(_job.hil_emergencycall.ToString()));
                Console.WriteLine("hil_emergencycall");
            }
            if (_job.hil_escalationcall != null)
            {
                entity["hil_escalationcall"] = new OptionSetValue(int.Parse(_job.hil_escalationcall.ToString()));
                Console.WriteLine("hil_escalationcall");
            }
            if (_job.hil_ifamceligible != null)
            {
                entity["hil_ifamceligible"] = new OptionSetValue(int.Parse(_job.hil_ifamceligible.ToString()));
                Console.WriteLine("hil_ifamceligible");
            }
            if (_job.hil_ischargeable != null)
            {
                entity["hil_ischargeable"] = new OptionSetValue(int.Parse(_job.hil_ischargeable.ToString()));
                Console.WriteLine("hil_ischargeable");
            }
            if (_job.hil_isclaimable != null)
            {
                entity["hil_isclaimable"] = new OptionSetValue(int.Parse(_job.hil_isclaimable.ToString()));
                Console.WriteLine("hil_isclaimable");
            }
            if (_job.hil_iswrongjobclosure != null)
            {
                entity["hil_iswrongjobclosure"] = new OptionSetValue(int.Parse(_job.hil_iswrongjobclosure.ToString()));
                Console.WriteLine("hil_iswrongjobclosure");
            }
            if (_job.hil_jobcancelreason != null)
            {
                entity["hil_jobcancelreason"] = new OptionSetValue(int.Parse(_job.hil_jobcancelreason.ToString()));
                Console.WriteLine("hil_jobcancelreason");
            }
            if (_job.hil_jobclass != null)
            {
                entity["hil_jobclass"] = new OptionSetValue(int.Parse(_job.hil_jobclass.ToString()));
                Console.WriteLine("hil_jobclass");
            }
            if (_job.hil_jobclassapproval != null)
            {
                entity["hil_jobclassapproval"] = new OptionSetValue(int.Parse(_job.hil_jobclassapproval.ToString()));
                Console.WriteLine("hil_jobclassapproval");
            }
            //if (_job.hil_jobreopenreason != null)
            //{
            //    entity["hil_jobreopenreason"] = new OptionSetValue(int.Parse(_job.hil_jobreopenreason.ToString()));
            //    Console.WriteLine("hil_jobreopenreason");
            //}
            if (_job.hil_kkgcode_sms != null)
            {
                entity["hil_kkgcode_sms"] = new OptionSetValue(int.Parse(_job.hil_kkgcode_sms.ToString()));
                Console.WriteLine("hil_kkgcode_sms");
            }
            if (_job.hil_locationofasset != null)
            {
                entity["hil_locationofasset"] = new OptionSetValue(int.Parse(_job.hil_locationofasset.ToString()));
                Console.WriteLine("hil_locationofasset");
            }
            if (_job.hil_modeofpayment != null)
            {
                entity["hil_modeofpayment"] = new OptionSetValue(int.Parse(_job.hil_modeofpayment.ToString()));
                Console.WriteLine("hil_modeofpayment");
            }
            //if (_job.hil_paymentstatus != null)
            //{
            //    entity["hil_paymentstatus"] = new OptionSetValue(int.Parse(_job.hil_paymentstatus.ToString()));
            //    Console.WriteLine("hil_paymentstatus");
            //}
            if (_job.hil_preferredtime != null)
            {
                entity["hil_preferredtime"] = new OptionSetValue(int.Parse(_job.hil_preferredtime.ToString()));
                Console.WriteLine("hil_preferredtime");
            }
            if (_job.hil_productreplacement != null)
            {
                entity["hil_productreplacement"] = new OptionSetValue(int.Parse(_job.hil_productreplacement.ToString()));
                Console.WriteLine("hil_productreplacement");
            }
            if (_job.hil_remindercall != null)
            {
                entity["hil_remindercall"] = new OptionSetValue(int.Parse(_job.hil_remindercall.ToString()));
                Console.WriteLine("hil_remindercall");
            }
            if (_job.hil_repairdone != null)
            {
                entity["hil_repairdone"] = new OptionSetValue(int.Parse(_job.hil_repairdone.ToString()));
                Console.WriteLine("hil_repairdone");
            }
            if (_job.hil_returned != null)
            {
                entity["hil_returned"] = new OptionSetValue(int.Parse(_job.hil_returned.ToString()));
                Console.WriteLine("hil_returned");
            }
            if (_job.hil_salutation != null)
            {
                entity["hil_salutation"] = new OptionSetValue(int.Parse(_job.hil_salutation.ToString()));
                Console.WriteLine("hil_returned");
            }
            if (_job.hil_slastatus != null)
            {
                entity["hil_slastatus"] = new OptionSetValue(int.Parse(_job.hil_slastatus.ToString()));
                Console.WriteLine("hil_slastatus");
            }
            if (_job.hil_sourceofjob != null)
            {
                entity["hil_sourceofjob"] = new OptionSetValue(int.Parse(_job.hil_sourceofjob.ToString()));
                Console.WriteLine("hil_sourceofjob");
            }
            if (_job.hil_requesttype != null)
            {
                entity["hil_requesttype"] = new OptionSetValue(int.Parse(_job.hil_requesttype.ToString()));
                Console.WriteLine("hil_requesttype");
            }
            if (_job.msdyn_systemstatus != null)
            {
                entity["msdyn_systemstatus"] = new OptionSetValue(int.Parse(_job.msdyn_systemstatus.ToString()));
                Console.WriteLine("msdyn_systemstatus");
            }
            if (_job.hil_warrantystatus != null)
            {
                entity["hil_warrantystatus"] = new OptionSetValue(int.Parse(_job.hil_warrantystatus.ToString()));
                Console.WriteLine("hil_warrantystatus");
            }
            if (_job.msdyn_worklocation != null)
            {
                entity["msdyn_worklocation"] = new OptionSetValue(int.Parse(_job.msdyn_worklocation.ToString()));
                Console.WriteLine("msdyn_worklocation");
            }
            if (_job.msdyn_addressname != null)
            {
                entity["msdyn_addressname"] = _job.msdyn_addressname;
                Console.WriteLine("msdyn_addressname");
            }
            if (_job.hil_alternate != null)
            {
                entity["hil_alternate"] = _job.hil_alternate;
                Console.WriteLine("hil_alternate");
            }
            if (_job.hil_areatext != null)
            {
                entity["hil_areatext"] = _job.hil_areatext;
                Console.WriteLine("hil_areatext");
            }
            //if (_job.msdyn_autonumbering != null)
            //{
            //    entity["msdyn_autonumbering"] = _job.msdyn_autonumbering;
            //    Console.WriteLine("msdyn_autonumbering");
            //}
            //if (_job.hil_b2bcompanyname != null)
            //{
            //    entity["hil_b2bcompanyname"] = _job.hil_b2bcompanyname;
            //    Console.WriteLine("hil_b2bcompanyname");
            //}
            if (_job.hil_branchtext != null)
            {
                entity["hil_branchtext"] = _job.hil_branchtext;
                Console.WriteLine("msdyn_autonumbering");
            }
            if (_job.hil_branchengineercity != null)
            {
                entity["hil_branchengineercity"] = _job.hil_branchengineercity;
                Console.WriteLine("hil_branchengineercity");
            }
            if (_job.hil_callingnumber != null)
            {
                entity["hil_callingnumber"] = _job.hil_callingnumber;
                Console.WriteLine("hil_callingnumber");
            }
            if (_job.hil_chequenumber != null)
            {
                entity["hil_chequenumber"] = _job.hil_chequenumber;
                Console.WriteLine("hil_chequenumber");
            }
            if (_job.msdyn_city != null)
            {
                entity["msdyn_city"] = _job.msdyn_city;
                Console.WriteLine("msdyn_city");
            }
            if (_job.hil_citytext != null)
            {
                entity["hil_citytext"] = _job.hil_citytext;
                Console.WriteLine("hil_citytext");
            }
            if (_job.msdyn_country != null)
            {
                entity["msdyn_country"] = _job.msdyn_country;
                Console.WriteLine("msdyn_country");
            }
            if (_job.createdbyname != null)
            {
                entity["hil_createdby"] = _job.createdbyname;
                Console.WriteLine("createdbyname");
            }
            //if (_job.createdby != null)
            //{
            //    entity["hil_createdonbehalfby"] = _job.createdbyname;
            //    Console.WriteLine("createdbyname");
            //}
            if (_job.hil_customername != null)
            {
                entity["hil_customername"] = _job.hil_customername;
                Console.WriteLine("hil_customername");
            }
            //if (_job.msdyn_phonenumber != null)
            //{
            //    entity["msdyn_phonenumber"] = _job.msdyn_phonenumber;
            //    Console.WriteLine("msdyn_phonenumber");
            //}
            //if (_job.hil_prno != null)
            //{
            //    entity["hil_prno"] = _job.hil_prno;
            //    Console.WriteLine("msdyn_phonenumber");
            //}
            //if (_job.hil_discountcouponcode != null)
            //{
            //    entity["hil_discountcouponcode"] = _job.hil_discountcouponcode;
            //    Console.WriteLine("hil_discountcouponcode");
            //}
            if (_job.hil_districttext != null)
            {
                entity["hil_districttext"] = _job.hil_districttext;
                Console.WriteLine("hil_districttext");
            }
            if (_job.hil_email != null)
            {
                entity["hil_email"] = _job.hil_email;
                Console.WriteLine("hil_email");
            }
            if (_job.hil_emolpyeenamecode != null)
            {
                entity["hil_emolpyeenamecode"] = _job.hil_emolpyeenamecode;
                Console.WriteLine("hil_emolpyeenamecode");
            }
            if (_job.hil_newserialnumber != null)
            {
                entity["hil_newserialnumber"] = _job.hil_newserialnumber;
                Console.WriteLine("hil_newserialnumber");
            }
            if (_job.msdyn_name != null)
            {
                entity["hil_name"] = _job.msdyn_name;
                Console.WriteLine("hil_name");
            }
            if (_job.createdby != null)
            {
                entity["msdyn_name"] = _job.msdyn_name;
                Console.WriteLine("hil_name");
            }
            if (_job.hil_jobstatuscode != null)
            {
                entity["hil_jobstatuscode"] = _job.hil_jobstatuscode;
                Console.WriteLine("hil_jobstatuscode");
            }
            if (_job.hil_kkgcode != null)
            {
                entity["hil_kkgcode"] = _job.hil_kkgcode;
                Console.WriteLine("hil_kkgcode");
            }
            if (_job.hil_kkgotp != null)
            {
                entity["hil_kkgotp"] = _job.hil_kkgotp;
                Console.WriteLine("hil_kkgotp");
            }
            if (_job.createdby != null)
            {
                entity["hil_mobilenumber"] = _job.hil_mobilenumber;
                Console.WriteLine("hil_mobilenumber");
            }
            if (_job.hil_modelid != null)
            {
                entity["hil_modelid"] = _job.hil_modelid;
                Console.WriteLine("hil_modelid");
            }
            if (_job.hil_modelname != null)
            {
                entity["hil_modelname"] = _job.hil_modelname;
                Console.WriteLine("hil_modelname");
            }
            if (_job.owneridname != null)
            {
                entity["hil_ownerid"] = _job.owneridname;
                Console.WriteLine("hil_ownerid");
            }
            if (_job.hil_pincodetext != null)
            {
                entity["hil_pincodetext"] = _job.hil_pincodetext;
                Console.WriteLine("hil_pincodetext");
            }
            if (_job.msdyn_postalcode != null)
            {
                entity["msdyn_postalcode"] = _job.msdyn_postalcode;
                Console.WriteLine("msdyn_postalcode");
            }
            //if (_job.hil_prnumbertext != null)
            //{
            //    entity["hil_prnumbertext"] = _job.hil_prnumbertext;
            //    Console.WriteLine("hil_prnumbertext");
            //}
            if (_job.hil_productsubcategorycallsubtype != null)
            {
                entity["hil_productsubcategorycallsubtype"] = _job.hil_productsubcategorycallsubtype;
                Console.WriteLine("hil_productsubcategorycallsubtype");
            }
            if (_job.hil_purchasedfrom != null)
            {
                entity["hil_purchasedfrom"] = _job.hil_purchasedfrom;
                Console.WriteLine("hil_purchasedfrom");
            }
            //if (_job.hil_reassignreasonincaseofothers != null)
            //{
            //    entity["hil_reassignreasonincaseofothers"] = _job.hil_reassignreasonincaseofothers;
            //    Console.WriteLine("hil_reassignreasonincaseofothers");
            //}
            if (_job.hil_receiptnumber != null)
            {
                entity["hil_receiptnumber"] = _job.hil_receiptnumber;
                Console.WriteLine("hil_receiptnumber");
            }
            //if (_job.hil_receiptsamountinwords != null)
            //{
            //    entity["hil_receiptsamountinwords"] = _job.hil_receiptsamountinwords;
            //    Console.WriteLine("hil_receiptsamountinwords");
            //}
            if (_job.hil_regiontext != null)
            {
                entity["hil_regiontext"] = _job.hil_regiontext;
                Console.WriteLine("hil_regiontext");
            }
            if (_job.hil_regionbranch != null)
            {
                entity["hil_regionbranch"] = _job.hil_regionbranch;
                Console.WriteLine("hil_regionbranch");
            }
            if (_job.msdyn_stateorprovince != null)
            {
                entity["msdyn_stateorprovince"] = _job.msdyn_stateorprovince;
                Console.WriteLine("msdyn_stateorprovince");
            }
            if (_job.hil_statetext != null)
            {
                entity["hil_statetext"] = _job.hil_statetext;
                Console.WriteLine("hil_statetext");
            }
            if (_job.msdyn_address1 != null)
            {
                entity["msdyn_address1"] = _job.msdyn_address1;
                Console.WriteLine("msdyn_address1");
            }
            if (_job.msdyn_address2 != null)
            {
                entity["msdyn_address2"] = _job.msdyn_address2;
                Console.WriteLine("msdyn_address2");
            }
            if (_job.msdyn_address3 != null)
            {
                entity["msdyn_address3"] = _job.msdyn_address3;
                Console.WriteLine("msdyn_address3");
            }
            if (_job.hil_technicianname != null)
            {
                entity["hil_technicianname"] = _job.hil_technicianname;
                Console.WriteLine("hil_technicianname");
            }
            //if (_job.msdyn_mapcontrol != null)
            //{
            //    entity["msdyn_mapcontrol"] = _job.msdyn_mapcontrol;
            //    Console.WriteLine("msdyn_mapcontrol");
            //}
            if (_job.traversedpath != null)
            {
                entity["hil_traversedpath"] = _job.traversedpath;
                Console.WriteLine("hil_traversedpath");
            }
            if (_job.hil_webclosureremarks != null)
            {
                entity["hil_webclosureremarks"] = _job.hil_webclosureremarks;
                Console.WriteLine("hil_webclosureremarks");
            }
            if (_job.msdyn_workorderid != null)
            {
                Console.WriteLine("msdyn_workorderid");
                entity["hil_jobsarchivedid"] = new Guid(_job.msdyn_workorderid);
            }
            if (_job.hil_address != null)
            {
                entity["hil_address"] = createEntityRef("hil_address", new Guid(_job.hil_address), service);
                Console.WriteLine("hil_address");
            }
            if (_job.msdyn_agreement != null)
            {
                entity["msdyn_agreement"] = createEntityRef("hil_address", new Guid(_job.msdyn_agreement), service);
                Console.WriteLine("msdyn_agreement");
            }
            if (_job.hil_allocatebykpi != null)
            {
                entity["hil_allocatebykpi"] = createEntityRef("slakpiinstance", new Guid(_job.hil_allocatebykpi), service);
                Console.WriteLine("slakpiinstance");
            }
            if (_job.hil_area != null)
            {
                entity["hil_area"] = createEntityRef("hil_area", new Guid(_job.hil_area), service);
                Console.WriteLine("hil_area");
            }
            if (_job.msdyn_customerasset != null)
            {
                entity["hil_customerasset"] = createEntityRef("msdyn_customerasset", new Guid(_job.msdyn_customerasset), service);
                Console.WriteLine("msdyn_customerasset");
            }
            if (_job.msdyn_billingaccount != null)
            {
                entity["hil_billingaccount"] = createEntityRef("account", new Guid(_job.msdyn_billingaccount), service);
                Console.WriteLine("hil_billingaccount");
            }
            if (_job.hil_branch != null)
            {
                entity["hil_branch"] = createEntityRef("hil_branch", new Guid(_job.hil_branch), service);
                Console.WriteLine("hil_branch");
            }
            if (_job.hil_bshname != null)
            {
                entity["hil_bsh"] = _job.hil_bshname;// createEntityRef("hil_branch", new Guid(_job.hil_branch), service);
                Console.WriteLine("hil_bshname");
            }
            if (_job.hil_callsubtype != null)
            {
                entity["hil_callsubtype"] = createEntityRef("hil_callsubtype", new Guid(_job.hil_callsubtype), service);
                Console.WriteLine("hil_calltype");
            }
            if (_job.hil_calltype != null)
            {
                entity["hil_calltype"] = createEntityRef("hil_calltype", new Guid(_job.hil_calltype), service);
                Console.WriteLine("hil_calltype");
            }
            //if (_job.hil_cancelledbyname != null)
            //{
            //    entity["hil_cancelledby"] = _job.hil_cancelledbyname;// createEntityRef("hil_calltype", new Guid(_job.hil_calltype), service);
            //    Console.WriteLine("hil_cancelledbyname");
            //}
            if (_job.msdyn_servicerequest != null)
            {
                entity["hil_servicerequest"] = createEntityRef("incident", new Guid(_job.msdyn_servicerequest), service);
                Console.WriteLine("msdyn_servicerequest");
            }
            if (_job.msdyn_primaryincidenttype != null)
            {
                entity["hil_primaryincidenttype"] = createEntityRef("msdyn_incidenttype", new Guid(_job.msdyn_primaryincidenttype), service);
                Console.WriteLine("msdyn_incidenttype");
            }
            if (_job.hil_city != null)
            {
                entity["hil_city"] = createEntityRef("hil_city", new Guid(_job.hil_city), service);
                Console.WriteLine("hil_city");
            }
            if (_job.hil_claimheader != null)
            {
                entity["hil_claimheader"] = createEntityRef("hil_claimheader", new Guid(_job.hil_claimheader), service);
                Console.WriteLine("hil_claimheader");
            }
            if (_job.hil_claimline != null)
            {
                entity["hil_claimline"] = createEntityRef("hil_claimlines", new Guid(_job.hil_claimline), service);
                Console.WriteLine("hil_claimheader");
            }
            if (_job.msdyn_closedby != null)
            {
                entity["hil_closedby"] = _job.msdyn_closedby;// createEntityRef("hil_branch", new Guid(_job.hil_branch), service);
                Console.WriteLine("msdyn_closedby");
            }
            //if (_job.hil_consumercategory != null)
            //{
            //    entity["hil_consumercategory"] = createEntityRef("hil_consumercategory", new Guid(_job.hil_consumercategory), service);
            //    Console.WriteLine("hil_consumercategory");
            //}
            //if (_job.hil_consumertype != null)
            //{
            //    entity["hil_consumertype"] = createEntityRef("hil_consumertype", new Guid(_job.hil_consumertype), service);
            //    Console.WriteLine("hil_consumertype");
            //}
            //            if (_job.msdyn_workorderid != null)
            //            {
            //                entity["hil_transactioncurrencyid"] =
            //}
            if (_job.hil_customerref != null)
            {
                Console.WriteLine("hil_customerref");
                entity["hil_customerref"] = createEntityRef("contact", new Guid(_job.hil_customerref), service);
            }
            //            if (_job.msdyn_workorderid != null)
            //            {
            //                entity["hil_customerref_contact"] =
            //}
            if (_job.hil_district != null)
            {
                entity["hil_district"] = createEntityRef("hil_district", new Guid(_job.hil_district), service);
                Console.WriteLine("hil_district");
            }
            if (_job.hil_emailcustomer != null)
            {
                entity["hil_emailcustomer"] = createEntityRef("contact", new Guid(_job.hil_emailcustomer), service);
                Console.WriteLine("hil_district");
            }
            if (_job.hil_firstresponsein != null)
            {
                entity["hil_firstresponsein"] = createEntityRef("slakpiinstance", new Guid(_job.hil_firstresponsein), service);
                Console.WriteLine("hil_firstresponsein");
            }
            //if (_job.hil_fiscalmonth != null)
            //{
            //    entity["hil_fiscalmonth"] = createEntityRef("hil_claimperiod", new Guid(_job.hil_fiscalmonth), service);
            //    Console.WriteLine("hil_claimperiod");
            //}
            //if (_job.msdyn_functionallocation != null)
            //{
            //    entity["hil_functionallocation"] = createEntityRef("msdyn_functionallocation", new Guid(_job.msdyn_functionallocation), service);
            //    Console.WriteLine("msdyn_functionallocation");
            //}
            //if (_job.msdyn_iotalert != null)
            //{
            //    entity["hil_iotalert"] = createEntityRef("msdyn_iotalert", new Guid(_job.msdyn_iotalert), service);
            //      Console.WriteLine("msdyn_iotalert");
            //}
            if (_job.hil_jobclosureby != null)
            {
                entity["hil_jobclosureby"] = _job.hil_jobclosureby;// createEntityRef("hil_branch", new Guid(_job.hil_branch), service);
                Console.WriteLine("hil_jobclosureby");
            }
            //if (_job.hil_jobextension != null)
            //{
            //    entity["hil_jobextension"] = createEntityRef("hil_jobsextension", new Guid(_job.hil_jobextension), service);
            //    Console.WriteLine("hil_jobsextension");
            //}
            if (_job.slainvokedid != null)
            {
                entity["hil_slainvokedid"] = createEntityRef("sla", new Guid(_job.slainvokedid), service);
                Console.WriteLine("hil_slainvokedid");
            }
            if (_job.modifiedbyname != null)
            {
                entity["hil_modifiedby"] = _job.modifiedbyname;// createEntityRef("hil_branch", new Guid(_job.hil_branch), service);
                Console.WriteLine("modifiedbyname");
            }
            //            if (_job.msdyn_workorderid != null)
            //            {
            //                entity["hil_modifiedonbehalfby"] =
            //}
            if (_job.hil_natureofcomplaint != null)
            {
                entity["hil_natureofcomplaint"] = createEntityRef("hil_natureofcomplaint", new Guid(_job.hil_natureofcomplaint), service);
                Console.WriteLine("hil_natureofcomplaint");
            }
            //if (_job.hil_newjobid != null)
            //{
            //    entity["hil_newjobid"] = createEntityRef("msdyn_workorder", new Guid(_job.hil_newjobid), service);
            //    Console.WriteLine("hil_newjobid");
            //}
            if (_job.hil_observation != null)
            {
                entity["hil_observation"] = createEntityRef("hil_observation", new Guid(_job.hil_observation), service);
                Console.WriteLine("hil_observation");
            }
            if (_job.hil_onbehalfofdealer != null)
            {
                entity["hil_onbehalfofdealer"] = createEntityRef("contact", new Guid(_job.hil_onbehalfofdealer), service);
                Console.WriteLine("hil_onbehalfofdealer");
            }
            if (_job.msdyn_opportunityid != null)
            {
                entity["hil_opportunityid"] = createEntityRef("opportunity", new Guid(_job.msdyn_opportunityid), service);
                Console.WriteLine("hil_opportunityid");
            }
            if (_job.hil_owneraccount != null)
            {
                entity["hil_owneraccount"] = createEntityRef("account", new Guid(_job.hil_owneraccount), service);
                Console.WriteLine("hil_owneraccount");
            }
            //            if (_job.msdyn_workorderid != null)
            //            {
            //                entity["hil_ownerpowerapp"] = createEntityRef("account", new Guid(_job.hil_owneraccount), service);
            //                Console.WriteLine("hil_owneraccount");
            //            }
            //            if (_job.msdyn_workorderid != null)
            //            {
            //                entity["hil_owningbusinessunit"] =
            //}
            //            if (_job.msdyn_workorderid != null)
            //            {
            //                entity["hil_owningteam"] =
            //}
            //            if (_job.msdyn_workorderid != null)
            //            {
            //                entity["hil_owninguser"] =
            //}
            if (_job.msdyn_parentworkorder != null)
            {
                entity["hil_parentworkorder"] = createEntityRef("msdyn_workorder", new Guid(_job.msdyn_parentworkorder), service);
                Console.WriteLine("msdyn_parentworkorder");
            }
            if (_job.hil_pincode != null)
            {
                entity["hil_pincode"] = createEntityRef("hil_pincode", new Guid(_job.hil_pincode), service);
                Console.WriteLine("hil_pincode");
            }
            if (_job.msdyn_preferredresource != null)
            {
                entity["hil_preferredresource"] = createEntityRef("bookableresource", new Guid(_job.msdyn_preferredresource), service);
                Console.WriteLine("msdyn_preferredresource");
            }
            if (_job.msdyn_pricelist != null)
            {
                entity["hil_pricelist"] = createEntityRef("pricelevel", new Guid(_job.msdyn_pricelist), service);
                Console.WriteLine("msdyn_pricelist");
            }
            //if (_job.msdyn_primaryresolution != null)
            //{
            //    entity["hil_primaryresolution"] = createEntityRef("msdyn_resolution", new Guid(_job.msdyn_primaryresolution), service);
            //    Console.WriteLine("msdyn_primaryresolution");
            //}
            if (_job.msdyn_priority != null)
            {
                entity["hil_priority"] = createEntityRef("msdyn_priority", new Guid(_job.msdyn_priority), service);
                Console.WriteLine("msdyn_pricelist");
            }
            if (_job.hil_productcatsubcatmapping != null)
            {
                entity["hil_productcatsubcatmapping"] = createEntityRef("hil_stagingdivisonmaterialgroupmapping", new Guid(_job.hil_productcatsubcatmapping), service);
                Console.WriteLine("hil_productcatsubcatmapping");
            }
            if (_job.hil_productcategory != null)
            {
                entity["hil_productcategory"] = createEntityRef("product", new Guid(_job.hil_productcategory), service);
                Console.WriteLine("hil_productcategory");
            }
            if (_job.msdyn_workorderid != null)
            {
                entity["hil_productsubcategory"] = createEntityRef("product", new Guid(_job.hil_productsubcategory), service);
                Console.WriteLine("hil_productsubcategory");
            }
            if (_job.hil_associateddealer != null)
            {
                entity["hil_associateddealer"] = createEntityRef("account", new Guid(_job.hil_associateddealer), service);
                Console.WriteLine("hil_associateddealer");
            }
            //if (_job.hil_reassignreason != null)
            //{
            //    entity["hil_reassignreason"] = createEntityRef("hil_jobreassignreason", new Guid(_job.hil_reassignreason), service);
            //    Console.WriteLine("hil_reassignreason");
            //}
            if (_job.hil_regardingemail != null)
            {
                entity["hil_regardingemail"] = createEntityRef("email", new Guid(_job.hil_regardingemail), service);
                Console.WriteLine("hil_regardingemail");
            }
            if (_job.hil_regardingfallback != null)
            {
                entity["hil_regardingfallback"] = createEntityRef("hil_sbubranchmapping", new Guid(_job.hil_regardingfallback), service);
                Console.WriteLine("hil_regardingfallback");
            }
            if (_job.hil_region != null)
            {
                entity["hil_region"] = createEntityRef("hil_region", new Guid(_job.hil_region), service);
                Console.WriteLine("hil_region");
            }
            if (_job.hil_assignmentmatrix != null)
            {
                entity["hil_relatedassignmentmatrix"] = createEntityRef("hil_assignmentmatrix", new Guid(_job.hil_assignmentmatrix), service);
                Console.WriteLine("hil_assignmentmatrix");
            }
            if (_job.hil_characterstics != null)
            {
                entity["hil_characterstics"] = createEntityRef("characteristic", new Guid(_job.hil_characterstics), service);
                Console.WriteLine("characteristic");
            }
            if (_job.msdyn_reportedbycontact != null)
            {
                entity["hil_reportedbycontact"] = createEntityRef("contact", new Guid(_job.msdyn_reportedbycontact), service);
                Console.WriteLine("msdyn_reportedbycontact");
            }
            if (_job.hil_resolvebykpi != null)
            {
                entity["hil_resolvebykpi"] = createEntityRef("slakpiinstance", new Guid(_job.hil_resolvebykpi), service);
                Console.WriteLine("slakpiinstance");
            }
            if (_job.hil_salesoffice != null)
            {
                entity["hil_salesoffice"] = createEntityRef("hil_salesoffice", new Guid(_job.hil_salesoffice), service);
                Console.WriteLine("hil_salesoffice");
            }
            if (_job.msdyn_taxcode != null)
            {
                entity["hil_taxcode"] = createEntityRef("msdyn_taxcode", new Guid(_job.msdyn_taxcode), service);
                Console.WriteLine("msdyn_taxcode");
            }
            //if (_job.hil_schemecode != null)
            //{
            //    entity["hil_schemecode"] = createEntityRef("hil_schemeincentive", new Guid(_job.hil_schemecode), service);
            //    Console.WriteLine("hil_schemeincentive");
            //}
            if (_job.msdyn_serviceaccount != null)
            {
                entity["hil_serviceaccount"] = createEntityRef("account", new Guid(_job.msdyn_serviceaccount), service);
                Console.WriteLine("hil_serviceaccount");
            }
            if (_job.hil_serviceaddress != null)
            {
                entity["hil_serviceaddress"] = createEntityRef("hil_address", new Guid(_job.hil_serviceaddress), service);
                Console.WriteLine("hil_serviceaddress");
            }
            if (_job.msdyn_serviceterritory != null)
            {
                entity["hil_serviceterritory"] = createEntityRef("territory", new Guid(_job.msdyn_serviceterritory), service);
                Console.WriteLine("territory");
            }
            if (_job.slaid != null)
            {
                entity["hil_slaid"] = createEntityRef("sla", new Guid(_job.slaid), service);
                Console.WriteLine("hil_slaid");
            }
            if (_job.hil_state != null)
            {
                entity["hil_state"] = createEntityRef("hil_state", new Guid(_job.hil_state), service);
                Console.WriteLine("hil_state");
            }
            if (_job.hil_nshname != null)
            {
                entity["hil_nsh"] = _job.hil_nshname;// createEntityRef("hil_state", new Guid(_job.hil_state), service);
                Console.WriteLine("hil_nsh");
            }
            if (_job.msdyn_substatus != null)
            {
                entity["hil_substatus"] = createEntityRef("msdyn_workordersubstatus", new Guid(_job.msdyn_substatus), service);
                Console.WriteLine("hil_substatus");
            }
            if (_job.msdyn_supportcontact != null)
            {
                entity["hil_supportcontact"] = createEntityRef("bookableresource", new Guid(_job.msdyn_supportcontact), service);
                Console.WriteLine("msdyn_supportcontact");
            }
            //if (_job.hil_tatachievementslab != null)
            //{
            //    entity["hil_tatachievementslab"] = createEntityRef("hil_tatachievementslabmaster", new Guid(_job.hil_tatachievementslab), service);
            //    Console.WriteLine("hil_tatachievementslab");
            //}
            //if (_job.hil_tatcategory != null)
            //{
            //    entity["hil_tatcategory"] = createEntityRef("hil_jobtatcategory", new Guid(_job.hil_tatcategory), service);
            //    Console.WriteLine("hil_tatcategory");
            //}
            if (_job.msdyn_timegroup != null)
            {
                entity["hil_timegroup"] = createEntityRef("msdyn_timegroup", new Guid(_job.msdyn_timegroup), service);
                Console.WriteLine("msdyn_timegroup");
            }
            if (_job.msdyn_timegroupdetailselected != null)
            {
                entity["hil_timegroupdetailselected"] = createEntityRef("msdyn_timegroupdetail", new Guid(_job.msdyn_timegroupdetailselected), service);
                Console.WriteLine("msdyn_timegroupdetailselected");
            }
            if (_job.hil_town != null)
            {
                entity["hil_town"] = createEntityRef("hil_city", new Guid(_job.hil_town), service);
                Console.WriteLine("hil_town");
            }
            if (_job.hil_typeofassignee != null)
            {
                entity["hil_typeofassignee"] = createEntityRef("position", new Guid(_job.hil_typeofassignee), service);
                Console.WriteLine("hil_typeofassignee");
            }
            //if (_job.hil_unitwarranty != null)
            //{
            //    entity["hil_unitwarranty"] = createEntityRef("hil_unitwarranty", new Guid(_job.hil_unitwarranty), service);
            //    Console.WriteLine("hil_unitwarranty");
            //}
            if (_job.hil_workdonebyname != null)
            {
                entity["hil_workdoneby"] = _job.hil_workdonebyname;// createEntityRef("position", new Guid(_job.hil_typeofassignee), service);
                Console.WriteLine("hil_workdonebyname");
            }
            //if (_job.msdyn_workhourtemplate != null)
            //{
            //    entity["hil_workhourtemplate"] = createEntityRef("msdyn_workhourtemplate", new Guid(_job.msdyn_workhourtemplate), service);
            //    Console.WriteLine("msdyn_workhourtemplate");
            //}
            //if (_job.msdyn_workorderarrivaltimekpiid != null)
            //{
            //    entity["hil_workorderarrivaltimekpiid"] = createEntityRef("slakpiinstance", new Guid(_job.msdyn_workorderarrivaltimekpiid), service);
            //    Console.WriteLine("msdyn_workorderarrivaltimekpiid");
            //}
            //if (_job.msdyn_workorderresolutionkpiid != null)
            //{
            //    entity["hil_workorderresolutionkpiid"] = createEntityRef("slakpiinstance", new Guid(_job.msdyn_workorderresolutionkpiid), service);
            //    Console.WriteLine("msdyn_workorderresolutionkpiid");
            //}
            if (_job.msdyn_workordertype != null)
            {
                entity["hil_workordertype"] = createEntityRef("msdyn_workordertype", new Guid(_job.msdyn_workordertype), service);
                Console.WriteLine("msdyn_workordertype");
            }

            Console.WriteLine("DataCollections method  end");
            return entity;
        }
        private static Entity DataCollectionsforIncident(IOrganizationService service, dtoIncidents _incident)
        {
            Entity entity = new Entity("hil_jobincidentarchived");
            Console.WriteLine("DataCollectionsforIncident method  started");

            if (_incident.msdyn_agreementbookingincident != null)
            {
                entity["hil_agreementbookingincident"] = createEntityRef("msdyn_agreementbookingincident", new Guid(_incident.msdyn_agreementbookingincident), service);
                Console.WriteLine("msdyn_agreementbookingincident");
            }
            if (_incident.msdyn_incidenttype != null)
            {
                entity["hil_cause"] = createEntityRef("msdyn_incidenttype", new Guid(_incident.msdyn_incidenttype), service);
                Console.WriteLine("msdyn_incidenttype");
                //
            }
            if (_incident.createdbyname != null)
            {
                entity["hil_createdby"] = _incident.createdbyname;
                Console.WriteLine("hil_createdby");
                //
            }
            //if (_incident.hil_purewatertds != null)
            //{
            //    entity["hil_createdbydelegate"] =
            // }
            if (_incident.hil_purewatertds != null)
            {
                entity["hil_createdon"] = dateFormater1(_incident.createdon.ToString(), tracingService);
                Console.WriteLine("hil_createdon");
            }
            if (_incident.hil_customer != null)
            {
                entity["hil_customer"] = createEntityRef("contact", new Guid(_incident.hil_customer), service);
                Console.WriteLine("hil_customer");
                //  
            }
            if (_incident.msdyn_customerasset != null)
            {
                entity["hil_customerasset"] = createEntityRef("msdyn_customerasset", new Guid(_incident.msdyn_customerasset), service);
                Console.WriteLine("msdyn_customerasset");
            }
            if (_incident.msdyn_description != null)
            {
                entity["msdyn_description"] = _incident.msdyn_description;
                Console.WriteLine("msdyn_description");
            }
            if (_incident.msdyn_estimatedduration != null)
            {
                entity["msdyn_estimatedduration"] = _incident.msdyn_estimatedduration;
                Console.WriteLine("msdyn_estimatedduration");
            }
            //if (_incident.msdyn_functionallocation != null)
            //{
            //    entity["hil_functionallocation"] = createEntityRef("msdyn_customerasset", new Guid(_incident.msdyn_customerasset), service);
            //    Console.WriteLine("msdyn_customerasset");
            //}
            //if (_incident.hil_purewatertds != null)
            //{
            //    entity["hil_importsequencenumber"] =
            // }
            if (_incident.msdyn_incidentresolved != null)
            {
                entity["msdyn_incidentresolved"] = _incident.msdyn_incidentresolved;
                Console.WriteLine("msdyn_incidentresolved");
            }
            if (_incident.hil_inputwatertds != null)
            {
                entity["hil_inputwatertds"] = _incident.hil_inputwatertds;
            }
            if (_incident.msdyn_internalflags != null)
            {
                entity["msdyn_internalflags"] = _incident.msdyn_internalflags;
                Console.WriteLine("msdyn_internalflags");
            }
            if (_incident.msdyn_ismobile != null)
            {
                entity["msdyn_ismobile"] = _incident.msdyn_ismobile;
                Console.WriteLine("msdyn_ismobile");
            }
            if (_incident.msdyn_isprimary != null)
            {
                entity["msdyn_isprimary"] = _incident.msdyn_isprimary;
                Console.WriteLine("msdyn_isprimary");
            }
            if (_incident.hil_iswarrantyvoid != null)
            {
                entity["hil_iswarrantyvoid"] = _incident.hil_iswarrantyvoid;
                Console.WriteLine("hil_iswarrantyvoid");
            }
            if (_incident.msdyn_itemspopulated != null)
            {
                entity["hil_itemspopulated"] = _incident.msdyn_itemspopulated;
                Console.WriteLine("hil_itemspopulated");
            }
            if (_incident.msdyn_itemspopulated != null)
            {
                entity["msdyn_itemspopulated"] = _incident.msdyn_itemspopulated;
                Console.WriteLine("msdyn_itemspopulated");
            }
            if (_incident.msdyn_workorder != null)
            {
                entity["hil_jobs"] = createEntityRef("hil_jobsarchived", new Guid(_incident.msdyn_workorder), service);
                Console.WriteLine("msdyn_workorder");
                //  
            }
            if (_incident.msdyn_workorderincidentid != null)
            {
                Console.WriteLine("msdyn_workorderid" + (_incident.msdyn_workorderincidentid));
                entity["hil_jobincidentarchivedid"] = new Guid(_incident.msdyn_workorderincidentid);
            }
            //if (_incident.hil_purewatertds != null)
            //{
            //    entity["hil_jobs"] =
            // }
            if (_incident.hil_modelcode != null)
            {
                Console.WriteLine("hil_modelcode");
                entity["hil_modelcode"] = createEntityRef("product", new Guid(_incident.hil_modelcode), service);
            }
            if (_incident.hil_modeldescription != null)
            {
                entity["hil_modeldescription"] = _incident.hil_modeldescription;
                Console.WriteLine("hil_modeldescription");
            }
            if (_incident.hil_modelname != null)
            {
                entity["hil_modelname"] = _incident.hil_modelname;
                Console.WriteLine("hil_modelname");
            }
            if (_incident.modifiedbyname != null)
            {
                entity["hil_modifiedby"] = _incident.modifiedbyname;
                Console.WriteLine("modifiedbyname");
            }
            //if (_incident.hil_purewatertds != null)
            //{
            //    entity["hil_modifiedbydelegate"] =
            // }
            if (_incident.hil_purewatertds != null)
            {
                entity["hil_modifiedon"] = dateFormater1(_incident.modifiedon.ToString(), tracingService);
                Console.WriteLine("hil_modifiedon");
            }
            if (_incident.msdyn_name != null)
            {
                entity["hil_name"] = _incident.msdyn_name;
            }
            if (_incident.msdyn_name != null)
            {
                entity["msdyn_name"] = _incident.msdyn_name;
            }
            if (_incident.hil_natureofcomplaint != null)
            {
                entity["hil_natureofcomplaint"] = createEntityRef("hil_natureofcomplaint", new Guid(_incident.hil_natureofcomplaint), service);
                Console.WriteLine("hil_natureofcomplaint");
                //
            }
            if (_incident.hil_noofpeopleinhome != null)
            {
                entity["hil_noofpeopleinhome"] = _incident.hil_noofpeopleinhome;
            }
            if (_incident.hil_observation != null)
            {
                entity["hil_observation"] = createEntityRef("hil_observation", new Guid(_incident.hil_observation), service);
                Console.WriteLine("hil_observation");
                //
            }
            if (_incident.owneridname != null)
            {
                Console.WriteLine("hil_owner");
                entity["hil_owner"] = _incident.owneridname;//createEntityRef("systemuser", new Guid(_incident.ownerid), service);
            }
            //if (_incident.hil_purewatertds != null)
            //{
            //    entity["hil_owningbusinessunit"] =
            // }
            //if (_incident.hil_purewatertds != null)
            //{
            //    entity["hil_owningteam"] =
            // }
            //if (_incident.hil_purewatertds != null)
            //{
            //    entity["hil_owninguser"] =
            // }
            if (_incident.hil_ph != null)
            {
                entity["hil_ph"] = _incident.hil_ph;
            }
            //if (_incident.msdyn_primaryresolution != null)
            //{
            //    entity["hil_primaryresolution"] = createEntityRef("msdyn_resolution", new Guid(_incident.msdyn_primaryresolution), service);
            //    Console.WriteLine("hil_observation");
            //}
            if (_incident.hil_productcategory != null)
            {
                Console.WriteLine("hil_productcategory");
                entity["hil_productcategory"] = createEntityRef("product", new Guid(_incident.hil_productcategory), service);
            }
            if (_incident.hil_productreplacement != null)
            {
                entity["hil_productreplacement"] = new OptionSetValue(int.Parse(_incident.hil_productreplacement.ToString()));
            }
            if (_incident.hil_productsubcategory != null)
            {
                Console.WriteLine("hil_productsubcategory");
                entity["hil_productsubcategory"] = createEntityRef("product", new Guid(_incident.hil_productsubcategory), service);
            }
            if (_incident.hil_purewatertds != null)
            {
                entity["hil_purewatertds"] = _incident.hil_purewatertds;
                Console.WriteLine("hil_purewatertds");
            }
            if (_incident.hil_quantity != null)
            {
                entity["hil_quantity"] = int.Parse(_incident.hil_quantity.ToString());
                Console.WriteLine("hil_quantity");
                //
            }
            //if (_incident.hil_purewatertds != null)
            //{
            //    entity["hil_overriddencreatedon"] =
            // }
            if (_incident.hil_rejectwatertds != null)
            {
                entity["hil_rejectwatertds"] = _incident.hil_rejectwatertds;
                Console.WriteLine("hil_rejectwatertds");
            }
            if (_incident.msdyn_resourcerequirement != null)
            {
                entity["hil_resourcerequirement"] = createEntityRef("msdyn_resourcerequirement", new Guid(_incident.msdyn_resourcerequirement), service);
                Console.WriteLine("hil_resourcerequirement");
            }
            if (_incident.hil_serialnumber != null)
            {
                entity["hil_serialnumber"] = _incident.hil_serialnumber;
                Console.WriteLine("hil_serialnumber");
            }
            if (_incident.msdyn_taskspercentcompleted != null)
            {
                entity["msdyn_taskspercentcompleted"] = _incident.msdyn_taskspercentcompleted;
                Console.WriteLine("msdyn_taskspercentcompleted");
            }
            if (_incident.hil_tds != null)
            {
                entity["hil_tds"] = _incident.hil_tds;
                Console.WriteLine("hil_tds");
            }
            //if (_incident.hil_purewatertds != null)
            //{
            //    entity["hil_timezoneruleversionnumber"] =
            // }
            //if (_incident.hil_purewatertds != null)
            //{
            //    entity["hil_utcconversiontimezonecode"] =
            // }
            //if (_incident.hil_purewatertds != null)
            //{
            //    entity["hil_versionnumber"] =
            // }
            if (_incident.new_warrantyenddate != null)
            {
                entity["hil_warrantyenddate"] = dateFormater(_incident.new_warrantyenddate.ToString(), tracingService);
                Console.WriteLine("new_warrantyenddate");
                //
            }
            if (_incident.new_warrantyenddate != null)
            {
                entity["hil_warrantyenddate"] = dateFormater(_incident.new_warrantyenddate.ToString(), tracingService);
                Console.WriteLine("new_warrantyenddate");
                //
            }
            if (_incident.hil_warrantystatus != null)
            {
                entity["hil_warrantystatus"] = int.Parse(_incident.hil_warrantystatus.ToString()) == 1 ? true : false;
                Console.WriteLine("hil_warrantystatus");
            }
            if (_incident.hil_warrantyvoidreasoncode != null)
            {
                entity["hil_warrantyvoidreasoncode"] = _incident.hil_warrantyvoidreasoncode.ToString();
                Console.WriteLine("hil_warrantyvoidreasoncode");
            }
            if (_incident.hil_watersource != null)
            {
                entity["hil_watersource"] = new OptionSetValue(int.Parse(_incident.hil_watersource.ToString()));
            }
            if (_incident.hil_waterstoragetype != null)
            {
                entity["hil_waterstoragetype"] = new OptionSetValue(int.Parse(_incident.hil_waterstoragetype.ToString()));
            }
            Console.WriteLine("DataCollections method  end");
            return entity;
        }
        private static Entity DataCollectionsforService(IOrganizationService service, dtoWorkOrderService _service)
        {
            Entity entity = new Entity("hil_jobservicearchived");
            Console.WriteLine("DataCollectionsforIncident method  started");
            //if (_service.createdby != null)
            //{
            //    entity["hil_agreementbookingservice"] = createEntityRef("msdyn_agreementbookingservice", new Guid(_service.msdyn_AgreementBookingService), service);
            //   Console.WriteLine("msdyn_customerasset");
            //}
            if (_service.msdyn_booking != null)
            {
                entity["hil_booking"] = createEntityRef("bookableresourcebooking", new Guid(_service.msdyn_booking), service);
                Console.WriteLine("hil_agreementbookingproduct");
            }
            if (_service.hil_charge != null)
            {
                entity["hil_charge"] = decimal.Parse(_service.hil_charge.ToString());
                Console.WriteLine("hil_charge");
            }
            if (_service.createdbyname != null)
            {
                entity["hil_createdby"] = _service.createdbyname;
                Console.WriteLine("createdbyname");
            }
            if (_service.createdon != null)
            {
                entity["hil_createdon"] = dateFormater1(_service.createdon.ToString(), tracingService);
                Console.WriteLine("hil_createdon");
            }
            //if (_service.createdby != null)
            //{
            //    entity["hil_createdonbehalfby"] =
            //   }
            if (_service.msdyn_customerasset != null)
            {
                entity["hil_customerasset"] = createEntityRef("msdyn_customerasset", new Guid(_service.msdyn_customerasset), service);
                Console.WriteLine("msdyn_customerasset");
            }
            if (_service.msdyn_description != null)
            {
                entity["hil_description"] = _service.msdyn_description;
                Console.WriteLine("hil_description");
            }
            if (_service.hil_discountedamount != null)
            {
                entity["hil_discountedamount"] = decimal.Parse(_service.hil_discountedamount);
                Console.WriteLine("hil_discountedamount");
            }
            if (_service.hil_effectiveamount != null)
            {
                entity["hil_effectivecharge"] = decimal.Parse(_service.hil_effectiveamount);
                Console.WriteLine("hil_effectiveamount");
            }
            //if (_service.createdby != null)
            //{
            //    entity["hil_entitlement"] =
            //   }
            if (_service.exchangerate != null)
            {
                entity["hil_exchangerate"] = decimal.Parse(_service.exchangerate.ToString());
                Console.WriteLine("exchangerate");
            }
            if (_service.hil_finalamount != null)
            {
                entity["hil_finalamount"] = decimal.Parse(_service.hil_finalamount.ToString());
                Console.WriteLine("hil_finalamount");
            }
            //if (_service.createdby != null)
            //{
            //    entity["hil_importsequencenumber"] =
            //   }
            if (_service.msdyn_workorderincident != null)
            {
                entity["hil_jobincident"] = new EntityReference()
                {
                    Id = new Guid(_service.msdyn_workorderincident),
                    LogicalName = "hil_jobincidentarchived",
                    Name = _service.msdyn_workorderincidentname,
                };
                Console.WriteLine("hil_jobincident");
            }
            if (_service.msdyn_workorder != null)
            {
                entity["hil_jobs"] = new EntityReference()
                {
                    Id = new Guid(_service.msdyn_workorder),
                    LogicalName = "hil_jobsarchived",
                    Name = _service.msdyn_workordername,
                };
                Console.WriteLine("job");
            }
            if (_service.msdyn_workorderserviceid != null)
            {
                entity["hil_jobservicearchivedid"] = new Guid(_service.msdyn_workorderserviceid);
                Console.WriteLine("msdyn_workorderserviceid");
            }
            if (_service.hil_markused != null)
            {
                entity["hil_markused"] = _service.hil_markused != "False" ? true : false; ;
                Console.WriteLine("hil_markused");
            }
            if (_service.modifiedbyname != null)
            {
                entity["hil_modifiedby"] = _service.modifiedbyname;
                Console.WriteLine("modifiedbyname");
            }
            if (_service.modifiedon != null)
            {
                entity["hil_modifiedon"] = dateFormater1(_service.modifiedon.ToString(), tracingService);
                Console.WriteLine("hil_modifiedon");
            }
            //if (_service.createdby != null)
            //{
            //    entity["hil_modifiedonbehalfby"] =
            //   }
            if (_service.msdyn_name != null)
            {
                entity["hil_name"] = _service.msdyn_name;
                Console.WriteLine("hil_name");
            }
            //if (_service.createdby != null)
            //{
            //    entity["hil_overriddencreatedon"] =
            //   }
            if (_service.owneridname != null)
            {
                entity["hil_owner"] = _service.owneridname;
                Console.WriteLine("Owner");
            }
            //if (_service.createdby != null)
            //{
            //    entity["hil_owningbusinessunit"] =
            //   }
            //if (_service.createdby != null)
            //{
            //    entity["hil_owningteam"] =
            //   }
            //if (_service.createdby != null)
            //{
            //    entity["hil_owninguser"] =
            //   }
            //if (_service.createdby != null)
            //{
            //    entity["hil_pricelist"] =
            //   }
            if (_service.hil_regardingestimate != null)
            {
                entity["hil_regardingestimate"] = createEntityRef("hil_estimate", new Guid(_service.hil_regardingestimate), service);
                Console.WriteLine("hil_regardingestimate");
            }
            if (_service.createdby != null)
            {
                entity["hil_service"] = createEntityRef("product", new Guid(_service.msdyn_service), service);
                Console.WriteLine("msdyn_service");
            }
            ////if (_service.createdby != null)
            ////{
            ////    entity["hil_timezoneruleversionnumber"] =
            ////   }
            //if (_service.createdby != null)
            //{
            //    entity["hil_transactioncurrencyid"] =
            //   }
            //if (_service.createdby != null)
            //{
            //    entity["hil_unit"] =
            //   }
            //if (_service.createdby != null)
            //{
            //    entity["hil_utcconversiontimezonecode"] =
            //   }
            //if (_service.createdby != null)
            //{
            //    entity["hil_versionnumber"] =
            //   }
            if (_service.hil_warrantystatus != null)
            {
                entity["hil_warrantystatus"] = new OptionSetValue(int.Parse(_service.hil_warrantystatus.ToString()));
                Console.WriteLine("hil_warrantystatus");
            }
            //if (_service.createdby != null)
            //{
            //    entity["hil_workorder"] =
            //   }
            //if (_service.createdby != null)
            //{
            //    entity["hil_workorderincident"] =
            //   }
            if (_service.msdyn_additionalcost != null)
            {
                entity["msdyn_additionalcost"] = decimal.Parse(_service.msdyn_additionalcost.ToString());
                Console.WriteLine("msdyn_additionalcost");
            }
            //if (_service.createdby != null)
            //{
            //    entity["msdyn_calculatedunitamount"] = decimal.Parse(_service.msdyn_cal.ToString());
            //   Console.WriteLine("hil_partamount");
            //}
            if (_service.msdyn_commissioncosts != null)
            {
                entity["msdyn_commissioncosts"] = decimal.Parse(_service.msdyn_commissioncosts.ToString());
                Console.WriteLine("msdyn_commissioncosts");
            }

            //if (_service.createdby != null)
            //{
            //    entity["msdyn_disableentitlement"] =
            //   }
            if (_service.msdyn_discountamount != null)
            {
                entity["msdyn_discountamount"] = decimal.Parse(_service.msdyn_discountamount.ToString());
                Console.WriteLine("msdyn_discountamount");
            }
            if (_service.msdyn_discountpercent != null)
            {
                entity["msdyn_discountpercent"] = decimal.Parse(_service.msdyn_discountpercent.ToString());
                Console.WriteLine("msdyn_discountamount");
            }
            //if (_service.msdyn_duration != null)
            //{
            //    entity["msdyn_duration"] = decimal.Parse(_service.msdyn_duration.ToString());
            //   Console.WriteLine("msdyn_discomsdyn_durationuntamount");
            //}
            //if (_service.msdyn_durationtobill != null)
            //{
            //    entity["msdyn_durationtobill"] =
            //   }
            //if (_service.msdyn_estimatecalculatedunitamount != null)
            //{
            //    entity["msdyn_estimatecalculatedunitamount"] =
            //   }
            if (_service.msdyn_estimatediscountamount != null)
            {
                entity["msdyn_estimatediscountamount"] = decimal.Parse(_service.msdyn_estimatediscountamount.ToString());
                Console.WriteLine("msdyn_estimatediscountamount");
            }
            if (_service.msdyn_estimatediscountpercent != null)
            {
                entity["msdyn_estimatediscountpercent"] = decimal.Parse(_service.msdyn_estimatediscountpercent.ToString());
                Console.WriteLine("msdyn_estimatediscountpercent");
            }
            //if (_service.msdyn_estimateduration != null)
            //{
            //    entity["msdyn_estimateduration"] =
            //   }
            if (_service.msdyn_estimatesubtotal != null)
            {
                entity["msdyn_estimatesubtotal"] = decimal.Parse(_service.msdyn_estimatesubtotal.ToString());
                Console.WriteLine("msdyn_estimatesubtotal");
            }
            if (_service.msdyn_estimatetotalamount != null)
            {
                entity["msdyn_estimatetotalamount"] = decimal.Parse(_service.msdyn_estimatetotalamount.ToString());
                Console.WriteLine("msdyn_estimatetotalamount");
            }
            if (_service.msdyn_estimatetotalcost != null)
            {
                entity["msdyn_estimatetotalcost"] = decimal.Parse(_service.msdyn_estimatetotalcost.ToString());
                Console.WriteLine("msdyn_estimatetotalcost");
            }
            if (_service.msdyn_estimateunitamount != null)
            {
                entity["msdyn_estimateunitamount"] = decimal.Parse(_service.msdyn_estimateunitamount.ToString());
                Console.WriteLine("msdyn_estimateunitamount");
            }
            if (_service.msdyn_estimateunitcost != null)
            {
                entity["msdyn_estimateunitcost"] = decimal.Parse(_service.msdyn_estimateunitcost.ToString());
                Console.WriteLine("msdyn_estimateunitcost");
            }
            if (_service.msdyn_internalflags != null)
            {
                entity["msdyn_internalflags"] = _service.msdyn_internalflags;
                Console.WriteLine("_service.msdyn_internalflags ");
            }
            if (_service.msdyn_lineorder != null)
            {
                entity["msdyn_lineorder"] = int.Parse(_service.msdyn_lineorder);
                Console.WriteLine("_service.msdyn_lineorder ");
            }
            if (_service.msdyn_linestatus != null)
            {
                entity["msdyn_linestatus"] = new OptionSetValue(int.Parse(_service.msdyn_linestatus));
                Console.WriteLine("_service.msdyn_linestatus ");
            }
            //if (_service.msdyn_minimumchargeamount != null)
            //{
            //    entity["msdyn_minimumchargeamount"] =
            //   }
            //if (_service.msdyn_minimumchargeduration != null)
            //{
            //    entity["msdyn_minimumchargeduration"] =
            //   }
            //if (_service.createdby != null)
            //{
            //    entity["msdyn_name"] =
            //   }
            if (_service.msdyn_subtotal != null)
            {
                entity["msdyn_subtotal"] = decimal.Parse(_service.msdyn_subtotal);
                Console.WriteLine("_service.msdyn_subtotal ");
            }
            if (_service.msdyn_taxable != null)
            {
                entity["msdyn_taxable"] = _service.msdyn_taxable != "False" ? true : false;
                Console.WriteLine("hil_replacesamepart");
            }
            if (_service.msdyn_totalamount != null)
            {
                entity["msdyn_totalamount"] = decimal.Parse(_service.msdyn_totalamount);
                Console.WriteLine("_service.msdyn_totalamount ");
            }
            if (_service.msdyn_unitamount != null)
            {
                entity["msdyn_unitamount"] = decimal.Parse(_service.msdyn_unitamount);
                Console.WriteLine("_service.msdyn_unitamount ");
            }
            if (_service.msdyn_unitcost != null)
            {
                entity["msdyn_unitcost"] = decimal.Parse(_service.msdyn_unitcost);
                Console.WriteLine("_service.msdyn_unitcost ");
            }
            Console.WriteLine("DataCollections method  end");
            return entity;
        }
        private static Entity DataCollectionsforProduct(IOrganizationService service, dtoWorkorderProduct _product)
        {
            Entity entity = new Entity("hil_jobproductarchived");
            Console.WriteLine("DataCollectionsforIncident method  started");
            if (_product.msdyn_agreementbookingproduct != null)
            {
                entity["hil_agreementbookingproduct"] = createEntityRef("msdyn_agreementbookingproduct", new Guid(_product.msdyn_agreementbookingproduct), service);
                Console.WriteLine("hil_agreementbookingproduct");
            }
            if (_product.msdyn_allocated != null)
            {
                entity["hil_allocated"] = _product.msdyn_allocated != "False" ? true : false;
                Console.WriteLine("hil_allocated");
            }
            if (_product.hil_availabilitystatus != null)
            {
                entity["hil_availabilitystatus"] = new OptionSetValue(int.Parse(_product.hil_availabilitystatus.ToString())); ;
                Console.WriteLine("hil_availabilitystatus");
            }
            if (_product.msdyn_booking != null)
            {
                entity["hil_booking"] = createEntityRef("bookableresourcebooking", new Guid(_product.msdyn_booking), service);
                Console.WriteLine("hil_agreementbookingproduct");
            }
            if (_product.hil_chargeableornot != null)
            {
                entity["hil_chargeableornot"] = new OptionSetValue(int.Parse(_product.hil_chargeableornot.ToString()));
                Console.WriteLine("hil_chargeableornot");
            }
            if (_product.hil_checkavailability != null)
            {
                entity["hil_checkavailability"] = _product.hil_checkavailability != "False" ? true : false;
                Console.WriteLine("hil_checkavailability");
            }
            if (_product.createdbyname != null)
            {
                entity["hil_createdby"] = _product.createdbyname; //createEntityRef("systemuser", new Guid(_product.createdby), service);
                Console.WriteLine("hil_createdby");
            }
            //if (_product.createdbyyominame != null)
            //{
            //    entity["hil_createdbydelegate"] = "";
            //}
            if (_product.createdon != null)
            {
                entity["hil_createdon"] = dateFormater1(_product.createdon.ToString(), tracingService);
                Console.WriteLine("createdon");
            }
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["hil_currency"] =
            //}
            if (_product.msdyn_customerasset != null)
            {
                entity["hil_customerasset"] = createEntityRef("msdyn_customerasset", new Guid(_product.msdyn_customerasset), service);
                Console.WriteLine("hil_customerasset");
            }
            if (_product.hil_defectiveserialnumber != null)
            {
                entity["hil_defectiveserialnumber"] = _product.hil_defectiveserialnumber;
                Console.WriteLine("hil_defectiveserialnumber");
            }
            if (_product.hil_delinkpo != null)
            {
                entity["hil_delinkpo"] = _product.hil_delinkpo != "False" ? true : false;
                Console.WriteLine("hil_delinkpo");
            }
            if (_product.msdyn_description != null)
            {
                entity["hil_description"] = _product.msdyn_description;
                Console.WriteLine("hil_description");
            }
            if (_product.hil_discountedamount != null)
            {
                entity["hil_discountedamount"] = decimal.Parse(_product.hil_discountedamount.ToString());
                Console.WriteLine("hil_discountedamount");
            }
            if (_product.hil_effectiveamount != null)
            {
                entity["hil_effectiveamount"] = decimal.Parse(_product.hil_effectiveamount.ToString());
                Console.WriteLine("hil_effectiveamount");
            }
            //if (_product.msdyn_Entitlement != null)
            //{
            //    entity["hil_entitlement"] = createEntityRef("msdyn_customerasset", new Guid(_product.msdyn_Entitlement), service);
            //     Console.WriteLine("hil_effectiveamount");
            //}
            if (_product.exchangerate != null)
            {
                entity["hil_exchangerate"] = decimal.Parse(_product.exchangerate.ToString());
                Console.WriteLine("exchangerate");
            }
            if (_product.hil_finalamount != null)
            {
                entity["hil_finalamount"] = decimal.Parse(_product.hil_finalamount.ToString());
                Console.WriteLine("hil_finalamount");
            }
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["hil_importsequencenumber"] =
            //}
            if (_product.hil_isserialized != null)
            {
                entity["hil_isserialized"] = new OptionSetValue(int.Parse(_product.hil_isserialized.ToString())); ;
                Console.WriteLine("hil_isserialized");
            }
            if (_product.hil_iswarrantyvoid != null)
            {
                entity["hil_iswarrantyvoid"] = _product.hil_iswarrantyvoid != "False" ? true : false; ;
                Console.WriteLine("hil_iswarrantyvoid");
            }
            if (_product.msdyn_workorderincident != null)
            {
                entity["hil_jobincident"] = new EntityReference()
                {
                    Id = new Guid(_product.msdyn_workorderincident),
                    LogicalName = "hil_jobincidentarchived",
                    Name = _product.msdyn_workorderincidentname,
                };
                Console.WriteLine("hil_jobincident");
            }
            if (_product.msdyn_workorderproductid != null)
            {
                entity["hil_jobproductarchivedid"] = new Guid(_product.msdyn_workorderproductid);
                Console.WriteLine("hil_jobproductarchivedid");
            }
            if (_product.msdyn_workorder != null)
            {
                entity["hil_jobs"] = new EntityReference()
                {
                    Id = new Guid(_product.msdyn_workorder),
                    LogicalName = "hil_jobsarchived",
                    Name = _product.msdyn_workordername,
                };
                Console.WriteLine("job");
            }
            if (_product.hil_lastpostedjournal != null)
            {
                entity["hil_lastpostedjournal"] = createEntityRef("hil_inventoryjournal", new Guid(_product.hil_lastpostedjournal), service);
                Console.WriteLine("hil_lastpostedjournal");
            }
            if (_product.hil_linestatus != null)
            {
                entity["hil_linestatus"] = new OptionSetValue(int.Parse(_product.hil_linestatus.ToString()));
                Console.WriteLine("hil_linestatus");
            }
            if (_product.hil_markused != null)
            {
                entity["hil_markused"] = _product.hil_markused != "False" ? true : false; ;
                Console.WriteLine("hil_markused");
            }
            if (_product.hil_maxquantity != null)
            {
                entity["hil_maxquantity"] = decimal.Parse(_product.hil_maxquantity.ToString());
                Console.WriteLine("hil_maxquantity");
            }
            if (_product.modifiedbyname != null)
            {
                entity["hil_modifiedby"] = _product.modifiedbyname;
                Console.WriteLine("hil_maxquantity");
            }
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["hil_modifiedbydelegate"] =
            //}
            if (_product.modifiedon != null)
            {
                entity["hil_modifiedon"] = dateFormater1(_product.modifiedon.ToString(), tracingService);
                Console.WriteLine("hil_modifiedon");
            }
            if (_product.msdyn_product != null)
            {
                entity["hil_msdyn_product"] = createEntityRef("product", new Guid(_product.msdyn_product), service);
                Console.WriteLine("hil_msdyn_product");
            }
            if (_product.msdyn_name != null)
            {
                entity["hil_name"] = _product.msdyn_name;
                Console.WriteLine("hil_name");
            }
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["hil_overriddencreatedon"] =
            //}
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["hil_owningbusinessunit"] =
            //}
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["hil_owningteam"] =
            //}
            if (_product.owneridname != null)
            {
                entity["hil_owninguser"] = _product.owneridname;
                Console.WriteLine("hil_owninguser");
            }
            if (_product.hil_part != null)
            {
                entity["hil_part"] = _product.hil_part;
                Console.WriteLine("hil_part");
            }
            if (_product.hil_partamount != null)
            {
                entity["hil_partamount"] = decimal.Parse(_product.hil_partamount.ToString());
                Console.WriteLine("hil_partamount");
            }
            if (_product.hil_linestatus != null)
            {
                entity["hil_partstatus"] = new OptionSetValue(int.Parse(_product.hil_linestatus.ToString()));
                Console.WriteLine("hil_partstatus");
            }
            if (_product.hil_pendingquantity != null)
            {
                entity["hil_pendingquantity"] = decimal.Parse(_product.hil_pendingquantity.ToString());
                Console.WriteLine("hil_pendingquantity");
            }
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["hil_pricelist"] =
            //}
            if (_product.hil_priority != null)
            {
                entity["hil_priority"] = _product.hil_priority;
                Console.WriteLine("hil_priority");
            }
            if (_product.hil_purchaseorder != null)
            {
                entity["hil_purchaseorder"] = createEntityRef("hil_productrequest", new Guid(_product.hil_purchaseorder), service);
                Console.WriteLine("hil_purchaseorder");
            }
            if (_product.msdyn_purchaseorderreceiptproduct != null)
            {
                entity["hil_purchaseorderreceiptproduct"] = createEntityRef("msdyn_purchaseorderreceiptproduct", new Guid(_product.msdyn_purchaseorderreceiptproduct), service);
                Console.WriteLine("hil_purchaseorder");
            }
            if (_product.hil_partamount != null)
            {
                entity["hil_quantity"] = decimal.Parse(_product.msdyn_quantity.ToString());
                Console.WriteLine("hil_quantity");
            }
            if (_product.hil_regardingestimate != null)
            {
                entity["hil_regardingestimate"] = createEntityRef("hil_estimate", new Guid(_product.hil_regardingestimate), service);
                Console.WriteLine("hil_purchaseorder");
            }
            if (_product.hil_replacedpart != null)
            {
                entity["hil_replacedpart"] = createEntityRef("product", new Guid(_product.hil_replacedpart), service);
                Console.WriteLine("hil_replacedpart");
            }
            if (_product.hil_replacedpartdescription != null)
            {
                entity["hil_replacedpartdescription"] = _product.hil_replacedpartdescription;
                Console.WriteLine("hil_replacedpartdescription");
            }
            if (_product.hil_replacedserialnumber != null)
            {
                entity["hil_replacedserialnumber"] = _product.hil_replacedserialnumber;
                Console.WriteLine("hil_replacedserialnumber");
            }
            if (_product.hil_replacesamepart != null)
            {
                entity["hil_replacesamepart"] = _product.hil_replacesamepart != "False" ? true : false;
                Console.WriteLine("hil_replacesamepart");
            }
            if (_product.hil_returned != null)
            {
                entity["hil_returned"] = new OptionSetValue(int.Parse(_product.hil_returned.ToString()));
                Console.WriteLine("hil_returned");
            }
            //if (_product.hil_partamount != null)
            //{
            //    entity["hil_sparepart"] = createEntityRef("product", new Guid(_product.), service);
            //     Console.WriteLine("hil_replacedpart");
            //}
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["hil_timezoneruleversionnumber"] =
            //}
            if (_product.hil_totalamount != null)
            {
                entity["hil_totalamount"] = decimal.Parse(_product.hil_totalamount.ToString());
                Console.WriteLine("hil_totalamount");
            }
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["hil_unit"] =
            //}
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["hil_utcconversiontimezonecode"] =
            //}
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["hil_versionnumber"] =
            //}
            if (_product.msdyn_warehouse != null)
            {
                entity["hil_warehouse"] = createEntityRef("msdyn_warehouse", new Guid(_product.msdyn_warehouse), service);
                Console.WriteLine("hil_warehouse");
            }
            if (_product.hil_warrantystatus != null)
            {
                entity["hil_warrantystatus"] = new OptionSetValue(int.Parse(_product.hil_warrantystatus.ToString()));
                Console.WriteLine("hil_warrantystatus");
            }
            if (_product.hil_warrantyvoidreason != null)
            {
                entity["hil_warrantyvoidreason"] = createEntityRef("hil_warrantyvoidreason", new Guid(_product.hil_warrantyvoidreason), service);
                Console.WriteLine("hil_warehouse");
            }
            if (_product.msdyn_additionalcost != null)
            {
                entity["msdyn_additionalcost"] = decimal.Parse(_product.msdyn_additionalcost.ToString());
                Console.WriteLine("msdyn_additionalcost");
            }
            if (_product.hil_partamount != null)
            {
                entity["msdyn_additionalcost_base"] = decimal.Parse(_product.hil_partamount.ToString());
                Console.WriteLine("hil_partamount");
            }
            if (_product.hil_partamount != null)
            {
                entity["msdyn_allocated"] = _product.msdyn_allocated != "False" ? true : false;
                Console.WriteLine("hil_allocated");
            }
            if (_product.msdyn_commissioncosts != null)
            {
                entity["msdyn_commissioncosts"] = decimal.Parse(_product.msdyn_commissioncosts.ToString());
                Console.WriteLine("msdyn_commissioncosts");
            }
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["msdyn_commissioncosts_base"] =
            //}
            //                if (_product.msdyn_disableentitlement != null)
            //    {
            //        entity["msdyn_disableentitlement"] =
            //}
            if (_product.msdyn_discountamount != null)
            {
                entity["msdyn_discountamount"] = decimal.Parse(_product.msdyn_discountamount.ToString());
                Console.WriteLine("msdyn_discountamount");
            }
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["msdyn_discountamount_base"] =
            //}
            if (_product.msdyn_discountpercent != null)
            {
                entity["msdyn_discountpercent"] = decimal.Parse(_product.msdyn_discountpercent.ToString());
                Console.WriteLine("msdyn_discountamount");
            }
            if (_product.msdyn_estimatediscountamount != null)
            {
                entity["msdyn_estimatediscountamount"] = decimal.Parse(_product.msdyn_estimatediscountamount.ToString());
                Console.WriteLine("msdyn_estimatediscountamount");
            }
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["msdyn_estimatediscountamount_base"] =
            //}
            if (_product.msdyn_estimatediscountpercent != null)
            {
                entity["msdyn_estimatediscountpercent"] = decimal.Parse(_product.msdyn_estimatediscountpercent.ToString());
                Console.WriteLine("msdyn_estimatediscountpercent");
            }
            if (_product.msdyn_estimatequantity != null)
            {
                entity["msdyn_estimatequantity"] = decimal.Parse(_product.msdyn_estimatequantity.ToString());
                Console.WriteLine("msdyn_estimatequantity");
            }
            if (_product.msdyn_estimatesubtotal != null)
            {
                entity["msdyn_estimatesubtotal"] = decimal.Parse(_product.msdyn_estimatesubtotal.ToString());
                Console.WriteLine("msdyn_estimatesubtotal");
            }
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["msdyn_estimatesubtotal_base"] =
            //}
            if (_product.msdyn_estimatetotalamount != null)
            {
                entity["msdyn_estimatetotalamount"] = decimal.Parse(_product.msdyn_estimatetotalamount.ToString());
                Console.WriteLine("msdyn_estimatetotalamount");
            }
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["msdyn_estimatetotalamount_base"] =
            //}
            if (_product.msdyn_estimatetotalcost != null)
            {
                entity["msdyn_estimatetotalcost"] = decimal.Parse(_product.msdyn_estimatetotalcost.ToString());
                Console.WriteLine("msdyn_estimatetotalcost");
            }
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["msdyn_estimatetotalcost_base"] =
            //}
            if (_product.msdyn_estimateunitamount != null)
            {
                entity["msdyn_estimateunitamount"] = decimal.Parse(_product.msdyn_estimateunitamount.ToString());
                Console.WriteLine("msdyn_estimateunitamount");
            }
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["msdyn_estimateunitamount_base"] =
            //}
            if (_product.msdyn_estimateunitcost != null)
            {
                entity["msdyn_estimateunitcost"] = decimal.Parse(_product.msdyn_estimateunitcost.ToString());
                Console.WriteLine("msdyn_estimateunitcost");
            }
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["msdyn_estimateunitcost_base"] =
            //}
            if (_product.msdyn_internaldescription != null)
            {
                entity["msdyn_internaldescription"] = _product.msdyn_internaldescription;
                Console.WriteLine("msdyn_internaldescription");
            }
            if (_product.msdyn_internalflags != null)
            {
                entity["msdyn_internalflags"] = _product.msdyn_internalflags;
                Console.WriteLine("msdyn_internalflags");
            }
            if (_product.msdyn_lineorder != null)
            {
                entity["msdyn_lineorder"] = int.Parse(_product.msdyn_lineorder);
                Console.WriteLine("msdyn_lineorder");
            }
            if (_product.msdyn_linestatus != null)
            {
                entity["msdyn_linestatus"] = new OptionSetValue(int.Parse(_product.msdyn_linestatus.ToString()));
                Console.WriteLine("msdyn_linestatus");
            }

            if (_product.msdyn_qtytobill != null)
            {
                entity["msdyn_qtytobill"] = decimal.Parse(_product.msdyn_qtytobill.ToString());
                Console.WriteLine("msdyn_qtytobill");
            }
            if (_product.msdyn_quantity != null)
            {
                entity["msdyn_quantity"] = decimal.Parse(_product.msdyn_quantity.ToString());
                Console.WriteLine("msdyn_quantity");
            }
            if (_product.msdyn_subtotal != null)
            {
                entity["msdyn_subtotal"] = decimal.Parse(_product.msdyn_subtotal.ToString());
                Console.WriteLine("msdyn_subtotal");
            }
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["msdyn_subtotal_base"] =
            //}
            if (_product.msdyn_taxable != null)
            {
                entity["msdyn_taxable"] = _product.msdyn_taxable != "False" ? true : false;
                Console.WriteLine("hil_replacesamepart");
            }
            if (_product.msdyn_totalamount != null)
            {
                entity["msdyn_totalamount"] = decimal.Parse(_product.msdyn_totalamount.ToString());
                Console.WriteLine("msdyn_totalamount");
            }
            //if (_product.hil_partamount != null)
            //{
            //    entity["msdyn_totalamount_base"] = decimal.Parse(_product.msdyn_totalamount.ToString());
            //     Console.WriteLine("msdyn_totalamount");
            //}
            if (_product.msdyn_totalcost != null)
            {
                entity["msdyn_totalcost"] = decimal.Parse(_product.msdyn_totalcost.ToString());
                Console.WriteLine("msdyn_totalcost");
            }
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["msdyn_totalcost_base"] = msdyn_totalcost;
            //}
            if (_product.msdyn_unitamount != null)
            {
                entity["msdyn_unitamount"] = decimal.Parse(_product.msdyn_unitamount.ToString());
                Console.WriteLine("msdyn_unitamount");
            }
            //    if (_product.hil_partamount != null)
            //    {
            //        entity["msdyn_unitamount_base"] =
            //}
            if (_product.msdyn_unitcost != null)
            {
                entity["msdyn_unitcost"] = decimal.Parse(_product.msdyn_unitcost.ToString());
                Console.WriteLine("msdyn_unitcost");
            }
            //if (_product.hil_partamount != null)
            //{
            //    entity["msdyn_unitcost_base"] =
            //        }


            Console.WriteLine("DataCollections method  end");
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
        private static DateTime dateFormater1(string date)
        {
            //2021-02-11T16:36:52
            Console.WriteLine("date " + date);
            string[] dateTime = date.Split(' ');
            string[] dateOnly = dateTime[0].Split('-');
            string[] timeOnly = dateTime[1].Split(':');
            DateTime myDate = DateTime.ParseExact(dateOnly[2] + "-" +
                dateOnly[1] + "-" +
                dateOnly[0] + " " +
                timeOnly[0] + ":" +
                timeOnly[1] + ":" +
                timeOnly[2], "yyyy-MM-dd HH:mm:ss",
                System.Globalization.CultureInfo.InvariantCulture);
            Console.WriteLine("myDate " + myDate);
            return myDate;
        }
        private static DateTime dateFormater(string date)
        {
            
            //2021-02-11T16:36:52
            Console.WriteLine("date " + date);
            if (date.Contains("T"))
            {
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
                Console.WriteLine("myDate " + myDate);
                return myDate;
            }
            else
            {
                string[] dateTime = date.Split(' ');
                string[] dateOnly = dateTime[0].Split('-');
                string[] timeOnly = dateTime[1].Split(':');
                DateTime myDate = DateTime.ParseExact(dateOnly[0] + "-" +
                    dateOnly[1] + "-" +
                    dateOnly[2] + " " +
                    timeOnly[0] + ":" +
                    timeOnly[1] + ":" +
                    timeOnly[2], "dd-MM-yyyy HH:mm:ss",
                    System.Globalization.CultureInfo.InvariantCulture);
                Console.WriteLine("myDate " + myDate);
                return myDate;
            }
            
            
        }
        private static EntityReference createEntityRef(String entName, Guid guid, IOrganizationService _service)
        {
            try
            {
                //if(entName == "systemuser")
                //_service = ConnecttoCRM();
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
