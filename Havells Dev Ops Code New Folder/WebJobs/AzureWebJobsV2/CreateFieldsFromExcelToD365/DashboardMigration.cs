using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CreateFieldsFromExcelToD365
{
    public class DashboardMigration
    {
        static Dictionary<string, string> viewIdsDict = new Dictionary<string, string>();
        public static void mainFunction(IOrganizationService service, IOrganizationService serviceSer)
        {
            CreateForm(serviceSer, service);
        }
        private static void CreateForm(IOrganizationService serviceSer, IOrganizationService service)
        {
            //Console.Clear();
            Console.WriteLine("********************** DASHBOARD CREATION IS STARTED **********************");
            var query = new QueryExpression("systemform");
            query.Criteria.AddCondition("type", ConditionOperator.Equal, 10);
            //query.Criteria.AddCondition("name", ConditionOperator.Equal, "ASH Dashboard");
            query.ColumnSet = new ColumnSet(true);

            EntityCollection results = service.RetrieveMultiple(query);
            Console.WriteLine("totla Form Count " + results.Entities.Count);
            int totalCount = results.Entities.Count;
            int done = 1;
            int error = 0;
            foreach (Entity entity in results.Entities)
            {
                try
                {
                    var query1 = new QueryExpression("systemform");
                    query1.Criteria.AddCondition("type", ConditionOperator.Equal, 10);
                    query1.Criteria.AddCondition("name", ConditionOperator.Equal, entity.GetAttributeValue<string>("name"));
                    query1.ColumnSet = new ColumnSet(true);

                    EntityCollection results1 = serviceSer.RetrieveMultiple(query1);
                    if (results1.Entities.Count == 0)
                    {
                        string formXML = entity.GetAttributeValue<string>("formxml");
                        string formJSON = entity.GetAttributeValue<string>("formjson");
                        viewIdsDict = new Dictionary<string, string>();

                        ChaneFormXml(formXML, service, serviceSer);
                        foreach (string key in viewIdsDict.Keys)
                        {
                            //Console.WriteLine(key + " : " + viewIdsDict[key]);
                            formXML = formXML.Replace(key.ToUpper(), viewIdsDict[key].ToUpper());
                            formJSON = formJSON.Replace(key.ToUpper(), viewIdsDict[key].ToUpper());
                        }


                        //formXML = formXML.Replace("8675a7e3-247b-4f83-8ffe-345723857150", "3D12F09F-1D84-43B8-8686-E3224C383D30");
                        //formJSON = formJSON.Replace("8675a7e3-247b-4f83-8ffe-345723857150", "3D12F09F-1D84-43B8-8686-E3224C383D30");

                        entity["formjson"] = formJSON;// entity.GetAttributeValue<string>("formjson");
                        entity["formxml"] = formXML; //entity.GetAttributeValue<string>("formxml");
                        try
                        {
                            serviceSer.Create(entity);
                            Console.WriteLine("Created " + entity.GetAttributeValue<string>("name"));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
                            Console.WriteLine("Error " + entity.GetAttributeValue<string>("name")+"  ||  "+ex.Message);
                            Console.WriteLine("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Skip " + entity.GetAttributeValue<string>("name"));
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
                    error++;
                    Console.WriteLine("-------------------------Failuer " + error + " in Form creation with name " + entity.GetAttributeValue<string>("name") +
                        ". Error is:-  " + ex.Message);
                    Console.WriteLine("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
                }
            }
            Console.WriteLine("********************** DASHBOARD CREATION IS ENDED **********************");
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
                xml = xml.Replace("<EntityViewId>", "|");
                xml = xml.Replace("</EntityViewId>", "^");
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
                //else if (savedQueries[0].GetAttributeValue<string>("name") == "Active Jobs")
                //    query1.Criteria.AddCondition("name", ConditionOperator.Equal, "Active Work Orders");
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

                foreach (Entity ent in savedQueries1)
                {
                    if (ent.GetAttributeValue<string>("returnedtypecode") == savedQueries[0].GetAttributeValue<string>("returnedtypecode"))
                    {
                        newViewId = ent.Id.ToString();
                    }
                    else
                    {
                        //  Console.WriteLine("dd");
                    }
                }
                if (savedQueries1.Count == 0)
                {
                    string nana = savedQueries[0].GetAttributeValue<string>("name");
                    Console.WriteLine(nana);
                    serviceSer.Create(savedQueries[0]);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


            return newViewId;
        }
    }
}
