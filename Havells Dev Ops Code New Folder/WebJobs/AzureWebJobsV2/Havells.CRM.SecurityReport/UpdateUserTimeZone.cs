using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Havells.CRM.SecurityReport
{
    public class UpdateUserTimeZone
    {
        public static void updateTimeZone(IOrganizationService service)
        {
            try
            {


                QueryExpression query = new QueryExpression("systemuser");
                query.ColumnSet = new ColumnSet("domainname");
                query.Criteria.AddCondition("isdisabled", ConditionOperator.Equal, false);
                query.PageInfo = new PagingInfo();
                query.PageInfo.Count = 5000;
                query.PageInfo.PageNumber = 1;
                query.PageInfo.ReturnTotalRecordCount = true;
                EntityCollection userColl = service.RetrieveMultiple(query);



                query = new QueryExpression("systemuser");
                query.ColumnSet = new ColumnSet(false);
                query.Criteria = new FilterExpression(LogicalOperator.Or);
                query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("560B3E66-C77C-EB11-A812-0022486EA516"));
                query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("C3565272-C77C-EB11-A812-0022486EA516"));//call center user
                query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("B7616B7E-C77C-EB11-A812-0022486EA516"));
                query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("34016996-C77C-EB11-A812-0022486EA516"));//call center user
                query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("30B7738A-C77C-EB11-A812-0022486EA516"));
                query.Criteria.AddCondition("positionid", ConditionOperator.Equal, new Guid("42024E5D-9080-EB11-A812-0022486E73EB"));//call center user
                EntityCollection entiColl = service.RetrieveMultiple(query);
                int count = 1;
                foreach (Entity user in entiColl.Entities)
                {
                    query = new QueryExpression("usersettings");
                    query.ColumnSet = new ColumnSet(false);
                    query.Criteria = new FilterExpression(LogicalOperator.And);
                    query.Criteria.AddCondition("systemuserid", ConditionOperator.Equal, user.Id);
                    entiColl = service.RetrieveMultiple(query);
                    if (entiColl.Entities.Count > 0)
                    {
                        try
                        {
                            Entity userSetting = new Entity("usersettings");
                            userSetting["timeformatcode"] = 0;
                            userSetting["timeformatstring"] = "HH:mm";
                            userSetting["timeseparator"] = ":";
                            userSetting["timezonebias"] = -330;
                            userSetting["timezonecode"] = 190;
                            userSetting["timezonedaylightbias"] = -60;
                            userSetting["timezonedaylightday"] = 0;
                            userSetting["timezonedaylightdayofweek"] = 0;
                            userSetting["timezonedaylighthour"] = 0;
                            userSetting["timezonedaylightminute"] = 0;
                            userSetting["timezonedaylightmonth"] = 0;
                            userSetting["timezonedaylightsecond"] = 0;
                            userSetting["timezonedaylightyear"] = 0;
                            userSetting["timezonestandardbias"] = 0;
                            userSetting["timezonestandardday"] = 0;
                            userSetting["timezonestandarddayofweek"] = 0;
                            userSetting["timezonestandardhour"] = 0;
                            userSetting["timezonestandardminute"] = 0;
                            userSetting["timezonestandardmonth"] = 0;
                            userSetting["timezonestandardsecond"] = 0;
                            userSetting["timezonestandardyear"] = 0;
                            userSetting.Id = entiColl[0].Id;
                            service.Update(userSetting);
                            Console.WriteLine("Time Zone Updated " + count);
                            count++;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error " + ex.Message);
                            continue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error " + ex.Message);

            }
        }
    }
}
