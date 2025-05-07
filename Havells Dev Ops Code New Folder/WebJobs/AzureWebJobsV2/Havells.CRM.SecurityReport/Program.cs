using Microsoft.Office.Interop.Excel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Configuration;

namespace Havells.CRM.SecurityReport
{
    class Program
    {
        #region Valiable
        static IOrganizationService service;
        static Guid loginUserGuid = Guid.Empty;
        static string _userId = string.Empty;
        static string _password = string.Empty;
        static string _soapOrganizationServiceUri = string.Empty;
        static List<string> systemUsersEmail = new List<string>();
        static Dictionary<String, List<string>> userrole = new Dictionary<string, List<string>>();
        static int count = 0;
        #endregion
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"].ToString();
            var CrmURL = ConfigurationManager.AppSettings["CRMUrl"].ToString();
            string finalString = string.Format(connStr, CrmURL);
            service = HavellsConnection.CreateConnection.createConnection(finalString);

            //UpdateUserTimeZone.updateTimeZone(service);
            //getPrivillage.getProvillage(service);

            //QueryExpression query = new QueryExpression("role");
            //query.ColumnSet = new ColumnSet(true);
            //EntityCollection userColl = service.RetrieveMultiple(query);



            //Console.WriteLine("Total User Retrived " + userColl.Entities.Count);


            //Entity entity = userColl[0];
            //foreach (string key in entity.Attributes.Keys)
            //{
            //    Console.WriteLine(key);
            //}
            //foreach (object key in entity.Attributes.Values)
            //{
            //    Console.WriteLine(key);
            //}

            //readExcel();
            //foreach (string email in systemUsersEmail)
            //{
            //    Guid userId = getUserGuid(email);
            //    List<string> roles = getUserSecurityRole(userId);
            //    userrole.Add(email, roles);
            //}
            Dictionary<Guid, string> userlist = getAllUserID();
            int totalcount = userlist.Count;
            Console.WriteLine("Total User Retrived " + totalcount);
            foreach (KeyValuePair<Guid, string> user in userlist)
            {
                List<string> roles = getUserSecurityRole(user.Key);
                userrole.Add(user.Value, roles);
            }
            createExcel(userrole);
        }
        protected static void createExcel(Dictionary<String, List<string>> userrole)
        {
            try
            {
                Application excel = new Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                Workbook worKbooK = excel.Workbooks.Add(Type.Missing);


                Worksheet worKsheeT = (Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = "New Sheet";

                worKsheeT.Cells[1, 1] = "User Email ID";

                int rowcount = 1;

                foreach (var user in userrole)
                {
                    string userEmail = user.Key;
                    rowcount += 1;
                    worKsheeT.Cells[rowcount, 1] = userEmail;
                    int col = 2;
                    int i = 0;
                    foreach (string role in user.Value)
                    {
                        string cc = "Role " + (col - 1);

                        worKsheeT.Cells[1, col] = cc;
                        worKsheeT.Cells[rowcount, col] = role;
                        col++;
                        i++;
                    }
                }

                worKbooK.SaveAs(@"C:\Users\35405\Downloads\SecurityReport.xlsx"); ;
                worKbooK.Close();
                excel.Quit();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        protected static Guid getUserGuid(string email)
        {
            QueryExpression query = new QueryExpression("systemuser");
            query.ColumnSet = new ColumnSet(false);
            query.Criteria.AddCondition("internalemailaddress", ConditionOperator.Equal, email);
            EntityCollection userColl = service.RetrieveMultiple(query);
            if (userColl.Entities.Count == 1)
            {
                return userColl[0].Id;
            }
            else
            {
                return Guid.Empty;
            }
        }
        protected static Dictionary<Guid, string> getAllUserID()
        {
            string name = "YOGESH1.KUMAR@HAVELLS.COM;vineet.mishra@havells.com;VINEESH.KR@HAVELLS.COM;Vinayak.K@havells.com;vikas.tiwari@havells.com;Vijay.Trivedi@havells.com;venkataraju.kt@havells.com;varun.arora@havells.com;v.prabakaran@havells.com;udaybhan.chauhan@havells.com;Uday.goud@havells.com;tushar.pramanick@havells.com;Tulsidas.rai@havells.com;SYED.AHSAN@HAVELLS.COM;SURESH.BABU@HAVELLS.COM;SURENDER.BANSAL@HAVELLS.COM;sumit.vyas@havells.com;SUBODH.PANDEY@HAVELLS.COM;7c276c5b4a3c4c54bf426aa6ed972ebfSubbarao.N@havells.com;Shivani.Modi@havells.com;SHIVA.KAMBHAMPAT@HAVELLS.COM;Sheesh.Ram@havells.com;sekhar.hariharan@havells.com;Sauradipthya.Paul@havells.com;saurabh.amar@havells.com;SATENDRA.SINGH@HAVELLS.COM;satapara.bharatkumar@havells.com;Sanku.Sarkar@havells.com;SANJEEV.MISHRA@HAVELLS.COM;SAMIR.KULKARNI@HAVELLS.COM;Sambarta.Mazumdar@havells.com;SAIF.ALI@HAVELLS.COM;RUPINDER.KANWAR@HAVELLS.COM;rudraprasad.jena@havells.com;Rohit.Vij@havells.com;RaviKumar.Mishra@havells.com;RAMESH1.SHARMA@HAVELLS.COM;RAKESH.RAJPUT@HAVELLS.COM;Rajnish.Kumar@havells.com;rajinder.cheema@havells.com;Rajib.Chakraborty@havells.com;RAJESH1.AGARWAL@HAVELLS.COM;RAJENDRA.DHANDHARE@HAVELLS.COM;RAHUL1.BHARDWAJ@HAVELLS.COM;rahul8.sharma@havells.com;rahul.kaushal@havells.com;R.SELVI@HAVELLS.COM;prashant.mathur@havells.com;pranay.tiwari@havells.com;Prakhar.Vasu@havells.com;pradeepkumar.misra@havells.com;Pradeep.Jain@havells.com;Prachi@havells.com;PAWAN2.KUMAR@HAVELLS.COM;parveenkumar.miglani@havells.com;PARTHASARTHI.AICH@HAVELLS.COM;pankaj.sethia@havells.com;panchal.neha@havells.com;Palanivelu.A@havells.com;NITIN1.TAYAL@HAVELLS.COM;nitin1.paliwal@havells.com;nitin.paliwal@havells.com;35b4a272ee3c4c919ea2ecf69edeb252nitin.bele@havells.com;nilangshu.shome@havells.com;Nikita.Gujar@havells.com;neha1.gupta@havells.com;Neerat.Mathur@havells.com;Neeraj.Gupta@havells.com;navneet.shrivastava@havells.com;NAVDEEPK.SINGH@HAVELLS.COM;navdeep.kapoor@havells.com;NARENDRAKUMAR.CHOUBEY@HAVELLS.COM;NARENDRA.YADAV@HAVELLS.COM;1f4f7139022644adb178c5b0e946a46fnarender.singh@havells.com;nagmaninath.tiwari@havells.com;mythri.as@havells.com;Muzaffar.Ahmed@havells.com;Muthu.C@havells.com;mukund.sahni@havells.com;MohammadLatief.Budoo@havells.com;ML.Virmani@Havells.Com;melvin.george@havells.com;manoj.varma@havells.com;MANOHAR.LAL@HAVELLS.COM;Manjunath.Telkar@havells.com;ManjitSingh.Puri@havells.com;Manish.Kaushik@havells.com;MAHENDRA.GUPTA@HAVELLS.COM;M.Kalaiarasan@havells.com;m.jeyachandran@havells.com;Kaushik.Basu@havells.com;Kanwar.Rajeev@havells.com;kanan.gandhi@havells.com;jogendra.poonia@havells.com;JJEYA.GANESH@HAVELLS.COM;JAIDEEP.TANTIA@HAVELLS.COM;husain.zakee@havells.com;himanshu.saini@havells.com;HEM.CHAND@HAVELLS.COM;HARENDRA.SHARMA@HAVELLS.COM;GOVIND.MIRWANI@HAVELLS.COM;gourinath.rachakonda@havells.com;GOPALK.KHATRI@HAVELLS.COM;Gitasish.Choudhury@havells.com;Geddam.Anjaneyulu@havells.com;G.Ramesh@havells.com;EKTA.AGARWAL@HAVELLS.COM;Durgesh.Roy@havells.com;dhrumit.sutariya@havells.com;dheeraj.Sharma@havells.com;Dhavalk.Shah@havells.com;DEEPTHI.BN@HAVELLS.COM;Deepika.Ahuja@havells.com;DeepakKumar.Banerjee@havells.com;Daxesh.Mistry@havells.com;DAULATRAM.BAIRWA@HAVELLS.COM;CHINMAYA.TRIPATHY@HAVELLS.COM;Chandrabhan.Gupta@havells.com;Chandan.Singh@havells.com;c.kalimuthu@havells.com;BK.Dinesh@havells.com;BHANUKUMAR.GARG@HAVELLS.COM;Barun.Mishra@havells.com;Banasree.Saha@havells.com;azad.kumar@havells.com;ASHWANI.CHUGH@HAVELLS.COM;ashok.malhotra@havells.com;ASHISHKR.GUPTA@HAVELLS.COM;ArunKumar.Jain@havells.com;b7283b90b3e34ceda62bf7a9135713a7Arun.Pandita@havells.com;archit.garg@havells.com;apoor.sabharwal@havells.com;anumukonda.anjani@havells.com;AnilRai.Gupta@havells.com;anil.thapa@havells.com;anand.pipalwa@havells.com;Anand.Garg@havells.com;Amit1.Mishra@havells.com;AMIT.MISHRA@HAVELLS.COM;AMIT.JANGRA@HAVELLS.COM;AKSHAT.SINGH@HAVELLS.COM;AkhilanJ.S@havells.com;AJOY.BHADURI@HAVELLS.COM;ajay.sunhotra@havells.com;ajay.grover@havells.com;ADITYAKUMAR.SONI@HAVELLS.COM;ADARSH.KAUSHIK@HAVELLS.COM;abhishek.singh@havells.com;abhishek1.chauhan@havells.com;abhinav.saxena@havells.com;abhijeet1.singh@havells.com;aatif.khan@havells.com;aadit.thakur@havells.com";
            string[] domainNAme = name.Split(';');
            Dictionary<Guid, string> allUserID = new Dictionary<Guid, string>();
            QueryExpression query = new QueryExpression("systemuser");
            query.ColumnSet = new ColumnSet("domainname");
            query.Criteria.AddCondition("isdisabled", ConditionOperator.Equal, false);
            query.Criteria.AddCondition("domainname", ConditionOperator.In, domainNAme);

            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;
            EntityCollection userColl = service.RetrieveMultiple(query);
            foreach (Entity user in userColl.Entities)
            {
                allUserID.Add(user.Id, user.GetAttributeValue<string>("domainname"));
            }
            do
            {
                query.PageInfo.PageNumber += 1;
                query.PageInfo.PagingCookie = userColl.PagingCookie;
                userColl = service.RetrieveMultiple(query);
                foreach (Entity user in userColl.Entities)
                    allUserID.Add(user.Id, user.GetAttributeValue<string>("domainname"));
            }
            while (userColl.MoreRecords);

            return allUserID;
        }
        protected static void readExcel()
        {
            string excelPath = @"C:\Users\35405\Downloads\UserEmail.xlsx";
            systemUsersEmail = new List<string>();
            Application excelApp = new Application();
            if (excelApp != null)
            {
                Workbook excelWorkbook = excelApp.Workbooks.Open(excelPath, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Worksheet excelWorksheet = (Worksheet)excelWorkbook.Sheets[1];
                Range excelRange = excelWorksheet.UsedRange;
                int rowCount = excelRange.Rows.Count;
                int iLeft = excelRange.Rows.Count;
                int colCount = excelRange.Columns.Count;
                int iDone = 0;
                int i = 2;
                bool hasRow = true;
                while (hasRow)  //rowCount
                {
                    try
                    {
                        if ((excelWorksheet.Cells[i, 1] as Range).Value != null)
                        {
                            var systemUserEmail = (excelWorksheet.Cells[i, 1] as Range).Value.ToString();
                            systemUsersEmail.Add(systemUserEmail);
                            Console.WriteLine(i);
                            i++;
                        }
                        else
                        {
                            hasRow = false;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: " + ex.Message);
                    }
                }
                excelWorkbook.Close();
                Console.WriteLine("Total row completed " + iDone);
                excelApp.Quit();
            }
        }
        protected static List<string> getUserSecurityRole(Guid userid)
        {

            List<string> roles = new List<string>();
            QueryExpression qe = new QueryExpression("systemuserroles");
            qe.ColumnSet.AddColumns("systemuserid");
            qe.Criteria.AddCondition("systemuserid", ConditionOperator.Equal, userid);

            LinkEntity link1 = qe.AddLink("systemuser", "systemuserid", "systemuserid", JoinOperator.Inner);
            link1.Columns.AddColumns("fullname", "internalemailaddress");
            LinkEntity link = qe.AddLink("role", "roleid", "roleid", JoinOperator.Inner);
            link.Columns.AddColumns("roleid", "name");
            EntityCollection results = service.RetrieveMultiple(qe);

            Console.WriteLine("Recod " + count + "/" + userid);
            count++;
            foreach (Entity Userrole in results.Entities)
            {
                if (Userrole.Attributes.Contains("role2.name"))
                {
                    roles.Add((Userrole.Attributes["role2.name"] as AliasedValue).Value.ToString());
                }
            }
            return roles;
        }
    }
}
