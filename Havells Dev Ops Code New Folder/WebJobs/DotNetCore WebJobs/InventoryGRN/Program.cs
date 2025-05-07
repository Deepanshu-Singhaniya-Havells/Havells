using Microsoft.PowerPlatform.Dataverse.Client;
using System.Configuration;

namespace InventoryGRN
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var connStr = ConfigurationManager.AppSettings["connStr"];
            var CRMUrlProd = ConfigurationManager.AppSettings["CRMUrlProd"];
            string finalConnStr = string.Format(connStr ?? "", CRMUrlProd);
            ServiceClient service = new ServiceClient(finalConnStr);

            #region Create GRN
            var StartDate = ConfigurationManager.AppSettings["StartDate"];
            var EndDate = ConfigurationManager.AppSettings["EndDate"];
            var SONumber = ConfigurationManager.AppSettings["SONumber"];
            var Hr = ConfigurationManager.AppSettings["Hr"];
            ClsInventoryGRN obj = new ClsInventoryGRN(service);
            try
            {
                if (!string.IsNullOrWhiteSpace(SONumber))
                {
                    obj.CreateGRNLineBySONumberNew(SONumber);
                }
                else if (!string.IsNullOrWhiteSpace(Hr))
                {
                    obj.CreateGRNLineForSONumberGeneratedInLast_Hr(Hr);
                }
                //else
                //{
                //    obj.CreateGRNLine(StartDate, EndDate);
                //}
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            #endregion

            //#region Update PO Job Sub-Status
            //ClsUpdateJobSubStatus obj1 = new ClsUpdateJobSubStatus(service);
            //try
            //{
            //    obj1.UpdatePOJobSubStatus();
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //}
            //#endregion
        }
    }
}
