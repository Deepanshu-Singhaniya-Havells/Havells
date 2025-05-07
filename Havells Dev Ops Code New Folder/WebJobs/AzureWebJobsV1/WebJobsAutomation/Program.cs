using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace WebJobsAutomation
{
    class Program
    {
        static void Main(string[] args)
        {
            var status = GetWebJobStatus("triggeredwebjobs/AMCJobDataSyncWithSAP");
            dynamic jsonStatus = JsonConvert.DeserializeObject(status);
            if (jsonStatus.latest_run != null)
            {
                var jobStatus = jsonStatus.latest_run.status;
            }
        }
        static string GetWebJobStatus(string call) {
            string ApiUrl = "https://havellscrmwsprod-as.scm.azurewebsites.net/api/";
            string result = string.Empty;
            string userPswd = "$HavellsCRMWSProd-AS" + ":" + "EqFKqiHAzlxJRcCcrqRj8Nytkss5bAsJpa5rmgXhofdiGTfvm4pHDurw4vaX";
            userPswd = Convert.ToBase64String(Encoding.Default.GetBytes(userPswd));
            string baseURL = string.Format("{0}", call);
            try
            {
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri(ApiUrl);
                    client.Timeout = TimeSpan.FromMinutes(30);
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", userPswd);
                    var response = new HttpResponseMessage();
                    response = client.GetAsync(baseURL).Result;
                    result = response.IsSuccessStatusCode ? (response.Content.ReadAsStringAsync().Result) : response.IsSuccessStatusCode.ToString();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return result;
        }

        static HttpResponseMessage POSTWebJobSync()
        {
            string ApiUrl = "https://havellscrmwsprod-as.scm.azurewebsites.net/api/";
            string call = "triggeredwebjobs/AMCJobDataSyncWithSAP/run";
            string result = string.Empty;
            string userPswd = "$HavellsCRMWSProd-AS" + ":" + "EqFKqiHAzlxJRcCcrqRj8Nytkss5bAsJpa5rmgXhofdiGTfvm4pHDurw4vaX";
            userPswd = Convert.ToBase64String(Encoding.Default.GetBytes(userPswd));
            var response = new HttpResponseMessage(HttpStatusCode.NotFound);
            try
            {
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri(ApiUrl);
                    client.Timeout = TimeSpan.FromMinutes(30);
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", userPswd);
                    response = client.PostAsync(call, new StringContent(string.Empty, System.Text.Encoding.UTF8, "application/json")).Result;
                }
                return response;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
