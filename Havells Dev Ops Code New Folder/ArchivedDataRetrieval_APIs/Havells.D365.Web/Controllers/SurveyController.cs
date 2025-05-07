using Havells.D365.Entities.Survey.Request;
using Havells.D365.Services.Abstract;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.ComponentModel.DataAnnotations;
using System.Net.Http;
using System.Net.Http.Headers;

namespace Havells.D365.Web.Controllers
{
    public class SurveyController : Controller
    {
        private readonly ISurveyRepository _surveyRepository;
        private readonly ILogger _logger;
        private readonly IConfiguration _configuration;
        public SurveyController(ISurveyRepository surveyRepository, ILoggerFactory logFactory, IConfiguration configuration)
        {
            this._surveyRepository = surveyRepository;
            this._logger = logFactory.CreateLogger<SurveyController>();
            this._configuration = configuration;
        }

        public ActionResult SurveyLandingPage()
        {
            return View();
        }
        public IActionResult Index([Required] string surveyCode)
        {
            _logger.LogInformation("Servey Code ="+ surveyCode);
            _logger.LogInformation("Base URl =" + _configuration["URLConnectionStrings:Survey"].ToString());
            if (ModelState.IsValid && !string.IsNullOrEmpty(surveyCode))
            {
                surveyCode = surveyCode.Replace(" ", "+");
                _logger.LogInformation("Servey Code After replace space =" + surveyCode);
                SurveyRequest request = new SurveyRequest();
                request.JobId = surveyCode;
                var surveyResult = _surveyRepository.GetSurveyDetail(request);

                _logger.LogInformation("Service Result =" + JsonConvert.SerializeObject(surveyResult));
                if (surveyResult.JobId==null)
                    return BadRequest("Invalid survey code");
                ViewBag.JobId = surveyCode;
                ViewBag.Status = surveyResult.SubmitStatus.ToString();
                ViewBag.BrandName = surveyResult.BrandName;
                _logger.LogInformation("Servey  Code Executed");
                return View();
            }
            else
                return BadRequest("Invalid URL");
        }


        public JsonResult SubmitFeedback([FromBody]SurveyUpdateRequest survey)
        {
            var result = _surveyRepository.SubmitSurvey(survey);
            return Json(result);
        }

        
       
       

    }
}
