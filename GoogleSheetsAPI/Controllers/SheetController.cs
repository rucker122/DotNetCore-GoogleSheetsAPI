using System;
using System.Linq;
using System.Collections.Generic;
using Microsoft.AspNetCore.Mvc;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Data = Google.Apis.Sheets.v4.Data;

using GoogleSheetsAPI;
using Newtonsoft.Json;

namespace GoogleSheetsAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class SheetController : Controller
    {
        private APIHelper aPIHelper;
        private SheetsService sheetsService;
        public SheetController()
        {
            aPIHelper = new APIHelper();
            sheetsService = aPIHelper.service;
        }

        [HttpGet]
        public IActionResult Get(string SPREADSHEET_ID, string SHEET_NAME, string RANGE)
        {


            string RANGE_COMBINE = $"{SHEET_NAME}!{RANGE}";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                   sheetsService.Spreadsheets.Values.Get(SPREADSHEET_ID, RANGE_COMBINE);

            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;


            return Json(values);
        }

        [HttpGet("GetSheetName")]
        public IActionResult GetSheetName(string SPREADSHEET_ID)
        {
            SpreadsheetsResource.GetRequest request = sheetsService.Spreadsheets.Get(SPREADSHEET_ID);
            request.Ranges = new List<string>();
            request.IncludeGridData = false;

            // To execute asynchronously in an async method, replace `request.Execute()` as shown:
            Data.Spreadsheet response = request.Execute();
            // Data.Spreadsheet response = await request.ExecuteAsync();
            var sheetlist = response.Sheets.Select(s => s.Properties.Title);

            return Json(sheetlist);
        }
    }
}
