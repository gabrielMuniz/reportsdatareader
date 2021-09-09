using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using System;
using System.IO;

namespace ReportsDataReader.Api.Controllers
{
    [ApiController]
    [Route("export")]
    public class ExportController : ControllerBase
    {

        private string FilePath = @"C:\temp\MapeamentoGrupo5.xlsx";

        [HttpGet]
        public ActionResult GetReports()
        {

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = System.IO.File.Open(FilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    do
                    {
                        while (reader.Read())
                        {
                           Console.WriteLine(reader.GetFieldType(0));
                        }
                    } while (reader.NextResult());
                }
            }
            return Ok();
        }
    }
}
