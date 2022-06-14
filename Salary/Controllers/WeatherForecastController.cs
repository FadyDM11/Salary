using System.Data.OleDb;
using Microsoft.AspNetCore.Mvc;

namespace Salary.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
        "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
    };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet(Name = "GrossByNet")]
        public ConverterDto Get(string Net)
        {
            var res = new ConverterDto("","","");
            string connString = "Provider= Microsoft.ACE.OLEDB.12.0;" + "Data Source=NetGrosssheet.xlsx" + ";Extended Properties='Excel 8.0;HDR=Yes'";
            using (OleDbConnection connection = new OleDbConnection(connString))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand($"select [Gross Monthly Salary],[Social Insurance],[Taxes Monthly] from [Sheet1$] Where [Net Salary Per Month] >= {Net}", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        res = new ConverterDto(dr[0].ToString()!, dr[1].ToString()!, dr[2].ToString()!);
                        //var row1Col0 = dr[0];
                        //var row1Col1 = dr[1];
                        //var row1Col2 = dr[2];
                        //Console.WriteLine(row1Col0);
                        //Console.WriteLine(row1Col1);
                        //Console.WriteLine(row1Col2);
                        break;
                    }
                }
            }




            return res;
        }
    }
}