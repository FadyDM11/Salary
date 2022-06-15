using System.Data.OleDb;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;

namespace Salary.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class SalaryController : ControllerBase
    {
        private readonly ILogger<SalaryController> _logger;

        public SalaryController(ILogger<SalaryController> logger)
        {
            _logger = logger;
        }
        [HttpGet(Name = "GrossByNet")]
        public ConverterDto Get(string net)
        {
            var res = new ConverterDto("","","");
            string connString = "Provider= Microsoft.ACE.OLEDB.12.0;" + "Data Source=NetGrosssheet.xlsx" + ";Extended Properties='Excel 8.0;HDR=Yes'";
            using OleDbConnection connection = new OleDbConnection(connString);
            connection.Open();
            OleDbCommand command = new OleDbCommand($"select [Gross Monthly Salary],[Social Insurance],[Taxes Monthly] from [Sheet1$] Where [Net Salary Per Month] >= {net}", connection);
            using (OleDbDataReader dr = command.ExecuteReader())
            {
                while (dr.Read())
                {
                    res = new ConverterDto(Math.Round((double)dr[0], 3).ToString(), Math.Round((double)dr[1], 3).ToString(), Math.Round((double)dr[0], 3).ToString());
                    break;
                }
            }
            connection.Close();
            return res;
        }
    }
}