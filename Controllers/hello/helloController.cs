using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.Sqlite;

namespace worddemo.Controllers.hello
{
    public class helloController : Controller
    {


        private readonly IWebHostEnvironment _webHostEnvironment;
        private string connString;
        public helloController(IWebHostEnvironment webHostEnvironment)
        {

            _webHostEnvironment = webHostEnvironment;
            string dataPath = _webHostEnvironment.WebRootPath.Replace("/", "\\");
            dataPath = dataPath.Substring(0, dataPath.Length - 7) + "appData\\" + "Worddemo.db";
            connString = "Data Source=" + dataPath;
        }

        public IActionResult Index()
        {
            string sql = "Select * from stream";
            SqliteConnection conn = new SqliteConnection(connString);
            conn.Open();
            return View();
        }
    }
}