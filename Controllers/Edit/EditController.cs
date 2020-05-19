using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.Sqlite;

namespace worddemo.Controllers.Edit
{
    public class EditController : Controller
    {

        private string connString;
        private readonly IWebHostEnvironment _webHostEnvironment;
        public EditController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
            string rootPath = _webHostEnvironment.WebRootPath.Replace("/", "\\");
            string dataPath = rootPath.Substring(0, rootPath.Length - 7) + "appData\\" + "Worddemo.db";
            connString = "Data Source=" + dataPath;
        }


        public IActionResult word2()
        {
            string DocID = Request.Query["ID"];
            string sql = "select * from word where id= " + DocID + ";";
            SqliteConnection conn = new SqliteConnection(connString);
            string lz = "张三批阅";//流转
            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;
            SqliteDataReader dr = cmd.ExecuteReader();
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            string Literal_Subject_Text = "";
            string Literal_Lc_Text = "";
            string fileName = "";
            while (dr.Read())
            {
                Literal_Subject_Text = dr["Subject"].ToString();//文件名称
                if ("在线编辑" == dr["Status"].ToString())
                {
                    Literal_Lc_Text = dr["Status"].ToString();//当前文件的流程
                    lz = "张三批阅";//流转
                }
                else
                {
                    Literal_Lc_Text = "已流转到“" + dr["Status"].ToString() + "”，当前是“修改无痕迹模式”打开文件的效果。";
                }
                fileName = dr["FileName"].ToString();
                string fileSubject = dr["Subject"].ToString();
                pageofficeCtrl.Caption = fileSubject;
            }
            //设置保存页面
            pageofficeCtrl.SaveFilePage = "/Edit/SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("/doc/" + fileName, PageOfficeNetCore.OpenModeType.docNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            ViewBag.Literal_Subject_Text = Literal_Subject_Text;
            ViewBag.Literal_Lc_Text = Literal_Lc_Text;
            ViewBag.DocID = DocID;
            ViewBag.lz = lz;
            ViewBag.fileName = fileName;
            return View();

        }

        public IActionResult word3()
        {
            string DocID = Request.Query["ID"];
            string sql = "select * from word where id= " + DocID + ";";
            SqliteConnection conn = new SqliteConnection(connString);
            String lz = "张三批阅";//流转
            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;
            SqliteDataReader dr = cmd.ExecuteReader();
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            string Literal_Subject_Text = "";
            string Literal_Lc_Text = "";
            string fileName = "";
            while (dr.Read())
            {
                Literal_Subject_Text = dr["Subject"].ToString();//文件名称

                if ("正式发文" == dr["Status"].ToString())
                {
                    Literal_Lc_Text = dr["Status"].ToString();//当前文件的流程  
                }
                else
                {
                    Literal_Lc_Text = "已流转到“" + dr["Status"].ToString() + "”，当前是“只读模式”打开文件的效果。";
                }
                fileName = dr["FileName"].ToString();
                string fileSubject = dr["Subject"].ToString();
                pageofficeCtrl.Caption = fileSubject;
            }
            //设置保存页面
            pageofficeCtrl.SaveFilePage = "/Edit/SaveDoc";

            pageofficeCtrl.AddCustomToolButton("另存到本地", "ShowDialog(0)", 5);
            pageofficeCtrl.AddCustomToolButton("页面设置", "ShowDialog(1)", 0);
            pageofficeCtrl.AddCustomToolButton("打印", "ShowDialog(2)", 6);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "IsFullScreen", 4);
            //打开Word文档
            pageofficeCtrl.WebOpen("/doc/" + fileName, PageOfficeNetCore.OpenModeType.docReadOnly, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            ViewBag.Literal_Subject_Text = Literal_Subject_Text;
            ViewBag.Literal_Lc_Text = Literal_Lc_Text;
            ViewBag.DocID = DocID;
            ViewBag.lz = lz;
            ViewBag.fileName = fileName;
            return View();

        }

        public IActionResult word()
        {
            string DocID = Request.Query["ID"];
            string userName = Request.Query["user"];
            string sql = "select * from word where id= " + DocID + ";";
            SqliteConnection conn = new SqliteConnection(connString);
            string lz = "李四批阅";//流转

            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;

            SqliteDataReader dr = cmd.ExecuteReader();
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            string Literal_Subject_Text = "";
            string Literal_Lc_Text = "";
            string fileName = "";
            while (dr.Read())
            {
                Literal_Subject_Text = dr["Subject"].ToString();//文件名称

                //if (("李四批阅" == dr["Status"].ToString() && "李四" == userName) ||("张三批阅" == dr["Status"].ToString() && "张三" == userName))
                if (("李四批阅" == dr["Status"].ToString()) ||("张三批阅" == dr["Status"].ToString()))
                {

                    Literal_Lc_Text = dr["Status"].ToString();//当前文件的流程
                    if ("张三批阅" == Literal_Lc_Text) lz = "李四批阅";
                    if ("李四批阅" == Literal_Lc_Text) lz = "文员清稿";
                }
                else
                {
                    Literal_Lc_Text = "已流转到“" + dr["Status"].ToString() + "”，当前是“强制留痕模式”打开文件的效果。";
                }
                fileName = dr["FileName"].ToString();
                string fileSubject = dr["Subject"].ToString();
                pageofficeCtrl.Caption = fileSubject;
            }
            pageofficeCtrl.CustomMenuCaption = "自定义菜单";
            pageofficeCtrl.AddCustomMenuItem("显示痕迹", "ShowRevisions", false);
            pageofficeCtrl.AddCustomMenuItem("隐藏痕迹", "HiddenRevisions", false);
            pageofficeCtrl.AddCustomMenuItem("-", "", false);
            pageofficeCtrl.AddCustomMenuItem("显示标题", "ShowTitle", true);
            pageofficeCtrl.AddCustomMenuItem("-", "", false);
            pageofficeCtrl.AddCustomMenuItem("领导圈阅", "StartHandDraw", true);
            pageofficeCtrl.AddCustomMenuItem("-", "", false);
            pageofficeCtrl.AddCustomMenuItem("分层显示手写批注", "ShowHandDrawDispBar", true);

            pageofficeCtrl.AddCustomToolButton("保存", "Save", 1);
            pageofficeCtrl.AddCustomToolButton("显示痕迹", "ShowRevisions", 5);
            pageofficeCtrl.AddCustomToolButton("隐藏痕迹", "HiddenRevisions", 5);
            pageofficeCtrl.AddCustomToolButton("领导圈阅", "StartHandDraw", 3);
            pageofficeCtrl.AddCustomToolButton("插入键盘批注", "StartRemark", 3);
            pageofficeCtrl.AddCustomToolButton("分层显示手写批注", "ShowHandDrawDispBar", 7);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "IsFullScreen", 4);

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "/Edit/SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("/doc/" + fileName, PageOfficeNetCore.OpenModeType.docRevisionOnly, userName);
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            ViewBag.Literal_Subject_Text = Literal_Subject_Text;
            ViewBag.Literal_Lc_Text = Literal_Lc_Text;
            ViewBag.DocID = DocID;
            ViewBag.lz = lz;
            ViewBag.fileName = fileName;
            return View();

        }

        public IActionResult word1()
        {
            string DocID = Request.Query["ID"];
            string userName = Request.Query["user"];
            string sql = "select * from word where id= " + DocID + ";";
            SqliteConnection conn = new SqliteConnection(connString);
            string lz = "李四批阅";//流转
            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;

            SqliteDataReader dr = cmd.ExecuteReader();
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            string Literal_Subject_Text = "";
            string Literal_Lc_Text = "";

            string fileName = "";
            while (dr.Read())
            {
                Literal_Subject_Text = dr["Subject"].ToString();//文件名称

                //if (("李四批阅" == dr["Status"].ToString() && "李四" == userName) ||("张三批阅" == dr["Status"].ToString() && "张三" == userName))
                if ("文员清稿" == dr["Status"].ToString() )
                {

                    Literal_Lc_Text = dr["Status"].ToString();//当前文件的流程
                    lz = "正式发文";
                    
                }
                else
                {
                    Literal_Lc_Text = "已流转到“" + dr["Status"].ToString() + "”，当前是“核稿模式”打开文件的效果。";
                }
                fileName = dr["FileName"].ToString();
                string fileSubject = dr["Subject"].ToString();
                pageofficeCtrl.Caption = fileSubject;
            }
            pageofficeCtrl.CustomMenuCaption = "自定义菜单";
            pageofficeCtrl.AddCustomMenuItem("显示痕迹", "ShowRevisions", false);
            pageofficeCtrl.AddCustomMenuItem("隐藏痕迹", "HiddenRevisions", false);
            pageofficeCtrl.AddCustomMenuItem("-", "", false);
            pageofficeCtrl.AddCustomMenuItem("显示标题", "ShowTitle", true);
            pageofficeCtrl.AddCustomMenuItem("-", "", false);
            pageofficeCtrl.AddCustomMenuItem("领导圈阅", "StartHandDraw", true);
            pageofficeCtrl.AddCustomMenuItem("-", "", false);
            pageofficeCtrl.AddCustomMenuItem("分层显示手写批注", "ShowHandDrawDispBar", true);

            pageofficeCtrl.AddCustomToolButton("保存", "Save", 1);
            pageofficeCtrl.AddCustomToolButton("另存为Html", "SaveAsHtml", 0);
            pageofficeCtrl.AddCustomToolButton("显示痕迹", "ShowRevisions", 5);
            pageofficeCtrl.AddCustomToolButton("隐藏痕迹", "HiddenRevisions", 5);
            pageofficeCtrl.AddCustomToolButton("领导圈阅", "StartHandDraw", 3);
            pageofficeCtrl.AddCustomToolButton("插入键盘批注", "StartRemark", 3);
            pageofficeCtrl.AddCustomToolButton("分层显示手写批注", "ShowHandDrawDispBar", 7);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "IsFullScreen", 4);

            //设置保存页面
            pageofficeCtrl.SaveFilePage = "/Edit/SaveDoc";
            //打开Word文档
            pageofficeCtrl.WebOpen("/doc/" + fileName, PageOfficeNetCore.OpenModeType.docRevisionOnly, userName);
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            ViewBag.Literal_Subject_Text = Literal_Subject_Text;
            ViewBag.Literal_Lc_Text = Literal_Lc_Text;
            ViewBag.DocID = DocID;
            ViewBag.lz = lz;
            ViewBag.fileName = fileName;
            return View();

        }

public IActionResult htmldoc()
        {
            string id = Request.Query["id"];
            string docFile = "";

            string sql = "select * from word where id= " + id + ";";
            SqliteConnection conn = new SqliteConnection(connString);
            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;
            SqliteDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                docFile = dr["FileName"].ToString();
                
            }

            docFile = docFile.Substring(0, docFile.Length - 3) + "mht";


            String strsql = "update word set htmlFile='" + docFile
            + "' where id=" + id;

            SqliteCommand cmd2 = new SqliteCommand(strsql, conn);
             
            cmd2.CommandType = CommandType.Text;
            cmd2.ExecuteNonQuery();

            return Redirect("/doc/"+docFile);

        }

        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/doc/" + fs.FileName);
            fs.Close();
            return Content("OK");
        }

        public IActionResult move()
        {
            string id = Request.Query["id"];
            string flg = Request.Query["flg"];

            SqliteConnection conn = new SqliteConnection(connString);
            conn.Open();
            string sql = "Update word set Status = '" + flg + "' where id=" + id;
            SqliteCommand cmd = new SqliteCommand(sql, conn);

            cmd.CommandType = CommandType.Text;
            cmd.ExecuteNonQuery();

            return Redirect("/");

        }


        public IActionResult create()
        {

            SqliteConnection conn = new SqliteConnection(connString);
            conn.Open();
            string sql = "select Max(ID) from word" ;
            SqliteCommand cmd = new SqliteCommand(sql, conn);

            cmd.CommandType = CommandType.Text;
            cmd.ExecuteNonQuery();

            cmd.CommandText = sql;
            SqliteDataReader dr = cmd.ExecuteReader();
            string newID = "1";
            if (dr.Read())
            {
                //newID = ((int)dr[0] + 1).ToString();
                newID = (int.Parse(dr[0].ToString()) + 1).ToString();
            }
            dr.Close();

            string fileName = "aabb" + newID + ".doc";

            string FileSubject = "请输入文档主题";
            if (Request.Query["FileSubject"] != "") FileSubject = Request.Query["FileSubject"];

            String strsql = "Insert into word(ID,FileName,Subject,SubmitTime,Status) values(" + newID
                + ",'" + fileName + "','" + FileSubject + "','" + DateTime.Now.ToString() + "','在线编辑')";

            SqliteCommand cmd2 = new SqliteCommand(strsql, conn);

            cmd2.CommandType = CommandType.Text;
            cmd2.ExecuteNonQuery();
            // 复制服务器端的模板文件命名为新的文件名
            string webRootPath = _webHostEnvironment.WebRootPath;
            System.IO.File.Copy(webRootPath + "\\doc\\"+ Request.Query["TemplateName"],
                webRootPath + "\\doc\\" + fileName, true);
            return Redirect("/");

        }



    }
}