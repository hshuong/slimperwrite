using AngleSharp.Parser.Html;
using AngleSharp.Extensions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Npgsql;
using NpgsqlTypes;

namespace slimperwrite
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
        public static void CreateRowNameInStatementRowTable(string path)
        {
            string[] loaibaocao = new string[3] { "bang-can-doi-ke-toan", "bao-cao-ket-qua-kinh-doanh", "bao-cao-luu-chuyen-tien-te" };
            var loaicongty = 1; // cong ty thuong
            using (NpgsqlConnection conn = new NpgsqlConnection("Server=localhost; Port=5432; User Id=postgres; Password=123456; Database=financial"))
            {
                conn.Open();
                var parser = new HtmlParser();
                for (int k = 0; k < loaibaocao.Length; k++)
                {
                    string loai = loaibaocao[k];
                    string html = File.ReadAllText(path + "\\bvh\\bvh_" + loai + "_2017_IN_YEAR.html");
                    var document = parser.Parse(html);
                    var tensolieu = document.QuerySelectorAll("table table tbody tr td:nth-child(1) div");
                    var solieu1 = document.QuerySelectorAll("table table tbody tr td:nth-child(2) div");
                    var solieu2 = document.QuerySelectorAll("table table tbody tr td:nth-child(3) div");
                    var solieu3 = document.QuerySelectorAll("table table tbody tr td:nth-child(4) div");
                    var solieu4 = document.QuerySelectorAll("table table tbody tr td:nth-child(5) div");
                    switch (loai)
                    {
                        case "bang-can-doi-ke-toan":
                            break;
                        case "bao-cao-ket-qua-kinh-doanh":
                        case "bao-cao-luu-chuyen-tien-te":
                            tensolieu = document.QuerySelectorAll("table table tbody tr td:nth-child(1)");
                            solieu1 = document.QuerySelectorAll("table table tbody tr td:nth-child(2)");
                            solieu2 = document.QuerySelectorAll("table table tbody tr td:nth-child(3)");
                            solieu3 = document.QuerySelectorAll("table table tbody tr td:nth-child(4)");
                            solieu4 = document.QuerySelectorAll("table table tbody tr td:nth-child(5)");
                            break;

                    }
                    Console.WriteLine("Bao cao ve " + loai);
                    if (loai == "bang-can-doi-ke-toan") { 
                        switch (tensolieu.Length)
                        {
                            case 80:
                                Console.WriteLine("Ngan hang");
                                loaicongty = 2;
                                break;
                            case 169:
                                Console.WriteLine("Cong ty bao hiem");
                                loaicongty = 3;
                                break;
                            case 123:
                                Console.WriteLine("Cong ty loai thuong");
                                loaicongty = 1;
                                break;

                        }
                    }
                    int thutucuastatementrow = 0;
                    for (int i = 0; i < tensolieu.Length; i++)
                    {
                        var test = tensolieu[i];
                        var rowtitle = "";
                        if (test.ChildElementCount > 0)
                        {
                            rowtitle = test.FirstElementChild.TextContent.Trim();
                            if (String.IsNullOrEmpty(rowtitle)) // truong hop loi the thuong mai truoc 2015, co the img, text nam ngoai
                            {
                                rowtitle = test.TextContent.Trim();
                            }
                            Console.WriteLine(rowtitle + ": " + solieu1[i].TextContent.Trim() + "; " + solieu2[i].TextContent.Trim() + "; " + solieu3[i].TextContent.Trim() + "; " + solieu4[i].TextContent.Trim());
                        }
                        else
                        {
                            rowtitle = test.TextContent.Trim();
                            Console.WriteLine(test.TextContent.Trim() + ": " + solieu1[i].TextContent.Trim() + "; " + solieu2[i].TextContent.Trim() + "; " + solieu3[i].TextContent.Trim() + "; " + solieu4[i].TextContent.Trim());
                        }
                        if (!String.IsNullOrEmpty(rowtitle))
                        {
                            thutucuastatementrow++; // vi co truong hop <td> khong co noi dung thi khong dat so thu tu
                            // Create insert command.
                            NpgsqlCommand command = new NpgsqlCommand("INSERT INTO " +
                                    "statementrow(statementid, companytypeid, rowtitle, roworder) " +
                                    "VALUES(:statementid, :companytypeid, :rowtitle, :roworder)", conn);
                            // Add paramaters.
                            command.Parameters.Add(new NpgsqlParameter("statementid",
                                    NpgsqlTypes.NpgsqlDbType.Integer));
                            command.Parameters.Add(new NpgsqlParameter("companytypeid",
                                    NpgsqlTypes.NpgsqlDbType.Integer));
                            command.Parameters.Add(new NpgsqlParameter("rowtitle",
                                    NpgsqlTypes.NpgsqlDbType.Varchar));
                            command.Parameters.Add(new NpgsqlParameter("roworder",
                                    NpgsqlTypes.NpgsqlDbType.Integer));
                            // Add value to the paramater.
                            command.Parameters[0].Value = k + 1;
                            command.Parameters[1].Value = loaicongty;
                            command.Parameters[2].Value = rowtitle;
                            command.Parameters[3].Value = thutucuastatementrow;
                            // Execute SQL command.
                            command.ExecuteNonQuery();
                        } else
                        {
                            Console.WriteLine(rowtitle);
                        }
                        
                    }

                }
                conn.Close();
            }
            MessageBox.Show("Tao xong bang ten cac khoan muc trong Financial Statement cua loai " + loaicongty.ToString());
        }

        public static void RunSlimpertowritefile(string path, string thamsocacsymbol)
        {
            string[] symbols = new string[] { "hbc", "hpg" };
            List<string> symbolschuacolist = new List<string>();
            // kiem tra co trong csdl chua, chua co moi chay
            if (thamsocacsymbol.Length > 0)
            {
                if (thamsocacsymbol.IndexOf(',') > 0)
                { // co hon 1 symbol
                    symbols = thamsocacsymbol.Split(',');
                }
                else // chi co 1 symbol
                {
                    symbols = new string[] { thamsocacsymbol };
                }
            }
            // kiem tra cac cong ty trong danh sach yeu cau xem
            // da co trong co so du lieu chua.
            for (int j = 0; j < symbols.Length; j++)
            {
                // kiem tra thong tin cong ty, neu co thi khong lam gi, neu chua co thi tao ra cong ty moi
                using (NpgsqlConnection conncheckexist = new NpgsqlConnection("Server=localhost; Port=5432; User Id=postgres; Password=123456; Database=financial"))
                {
                    conncheckexist.Open();
                    NpgsqlCommand commandcheckexist = new NpgsqlCommand("SELECT id, name FROM company where Name = '" + symbols[j] + "'", conncheckexist);
                    NpgsqlDataReader readercheckexist = commandcheckexist.ExecuteReader();

                    if (!readercheckexist.HasRows)
                    {
                        symbolschuacolist.Add(symbols[j]);
                    }
                    conncheckexist.Close();
                }
            }
            string[] symbolschuaco = symbolschuacolist.ToArray();
            if (symbolschuaco.Length > 0) {
                // sua lai tham so, da loai bo cac symbol da co trong database
                thamsocacsymbol = string.Join(",", symbolschuaco);
                // Comment cac dong sau de khong chay slimerjs
                Process proc = null;
                //string _batDir = string.Format(@"I:\web load\slimerjs-1.0.0\");
                string _batDir = path;
                proc = new Process();
                proc.StartInfo.WorkingDirectory = _batDir;
                proc.StartInfo.FileName = "slimerjs.bat";
                proc.StartInfo.Arguments = "nhieuwebnhieuquy.js " + thamsocacsymbol;
                proc.StartInfo.CreateNoWindow = false;
                proc.Start();
                proc.WaitForExit();
                //ExitCode = proc.ExitCode;
                proc.Close();
                // Comment cac dong tren de khong chay slimerjs
                
                int statementid = 1;
                int startstatementid = 0;
                int loaicongty = 1;
                var thutuinsertsql = 0;
                var parser = new HtmlParser();
                for (int j = 0; j < symbolschuaco.Length; j++)
                {
                    // neu khong co thi them cong ty
                    int cocongtyroi = 0;
                    using (NpgsqlConnection conninsert = new NpgsqlConnection("Server=localhost; Port=5432; User Id=postgres; Password=123456; Database=financial"))
                    {
                        conninsert.Open();
                        NpgsqlCommand commandinsert = new NpgsqlCommand("INSERT INTO " +
                                "company(name, companytypeid) " +
                                "VALUES(:name, :companytypeid)", conninsert);
                        // Add paramaters.
                        commandinsert.Parameters.Add(new NpgsqlParameter("name",
                                NpgsqlTypes.NpgsqlDbType.Varchar));
                        commandinsert.Parameters.Add(new NpgsqlParameter("companytypeid",
                                NpgsqlTypes.NpgsqlDbType.Integer));
                        // Add value to the paramater.
                        commandinsert.Parameters[0].Value = symbolschuaco[j];
                        commandinsert.Parameters[1].Value = loaicongty; // tam thoi xem cong ty nay la cong ty thuong
                                                                        // Execute SQL command.
                        commandinsert.ExecuteNonQuery();
                        conninsert.Close();
                        conninsert.Open();
                        commandinsert = new NpgsqlCommand("SELECT id FROM company where Name = '" + symbolschuaco[j] + "'", conninsert);
                        NpgsqlDataReader readerinsert = commandinsert.ExecuteReader();
                        while (readerinsert.Read())
                        {
                            cocongtyroi = Int32.Parse(readerinsert[0].ToString());
                            //do whatever you like
                        }
                        conninsert.Close();
                    }
                    // them cong ty xong moi ghi du lieu bao cao cua cong ty moi
                    string[] htmlFiles = Directory.GetFiles(path + "\\" + symbolschuaco[j], "*.html").Select(Path.GetFileName).ToArray();

                    foreach (var tenfile in htmlFiles)
                    {
                        // tim dau _ thu 2 trong tenfile
                        int vitridaugachduoithu1 = tenfile.IndexOf("_", 0);
                        int vitridaugachduoithu2 = tenfile.IndexOf("_", vitridaugachduoithu1 + 1);
                        int vitrihtml = tenfile.IndexOf(".html", 0);
                        string loaibaocao = tenfile.Substring(vitridaugachduoithu1 + 1, vitridaugachduoithu2 - (vitridaugachduoithu1 + 1));
                        string thoigiansautenloaibaocao = tenfile.Substring(vitridaugachduoithu2 + 1, vitrihtml - (vitridaugachduoithu2 + 1));
                        //if (thoigiansautenloaibaocao.Contains("2016_Q3") && loaibaocao == "bao-cao-luu-chuyen-tien-te")
                        //{
                        //    continue;
                        //}
                        string html = File.ReadAllText(path + "\\" + symbolschuaco[j] + "\\" + tenfile);
                        var document = parser.Parse(html);
                        // lay thoi gian cua bao cao
                        var datebyyear = document.QuerySelectorAll("table table tbody tr").First().Children;
                        // lay ra duoc 6 td, chi co 4 td o giua co thoi gian

                        // lay ten tai khoan va khoi luong
                        var tensolieu = document.QuerySelectorAll("table table tbody tr td:nth-child(1) div");
                        var solieu = document.QuerySelectorAll("table table tbody tr td div").ToArray();
                        //var tencongty = tenfile.Substring(0, 3);
                        switch (loaibaocao)
                        {
                            case "bang-can-doi-ke-toan":
                                statementid = 1;
                                break;
                            case "bao-cao-ket-qua-kinh-doanh":
                                statementid = 2;
                                datebyyear = document.QuerySelectorAll("table table thead tr").First().Children;
                                tensolieu = document.QuerySelectorAll("table table tbody tr td:nth-child(1)");
                                solieu = document.QuerySelectorAll("table table tbody tr td").ToArray();
                                break;
                            case "bao-cao-luu-chuyen-tien-te":
                                statementid = 3;
                                tensolieu = document.QuerySelectorAll("table table tbody tr td:nth-child(1)");
                                solieu = document.QuerySelectorAll("table table tbody tr td").Skip(6).ToArray();
                                break;

                        }
                        Console.WriteLine("Bao cao ve " + loaibaocao);
                        // cap nhat bang cong ty ve loai cong ty
                        if (loaibaocao == "bang-can-doi-ke-toan")
                        {
                            switch (tensolieu.Length)
                            {
                                case 80:
                                    Console.WriteLine("Ngan hang");
                                    loaicongty = 2;
                                    break;
                                case 169:
                                    Console.WriteLine("Cong ty bao hiem");
                                    loaicongty = 3;
                                    break;
                                case 123:
                                    Console.WriteLine("Cong ty loai thuong");
                                    loaicongty = 1;
                                    break;

                            }
                            // update bang cong ty ve loai cong ty
                            using (NpgsqlConnection conn = new NpgsqlConnection("Server=localhost; Port=5432; User Id=postgres; Password=123456; Database=financial"))
                            {
                                conn.Open();
                                NpgsqlCommand commandupdate = new NpgsqlCommand("update company set \"companytypeid\" = :companytypeid where \"id\" = '" + cocongtyroi + "' ;", conn);
                                // Add paramaters.
                                commandupdate.Parameters.Add(new NpgsqlParameter("companytypeid",
                                    NpgsqlTypes.NpgsqlDbType.Integer));
                                commandupdate.Parameters[0].Value = loaicongty; // loai cong ty
                                commandupdate.ExecuteNonQuery();
                                conn.Close();
                            }
                        }
                        // da xac dinh duoc loai cong ty, id bat dau cua cac bang bao cao 
                        // lay thu tu row cua tung loai bao cao ung voi tung loai cong ty
                        // select id from statementrow where companytypeid = 1 and statementid = 3
                        // order by id, statementid, roworder limit 1
                        using (NpgsqlConnection conn = new NpgsqlConnection("Server=localhost; Port=5432; User Id=postgres; Password=123456; Database=financial"))
                        {
                            conn.Open();
                            // lay thu tu row cua tung loai bao cao ung voi tung loai cong ty
                            //select id from statementrow where companytypeid = 1 and statementid = 3
                            //order by id, statementid, roworder limit 1
                            NpgsqlCommand commandcheckexist = new NpgsqlCommand("SELECT id FROM statementrow where companytypeid = '" + loaicongty
                                + "'" + " and statementid = '" + statementid.ToString() + "' order by id, statementid, roworder limit 1", conn);
                            NpgsqlDataReader readercheckexist = commandcheckexist.ExecuteReader();

                            if (readercheckexist.HasRows)
                            {
                                while (readercheckexist.Read())
                                {
                                    startstatementid = Int32.Parse(readercheckexist[0].ToString());
                                    //do whatever you like
                                }
                            }
                            conn.Close();
                        }
                        using (NpgsqlConnection conn = new NpgsqlConnection("Server=localhost; Port=5432; User Id=postgres; Password=123456; Database=financial"))
                        {
                            conn.Open();
                            for (int i = 0; i < tensolieu.Length; i++) // moi mot hang cua bao cao
                            {
                                var test = tensolieu[i];
                                var rowtitle = "";
                                if (test.ChildElementCount > 0)
                                {
                                    rowtitle = test.FirstElementChild.TextContent.Trim();
                                    if (String.IsNullOrEmpty(rowtitle))
                                    // truong hop loi the thuong mai truoc 2015, co the img, text nam ngoai
                                    {
                                        rowtitle = test.TextContent.Trim();
                                    }
                                    //Console.WriteLine(rowtitle + ": " + solieu1[i].TextContent.Trim() + "; " + solieu2[i].TextContent.Trim() + "; " + solieu3[i].TextContent.Trim() + "; " + solieu4[i].TextContent.Trim());
                                }
                                else
                                {
                                    rowtitle = test.TextContent.Trim();
                                    //Console.WriteLine(test.TextContent.Trim() + ": " + solieu1[i].TextContent.Trim() + "; " + solieu2[i].TextContent.Trim() + "; " + solieu3[i].TextContent.Trim() + "; " + solieu4[i].TextContent.Trim());
                                }
                                //trong mot hang cua bao cao, lay ra 5 cot theo thoi gian
                                for (int m = 1; m < datebyyear.Length; m++)
                                {
                                    var namorquy = datebyyear[m].TextContent.Trim();
                                    // neu co thoi gian moi insert vao co so du lieu
                                    // vi co cot khong co gia tri
                                    if (loaibaocao == "bao-cao-luu-chuyen-tien-te" && loaicongty ==1 && i == 54)//71 bao hiem, 72 ngan hang   
                                    {
                                        break;
                                    }
                                    if (loaibaocao == "bao-cao-luu-chuyen-tien-te" && loaicongty == 2 && i == 72)//71 bao hiem, 72 ngan hang   
                                    {
                                        break;
                                    }
                                    if (loaibaocao == "bao-cao-luu-chuyen-tien-te" && loaicongty == 3 && i == 71)//71 bao hiem, 72 ngan hang   
                                    {
                                        break;
                                    }
                                    if (!String.IsNullOrEmpty(namorquy))
                                    {
                                        thutuinsertsql = i * 6 + m;// hang 0 cach hang 1 6 vi tri
                                                                    //Console.WriteLine(rowtitle + " namorquy " + namorquy + ": "+ solieu[thutuinsertsql].TextContent.Trim());
                                                                    //Console.WriteLine(namorquy);
                                                                    // Create insert command.
                                        NpgsqlCommand command = new NpgsqlCommand("INSERT INTO " +
                                            "statementfact(companyid, companytypeid,statementid,statementrowid,date,amount) " +
                                            "VALUES(:companyid, :companytypeid, :statementid, :statementrowid, :date, :amount)", conn);
                                        //cmd.CommandText = "INSERT INTO " +
                                        //        "statementrow(statementid,rowtitle, roworder) " +
                                        //        "VALUES(:statementid, :rowtitle, :roworder)";
                                        // Add paramaters.
                                        command.Parameters.Add(new NpgsqlParameter("companyid",
                                                NpgsqlTypes.NpgsqlDbType.Integer));
                                        command.Parameters.Add(new NpgsqlParameter("companytypeid",
                                                NpgsqlTypes.NpgsqlDbType.Integer));
                                        command.Parameters.Add(new NpgsqlParameter("statementid",
                                                NpgsqlTypes.NpgsqlDbType.Integer));
                                        command.Parameters.Add(new NpgsqlParameter("statementrowid",
                                                NpgsqlTypes.NpgsqlDbType.Integer));
                                        command.Parameters.Add(new NpgsqlParameter("date",
                                                NpgsqlTypes.NpgsqlDbType.Date));
                                        command.Parameters.Add(new NpgsqlParameter("amount",
                                                NpgsqlTypes.NpgsqlDbType.Numeric));
                                        // Prepare the command.
                                        //command.Prepare();
                                        // Add value to the paramater.
                                        //command.Parameters[0].Value = j + 1; // cong ty
                                        command.Parameters[0].Value = cocongtyroi; // cong ty
                                        command.Parameters[1].Value = loaicongty; // companytypeid loai cong ty
                                        command.Parameters[2].Value = statementid; // loai balance
                                        switch (loaibaocao)
                                        {// 124 bat dau income statement, 147 bat dau cash flow, to do: dung sql lay so row cua 1 loai bang
                                            // trong bang statementrow
                                            case "bang-can-doi-ke-toan":
                                                command.Parameters[3].Value = i + startstatementid; // hang thu // ghi vao truong statementrowid
                                                break;
                                            case "bao-cao-ket-qua-kinh-doanh":
                                                command.Parameters[3].Value = i + startstatementid; // hang thu // ghi vao truong statementrowid
                                                break;
                                            case "bao-cao-luu-chuyen-tien-te":
                                                command.Parameters[3].Value = i + startstatementid; // hang thu // ghi vao truong statementrowid
                                                break;
                                        }
                                        if (namorquy.ToUpper().Contains('Q')) // la quy
                                        {
                                            string tinhquytheokytu = namorquy.Substring(namorquy.ToUpper().IndexOf('Q') + 1, 1).Trim();
                                            string namtheoquy = namorquy.Substring(namorquy.ToUpper().IndexOf('Q') + 3).Trim();
                                            command.Parameters[4].Value = new DateTime(Convert.ToInt16(namtheoquy), Convert.ToInt16(tinhquytheokytu) * 3, 30);
                                        }
                                        else
                                        {
                                            command.Parameters[4].Value = new DateTime(Convert.ToInt16(namorquy), 12, 31);
                                        }
                                        //command.Parameters[4].Value = new DateTime(Convert.ToInt16(namorquy), 12, 31); 
                                        // ngay thang
                                        // kiem tra so lieu truoc, neu khong null thi ghi, neu null thi ghi 0
                                        var giatriamount = solieu[thutuinsertsql].TextContent.Trim();
                                        if (!String.IsNullOrEmpty(giatriamount))
                                        {
                                            command.Parameters[5].Value = Convert.ToDouble(giatriamount, System.Globalization.CultureInfo.InvariantCulture); // gia tri
                                        }
                                        else
                                        {
                                            command.Parameters[5].Value = Convert.ToDouble(0); // gia tri
                                        }
                                        command.ExecuteNonQuery();
                                        // neu bao loi dupllicate pkey thi kiem tra file .html bao cao tai chinh xem
                                        // cot thong tin cac con so theo quy co bi sai so voi ten file html hay ko
                                        // vi du neu ten la Q3/2016 thi phai co cot dau tien la Q3/2016, khong phai la Q4/2017
                                    }

                                }
                            }
                            // ket thuc insert du lieu vao bang bao cao
                            conn.Close();
                        }
                    } //foreach
                }
                MessageBox.Show("Đã lấy xong báo cáo tài chính...");
            }
            else
            {
                showInformation("Đã có dữ liệu của các CK này rồi!");
            }
            
        }
        // Show information to message box.
        private static void showInformation(String message)
        {
            MessageBox.Show(message, "Information", MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        public static void getFileNames(string path,string thamsocacsymbol)
        {
            string[] symbols = new string[] { "hbc", "hpg" };
            // kiem tra co trong csdl chua, chua co moi chay
            if (thamsocacsymbol.Length > 0)
            {
                if (thamsocacsymbol.IndexOf(',') > 0)
                { // co hon 1 symbol
                    symbols = thamsocacsymbol.Split(',');
                }
                else // chi co 1 symbol
                {
                    symbols = new string[] { thamsocacsymbol };
                }
            }
            
            for (int j = 0; j < symbols.Length; j++)
            {
                
                string[] htmlFiles = Directory.GetFiles(path + "\\" + symbols[j], "*.html").Select(Path.GetFileName).ToArray();
                
                foreach (var tenfile in htmlFiles)
                {
                    // tim dau _ thu 2 trong tenfile
                    int vitridaugachduoithu1 = tenfile.IndexOf("_",0);
                    int vitridaugachduoithu2 = tenfile.IndexOf("_", vitridaugachduoithu1 + 1);
                    int vitrihtml = tenfile.IndexOf(".html", 0);
                    string loaibaocao = tenfile.Substring(vitridaugachduoithu1 + 1, vitridaugachduoithu2 - (vitridaugachduoithu1 + 1));
                    string thoigiansautenloaibaocao = tenfile.Substring(vitridaugachduoithu2 + 1, vitrihtml - (vitridaugachduoithu2 + 1));
                    string html = File.ReadAllText(path + "\\" + symbols[j] + "\\" + tenfile);
                    
                }
               

            }
        }
        public static void RunSlimpertowritefile_old(string path, string thamsocacsymbol)
        {
            string[] symbols = new string[] { "hbc", "hpg" };
            // kiem tra co trong csdl chua, chua co moi chay
            if (thamsocacsymbol.Length > 0)
            {
                if (thamsocacsymbol.IndexOf(',') > 0)
                { // co hon 1 symbol
                    symbols = thamsocacsymbol.Split(',');
                }
                else // chi co 1 symbol
                {
                    symbols = new string[] { thamsocacsymbol };
                }
            }
            // lay danh sach cong ty co trong database
            using (NpgsqlConnection conncheckexist = new NpgsqlConnection("Server=localhost; Port=5432; User Id=postgres; Password=123456; Database=financial"))
            {
                conncheckexist.Open();
                NpgsqlCommand commandcheckexist = new NpgsqlCommand("SELECT name FROM company", conncheckexist);
                NpgsqlDataReader readercheckexist = commandcheckexist.ExecuteReader();

                if (readercheckexist.HasRows)
                {
                    var daco = -1;
                    while (readercheckexist.Read())
                    {

                        for (int i = 0; i < symbols.Length; i++)
                        {
                            if (readercheckexist[0].ToString() == symbols[i]) // da co khong lam gi, chua co thi them vao symbols
                                daco = i;
                        }
                        if (daco >= 0)
                        {
                            symbols = symbols.Where((source, index) => index != daco).ToArray();
                        }
                        // gan lai daco
                        daco = -1;
                    }
                }
                conncheckexist.Close();
            }
            // sua lai tham so, da loai bo cac symbol da co trong database
            thamsocacsymbol = string.Join(",", symbols);
            if (thamsocacsymbol.Length > 0)
            {
                //// Comment cac dong sau de khong chay slimerjs
                //Process proc = null;
                ////string _batDir = string.Format(@"I:\web load\slimerjs-1.0.0\");
                //string _batDir = path;
                //proc = new Process();
                //proc.StartInfo.WorkingDirectory = _batDir;
                //proc.StartInfo.FileName = "slimerjs.bat";
                //proc.StartInfo.Arguments = "nhieuwebnhieuquy42.js " + thamsocacsymbol;
                //proc.StartInfo.CreateNoWindow = false;
                //proc.Start();
                //proc.WaitForExit();
                ////ExitCode = proc.ExitCode;
                //proc.Close();
                //// Comment cac dong tren de khong chay slimerjs
                string[] loaibaocao = new string[3] { "bang-can-doi-ke-toan", "bao-cao-ket-qua-kinh-doanh", "bao-cao-luu-chuyen-tien-te" };
                var loaicongty = 1; // cong ty thuong
                //string[] thoigian = new string[] { "2016_IN_YEAR", "2017_Q2", "2016_Q1" };
                // chi can chay 1 nam in_year la duoc 5 nam, tu 2017 ve 2013
                string[] thoigian = new string[] { "2017_IN_YEAR", "2017_Q4", "2016_Q3" };
                int cocongtyroi = 0;
                using (NpgsqlConnection conn = new NpgsqlConnection("Server=localhost; Port=5432; User Id=postgres; Password=123456; Database=financial"))
                {
                    conn.Open();
                    var parser = new HtmlParser();
                    for (int j = 0; j < symbols.Length; j++)
                    {
                        // truoc khi kiem tra co cong ty, khoi tao lai bien cocongtyroi = 0
                        cocongtyroi = 0; // de huy ket qua da tao cong ty o buoc j truoc do
                                         // kiem tra thong tin cong ty, neu co thi khong lam gi, neu chua co thi tao ra cong ty moi
                        using (NpgsqlConnection conncheckexist = new NpgsqlConnection("Server=localhost; Port=5432; User Id=postgres; Password=123456; Database=financial"))
                        {
                            conncheckexist.Open();
                            NpgsqlCommand commandcheckexist = new NpgsqlCommand("SELECT id FROM company where Name = '" + symbols[j] + "'", conncheckexist);
                            NpgsqlDataReader readercheckexist = commandcheckexist.ExecuteReader();

                            if (readercheckexist.HasRows)
                            {
                                while (readercheckexist.Read())
                                {
                                    cocongtyroi = Int32.Parse(readercheckexist[0].ToString());
                                    //do whatever you like
                                }
                            }
                            conncheckexist.Close();
                        }
                        if (cocongtyroi == 0) // khong co cong ty
                        {
                            // neu khong co thi them cong ty
                            using (NpgsqlConnection conninsert = new NpgsqlConnection("Server=localhost; Port=5432; User Id=postgres; Password=123456; Database=financial"))
                            {
                                conninsert.Open();
                                NpgsqlCommand commandinsert = new NpgsqlCommand("INSERT INTO " +
                                        "company(name, companytypeid) " +
                                        "VALUES(:name, :companytypeid)", conninsert);
                                // Add paramaters.
                                commandinsert.Parameters.Add(new NpgsqlParameter("name",
                                        NpgsqlTypes.NpgsqlDbType.Varchar));
                                commandinsert.Parameters.Add(new NpgsqlParameter("companytypeid",
                                        NpgsqlTypes.NpgsqlDbType.Integer));
                                // Add value to the paramater.
                                commandinsert.Parameters[0].Value = symbols[j];
                                commandinsert.Parameters[1].Value = loaicongty;
                                // Execute SQL command.
                                commandinsert.ExecuteNonQuery();
                                conninsert.Close();
                                conninsert.Open();
                                commandinsert = new NpgsqlCommand("SELECT id FROM company where Name = '" + symbols[j] + "'", conninsert);
                                NpgsqlDataReader readerinsert = commandinsert.ExecuteReader();
                                while (readerinsert.Read())
                                {
                                    cocongtyroi = Int32.Parse(readerinsert[0].ToString());
                                    //do whatever you like
                                }
                                conninsert.Close();
                            }
                            // them cong ty xong moi ghi du lieu bao cao cua cong ty moi
                            for (int quy = 0; quy < thoigian.Length; quy++)
                            {
                                //
                                for (int k = 0; k < loaibaocao.Length; k++)
                                {
                                    string loai = loaibaocao[k];
                                    int startstatementid = 0;
                                    if (thoigian.Contains("2016_Q3") && loai == "bao-cao-luu-chuyen-tien-te")
                                    {
                                        continue;
                                    }
                                    string html = File.ReadAllText(path + "\\" + symbols[j] + "\\" + symbols[j] + "_" + loai + "_" + thoigian[quy] + ".html");
                                    var document = parser.Parse(html);
                                    // lay thoi gian cua bao cao
                                    var datebyyear = document.QuerySelectorAll("table table tbody tr").First().Children;
                                    // lay ra duoc 6 td, chi co 4 td o giua co thoi gian

                                    // lay ten tai khoan va khoi luong
                                    var tensolieu = document.QuerySelectorAll("table table tbody tr td:nth-child(1) div");
                                    var solieu = document.QuerySelectorAll("table table tbody tr td div").ToArray();
                                    //var solieu1 = document.QuerySelectorAll("table table tbody tr td:nth-child(2) div");
                                    switch (loai)
                                    {
                                        case "bang-can-doi-ke-toan":
                                            break;
                                        case "bao-cao-ket-qua-kinh-doanh":
                                            datebyyear = document.QuerySelectorAll("table table thead tr").First().Children;
                                            tensolieu = document.QuerySelectorAll("table table tbody tr td:nth-child(1)");
                                            solieu = document.QuerySelectorAll("table table tbody tr td").ToArray();
                                            break;
                                        case "bao-cao-luu-chuyen-tien-te":
                                            tensolieu = document.QuerySelectorAll("table table tbody tr td:nth-child(1)");
                                            solieu = document.QuerySelectorAll("table table tbody tr td").Skip(6).ToArray();
                                            //Console.WriteLine("Bao cao ve " + loai);
                                            break;

                                    }
                                    Console.WriteLine("Bao cao ve " + loai);
                                    if (loai == "bang-can-doi-ke-toan")
                                    {
                                        switch (tensolieu.Length)
                                        {
                                            case 80:
                                                Console.WriteLine("Ngan hang");
                                                loaicongty = 2;
                                                break;
                                            case 169:
                                                Console.WriteLine("Cong ty bao hiem");
                                                loaicongty = 3;
                                                break;
                                            case 123:
                                                Console.WriteLine("Cong ty loai thuong");
                                                loaicongty = 1;
                                                break;

                                        }
                                        // update bang cong ty ve loai cong ty
                                        NpgsqlCommand commandupdate = new NpgsqlCommand("update company set \"companytypeid\" = :companytypeid where \"id\" = '" + cocongtyroi + "' ;", conn);
                                        // Add paramaters.
                                        commandupdate.Parameters.Add(new NpgsqlParameter("companytypeid",
                                            NpgsqlTypes.NpgsqlDbType.Integer));
                                        commandupdate.Parameters[0].Value = loaicongty; // loai cong ty
                                        commandupdate.ExecuteNonQuery();
                                    }
                                    // lay thu tu row cua tung loai bao cao ung voi tung loai cong ty
                                    using (NpgsqlConnection conncheckexist = new NpgsqlConnection("Server=localhost; Port=5432; User Id=postgres; Password=123456; Database=financial"))
                                    {
                                        conncheckexist.Open();
                                        //select id from statementrow where companytypeid = 1 and statementid = 3
                                        //order by id, statementid, roworder limit 1
                                        NpgsqlCommand commandcheckexist = new NpgsqlCommand("SELECT id FROM statementrow where companytypeid = '" + loaicongty
                                            + "'" + " and statementid = '" + (k + 1).ToString() + "' order by id, statementid, roworder limit 1", conncheckexist);
                                        NpgsqlDataReader readercheckexist = commandcheckexist.ExecuteReader();

                                        if (readercheckexist.HasRows)
                                        {
                                            while (readercheckexist.Read())
                                            {
                                                startstatementid = Int32.Parse(readercheckexist[0].ToString());
                                                //do whatever you like
                                            }
                                        }
                                        conncheckexist.Close();
                                    }
                                    var thutuinsertsql = 0;
                                    for (int i = 0; i < tensolieu.Length; i++) // moi mot hang cua bao cao
                                    {
                                        var test = tensolieu[i];
                                        var rowtitle = "";
                                        if (test.ChildElementCount > 0)
                                        {
                                            rowtitle = test.FirstElementChild.TextContent.Trim();
                                            if (String.IsNullOrEmpty(rowtitle))
                                            // truong hop loi the thuong mai truoc 2015, co the img, text nam ngoai
                                            {
                                                rowtitle = test.TextContent.Trim();
                                            }
                                            //Console.WriteLine(rowtitle + ": " + solieu1[i].TextContent.Trim() + "; " + solieu2[i].TextContent.Trim() + "; " + solieu3[i].TextContent.Trim() + "; " + solieu4[i].TextContent.Trim());
                                        }
                                        else
                                        {
                                            rowtitle = test.TextContent.Trim();
                                            //Console.WriteLine(test.TextContent.Trim() + ": " + solieu1[i].TextContent.Trim() + "; " + solieu2[i].TextContent.Trim() + "; " + solieu3[i].TextContent.Trim() + "; " + solieu4[i].TextContent.Trim());
                                        }
                                        //trong mot hang cua bao cao, lay ra 5 cot theo thoi gian
                                        for (int m = 1; m < datebyyear.Length; m++)
                                        {
                                            var namorquy = datebyyear[m].TextContent.Trim();
                                            // neu co thoi gian moi insert vao co so du lieu
                                            // vi co cot khong co gia tri
                                            if (loai == "bao-cao-luu-chuyen-tien-te" && i == 54)
                                            {
                                                break;
                                            }
                                            if (!String.IsNullOrEmpty(namorquy))
                                            {
                                                thutuinsertsql = i * 6 + m;// hang 0 cach hang 1 6 vi tri
                                                                           //Console.WriteLine(rowtitle + " namorquy " + namorquy + ": "+ solieu[thutuinsertsql].TextContent.Trim());
                                                                           //Console.WriteLine(namorquy);
                                                                           // Create insert command.
                                                NpgsqlCommand command = new NpgsqlCommand("INSERT INTO " +
                                                    "statementfact(companyid, companytypeid,statementid,statementrowid,date,amount) " +
                                                    "VALUES(:companyid, :companytypeid, :statementid, :statementrowid, :date, :amount)", conn);
                                                //cmd.CommandText = "INSERT INTO " +
                                                //        "statementrow(statementid,rowtitle, roworder) " +
                                                //        "VALUES(:statementid, :rowtitle, :roworder)";
                                                // Add paramaters.
                                                command.Parameters.Add(new NpgsqlParameter("companyid",
                                                        NpgsqlTypes.NpgsqlDbType.Integer));
                                                command.Parameters.Add(new NpgsqlParameter("companytypeid",
                                                        NpgsqlTypes.NpgsqlDbType.Integer));
                                                command.Parameters.Add(new NpgsqlParameter("statementid",
                                                        NpgsqlTypes.NpgsqlDbType.Integer));
                                                command.Parameters.Add(new NpgsqlParameter("statementrowid",
                                                        NpgsqlTypes.NpgsqlDbType.Integer));
                                                command.Parameters.Add(new NpgsqlParameter("date",
                                                        NpgsqlTypes.NpgsqlDbType.Date));
                                                command.Parameters.Add(new NpgsqlParameter("amount",
                                                        NpgsqlTypes.NpgsqlDbType.Numeric));
                                                // Prepare the command.
                                                //command.Prepare();
                                                // Add value to the paramater.
                                                //command.Parameters[0].Value = j + 1; // cong ty
                                                command.Parameters[0].Value = cocongtyroi; // cong ty
                                                command.Parameters[1].Value = loaicongty; // companytypeid loai cong ty
                                                command.Parameters[2].Value = k + 1; // loai balance
                                                switch (loai)
                                                {// 124 bat dau income statement, 147 bat dau cash flow, to do: dung sql lay so row cua 1 loai bang
                                                 // trong bang statementrow
                                                    case "bang-can-doi-ke-toan":
                                                        command.Parameters[3].Value = i + startstatementid; // hang thu // ghi vao truong statementrowid
                                                        break;
                                                    case "bao-cao-ket-qua-kinh-doanh":
                                                        command.Parameters[3].Value = i + startstatementid; // hang thu // ghi vao truong statementrowid
                                                        break;
                                                    case "bao-cao-luu-chuyen-tien-te":
                                                        command.Parameters[3].Value = i + startstatementid; // hang thu // ghi vao truong statementrowid
                                                        break;
                                                }
                                                if (namorquy.ToUpper().Contains('Q')) // la quy
                                                {
                                                    string tinhquytheokytu = namorquy.Substring(namorquy.ToUpper().IndexOf('Q') + 1, 1).Trim();
                                                    string namtheoquy = namorquy.Substring(namorquy.ToUpper().IndexOf('Q') + 3).Trim();
                                                    command.Parameters[4].Value = new DateTime(Convert.ToInt16(namtheoquy), Convert.ToInt16(tinhquytheokytu) * 3, 30);
                                                }
                                                else
                                                {
                                                    command.Parameters[4].Value = new DateTime(Convert.ToInt16(namorquy), 12, 31);
                                                }
                                                //command.Parameters[4].Value = new DateTime(Convert.ToInt16(namorquy), 12, 31); 
                                                // ngay thang
                                                // kiem tra so lieu truoc, neu khong null thi ghi, neu null thi ghi 0
                                                var giatriamount = solieu[thutuinsertsql].TextContent.Trim();
                                                if (!String.IsNullOrEmpty(giatriamount))
                                                {
                                                    command.Parameters[5].Value = Convert.ToDouble(giatriamount, System.Globalization.CultureInfo.InvariantCulture); // gia tri
                                                }
                                                else
                                                {
                                                    command.Parameters[5].Value = Convert.ToDouble(0); // gia tri
                                                }
                                                command.ExecuteNonQuery();
                                                // neu bao loi dupllicate pkey thi kiem tra file .html bao cao tai chinh xem
                                                // cot thong tin cac con so theo quy co bi sai so voi ten file html hay ko
                                                // vi du neu ten la Q3/2016 thi phai co cot dau tien la Q3/2016, khong phai la Q4/2017
                                            }

                                        }
                                    }
                                }
                                //
                            }
                        }
                        // neu cong ty co roi thi khong lam gi


                    }
                    conn.Close();
                }
                MessageBox.Show("Đã lấy xong báo cáo tài chính...");
            }
            else
            {
                showInformation("Đã có dữ liệu của các CK này rồi!");
            }

        }

    }
}
