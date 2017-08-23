using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelParser
{
    public partial class Form1 : Form
    {
        private Excel.Application ex_serv;
        Stopwatch swatch = new Stopwatch();
        Int64 sec = 0;
        Thread pars_thread = null;
        System.Windows.Forms.Timer timer = null;
        char[] separator_size = new char[5] { 'X', 'x', '*', 'Х', 'х' };
        String fileName;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Logger.LogMessage("INFO", "Start application");
            Logger.LogMessage("INFO", "Loading excel server");
            ex_serv = new Excel.Application();
            Logger.LogMessage("INFO", "Loaded excel server");
            timer = new System.Windows.Forms.Timer();
            timer.Tick += Timer_Tick;
            timer.Interval = 1000;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (pars_thread != null)
                pars_thread.Abort();
            ex_serv.Visible = true;
            Logger.LogMessage("INFO", "Closing excel server");
            ex_serv.Quit();
            Logger.LogMessage("INFO", "Closed excel server");
        }

        private void AddHeaderColumn(Excel.Worksheet dst, String range, String text)
        {
            Excel.Range rng = dst.get_Range(range);
            rng.Value2 = text;
            rng.Font.Bold = true;
            rng.HorizontalAlignment = Excel.Constants.xlCenter;
            rng.VerticalAlignment = Excel.Constants.xlCenter;
            rng.EntireColumn.ColumnWidth = 8.38;
            rng.WrapText = true;
            rng.Font.Size = 9;
            rng.Interior.Color = Color.FromArgb(245, 245, 220);
            rng.Borders.ColorIndex = 0;
            rng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            rng.Borders.Weight = Excel.XlBorderWeight.xlThin;
        }


        private void PrintHeader(Excel.Worksheet dst)
        {
            AddHeaderColumn(dst, "A1", "Дата");
            AddHeaderColumn(dst, "B1", "G31_11 (Наименование фирмы изготовителя)");
            AddHeaderColumn(dst, "C1", "G022 (Наименование отправителя)");
            AddHeaderColumn(dst, "D1", "G082 (Наименование получателя)");
            AddHeaderColumn(dst, "E1", "G17B (Страна назначения)");
            AddHeaderColumn(dst, "F1", "G202 - Условия поставки");
            AddHeaderColumn(dst, "G1", "G2021 (пункт поставки товара)");
            AddHeaderColumn(dst, "H1", "Сорт");
            AddHeaderColumn(dst, "I1", "Марка");
            AddHeaderColumn(dst, "J1", "Длина");
            AddHeaderColumn(dst, "K1", "Ширина");
            AddHeaderColumn(dst, "L1", "Толщина");
            AddHeaderColumn(dst, "M1", "Объем, м3");
            AddHeaderColumn(dst, "N1", "Средняя фактурная");
            AddHeaderColumn(dst, "O1", "G42 (Фактурная стоимость товара)");
            AddHeaderColumn(dst, "P1", "G221 (Код валюты фактурной стоимости)");
            AddHeaderColumn(dst, "Q1", "Средняя статистическая");
            AddHeaderColumn(dst, "R1", "G46 (Статистическая стоимость товара в USD)");
        }

        private int GetColumnNumber(Excel.Worksheet src, string prefix)
        {
            int cur_column = 1;
            Excel.Range rng = src.Cells[1, 1];
            while (rng.Value != null)
            {
                string tmp = rng.Value.ToString();
                for (int i = 0; i < prefix.Length; ++i)
                {
                    if (i >= tmp.Length)
                        break;
                    if (prefix[i] != tmp[i])
                        break;
                    if (i == prefix.Length - 1)
                        return cur_column;
                }
                cur_column++;
                rng = src.Cells[1, cur_column];
            }
            return -1;
        }

        private void Parsing()
        {
            Logger.LogMessage("INFO", "Start parsing");
            Excel.Worksheet src = ex_serv.Workbooks[1].Worksheets[1];
            Excel.Worksheet dst = ex_serv.Workbooks[2].Worksheets[1];
            PrintHeader(dst);
            Excel.Worksheet err = ex_serv.Workbooks[3].Worksheets[1];
            src.Cells[1, 1].EntireRow.Copy(err.Cells[1, 1].EntireRow);
            dst.Cells[1, 1].EntireColumn.ColumnWidth = 9.5;

            //инициализация прогресс-бара
            int srcMaxRows = 0;
            while (src.Cells[srcMaxRows + 1, 1].Value != null)
                srcMaxRows++;
            pg_bar.Invoke((MethodInvoker)delegate
            {
                pg_bar.Maximum = srcMaxRows - 1;
                pg_bar.Value = 0;
            });
            l_nproc.Invoke((MethodInvoker)delegate
            {
                l_nproc.Text = l_all.Text = (srcMaxRows - 1).ToString();
            });

            //поиск столбцов
            //Дата G072
            int dateColumn = GetColumnNumber(src, "G072");
            //G31_11 (Наименование фирмы изготовителя)
            int G31_11 = GetColumnNumber(src, "G31_11");
            //G022(Наименование отправителя)
            int G022 = GetColumnNumber(src, "G022");
            //G082(Наименование получателя)
            int G082 = GetColumnNumber(src, "G082");
            //G17B(Страна назначения)
            int G17 = GetColumnNumber(src, "G17");
            //G202 - Условия поставки
            int G202 = GetColumnNumber(src, "G202");
            //G2021(пункт поставки товара)
            int G2021 = GetColumnNumber(src, "G2021");
            //G42(Фактурная стоимость товара)
            int G42 = GetColumnNumber(src, "G42");
            //G221(Код валюты фактурной стоимости)
            int G221 = GetColumnNumber(src, "G221");
            //G46(Статистическая стоимость товара в USD)
            int G46 = GetColumnNumber(src, "G46");
            //G31_1 - Наименование
            int G31_1 = GetColumnNumber(src, "G31_1");
            //G5441 - Брокер Ф.И.О.
            int G5441 = GetColumnNumber(src, "G5441");

            int dst_row = 2;
            int err_row = 2;

            int lproc = 0;
            long sumTicks = 0;

            List<string> sep_sort = new List<string>();
            for (int i = 1; i < 50; ++i)
            {
                sep_sort.Add(String.Format("{0}. ", i));
            }

            List<string> izg_sep = new List<string>();
            for (int i = 1; i < 30; ++i)
                for (int j = 1; j < 30; ++j)
                    izg_sep.Add(String.Format("{0}.{1}. ", i, j));

            string[] dbr = new string[] { "DBR 120/120", "DBR120/120", "DBR 220/220", "DBR220/220",
                                    "S2S", "DB 120/120", "DB120/120", "DB220/220", "DB 220/220",
                "DECK350/DECK350", "DB LOGO 120/120",
                "DB 120/220", "DBPERIBRCH120/DBPERIBRC", "120/120 DB LOGO", "GRN120/120", "GRY 220/220", "BLK120/120", "DBR 120/220", "LBR 120/120", "LBR120/120:*2", "LBR120/120"};

            string[] separator = new string[1] { "Изготовитель" };
            for (int src_row = 2; src_row <= srcMaxRows; src_row++)
            {
                swatch.Reset();
                swatch.Start();
                string nonPars = src.Cells[src_row, G31_1].Value;
                int tl;
                while ((tl = nonPars.IndexOf("_[=")) != -1)
                {
                    int tr = nonPars.IndexOf("=]", tl);
                    nonPars = nonPars.Remove(tl, tr - tl + 3);
                }
                foreach (var c in dbr)
                {
                    nonPars = nonPars.Replace(c, "");
                }
                try
                {
                    dst.Cells[dst_row, 1].Value = src.Cells[src_row, dateColumn].Value;
                    dst.Cells[dst_row, 2].Value = src.Cells[src_row, G31_11].Value;
                    dst.Cells[dst_row, 3].Value = src.Cells[src_row, G022].Value;
                    dst.Cells[dst_row, 4].Value = src.Cells[src_row, G082].Value;
                    dst.Cells[dst_row, 5].Value = src.Cells[src_row, G17].Value;
                    dst.Cells[dst_row, 6].Value = src.Cells[src_row, G202].Value;
                    dst.Cells[dst_row, 7].Value = src.Cells[src_row, G2021].Value;
                    dst.Cells[dst_row, 8].EntireColumn.NumberFormat = "@";
                    dst.Cells[dst_row, 9].EntireColumn.NumberFormat = "@";
                    dst.get_Range(String.Format("{0}1", (char)('A' + 10)), String.Format("{0}65535", (char)('A' + 10))).NumberFormat = "#.0####";
                    dst.get_Range(String.Format("{0}1", (char)('A' + 11)), String.Format("{0}65535", (char)('A' + 11))).NumberFormat = "#.0####";
                    dst.get_Range(String.Format("{0}1", (char)('A' + 12)), String.Format("{0}65535", (char)('A' + 12))).NumberFormat = "#.0####";
                    dst.get_Range(String.Format("{0}1", (char)('A' + 13)), String.Format("{0}65535", (char)('A' + 13))).NumberFormat = "#.0####";
                    dst.Cells[dst_row, 14].Formula = String.Format("=O{0}/M{0}", dst_row);
                    dst.Cells[dst_row, 15].Value = src.Cells[src_row, G42].Value;
                    dst.Cells[dst_row, 16].Value = src.Cells[src_row, G221].Value;
                    dst.Cells[dst_row, 17].Formula = String.Format("=R{0}/M{0}", dst_row);
                    dst.Cells[dst_row, 18].Value = src.Cells[src_row, G46].Value;
                    dst.Cells[dst_row, 14].Formula = String.Format("=O{0}/M{0}", dst_row);
                    dst.Cells[dst_row, 15].Value = src.Cells[src_row, G42].Value;
                    dst.Cells[dst_row, 16].Value = src.Cells[src_row, G221].Value;
                    dst.Cells[dst_row, 17].Formula = String.Format("=R{0}/M{0}", dst_row);
                    dst.Cells[dst_row, 18].Value = src.Cells[src_row, G46].Value;


                    switch ((string)src.Cells[src_row, G5441].Value2.ToString())
                    {
                        case "ЦЫГАНКОВ":
                        case "НЕРОНОВА":
                        case "ЗАЙЦЕВ":
                        case "ПРОЗОРОВ":
                        case "ПАШЕНЦЕВА":
                        case "МОРОЗОВА":
                        case "МИХАЙЛОВА":
                            {
                                int indst = nonPars.IndexOf(')') + 1;
                                int indfin = nonPars.IndexOf(":_1.1", indst);
                                string tmp = nonPars.Substring(indst, indfin - indst).Trim().Replace("_1.", "").Replace(" L ", "");
                                char[] seps = new char[] { ';', '.', '(', ')', '-', ':' };
                                dst.Cells[dst_row, 9].Value = get_value(nonPars, "марка");
                                int tr = -1;
                                while ((tr = tmp.IndexOf(" м3")) != -1)
                                {
                                    tmp = tmp.Replace(" м3", "м3");
                                }

                                while ((tr = tmp.IndexOf(" М3")) != -1)
                                {
                                    tmp = tmp.Replace(" М3", "М3");
                                }
                                while ((tr = tmp.IndexOf(" m3")) != -1)
                                {
                                    tmp = tmp.Replace(" m3", "m3");
                                }
                                while ((tr = tmp.IndexOf(" M3")) != -1)
                                {
                                    tmp = tmp.Replace(" M3", "M3");
                                }
                                while ((tr = tmp.IndexOf(" mm")) != -1)
                                {
                                    tmp = tmp.Replace(" mm", "mm");
                                }
                                while ((tr = tmp.IndexOf(" MM")) != -1)
                                {
                                    tmp = tmp.Replace(" MM", "MM");
                                }
                                foreach (var c in new string[] { "мм", "ММ", "mm", "MM", "m3", "M3", "м3", "М3", "*" })
                                {
                                    int k = 0;
                                    while ((k = tmp.IndexOf(c, k + c.Length)) != -1)
                                    {
                                        StringBuilder tmpsb = new StringBuilder(tmp);
                                        int i = k - 1;
                                        while (i > 0 && char.IsDigit(tmp[i])) i--;
                                        if (tmp[i] == ' ')
                                        {
                                            tmpsb[i] = '-';
                                        }
                                        i = k + 1;
                                        while (i > 0 && char.IsDigit(tmp[i])) i++;
                                        if (tmp[i] == ' ')
                                        {
                                            tmpsb[i] = '-';
                                        }
                                        tmp = tmpsb.ToString();
                                    }
                                }
                                List<string> strs = tmp.Split(seps, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim().Trim(',').Trim()).ToList();

                                bool sz = false, srt = false, w = false, v = false;
                                string[] add_sort = new string[] { "1/2", "1/1", "1/3" };
                                string sort = "";
                                foreach (string str in strs)
                                {
                                    if (sz && srt && w && v)
                                    {
                                        dst_row++;
                                        dst.Cells[dst_row - 1, 1].EntireRow.Copy(dst.Cells[dst_row, 1].EntireRow);
                                        v = false;
                                        if (srt)
                                            dst.Cells[dst_row, 8].Value = sort;
                                    }
                                    if (add_sort.Contains(str))
                                    {
                                        dst.Cells[dst_row, 8].Value = sort + " " + str;
                                        srt = true;
                                        v = false;
                                        continue;
                                    }
                                    string low = str.ToLower();
                                    if (low.Contains("м3") || low.Contains("m3")) //объем
                                    {
                                        dst.Cells[dst_row, 13].Value2 = get_numeric(str);
                                        v = true;
                                    }
                                    else if (low.Contains("mm") || low.Contains("мм")) //толщина
                                    {
                                        dst.Cells[dst_row, 12].Value2 = get_numeric(str);
                                        w = true;
                                        v = false;
                                    }
                                    else if (check_size_string(low)) //размер
                                                                     //else if (low.Contains("*") || low.Contains("x") || low.Contains("х")) //размер
                                    {
                                        var s = low.Split(separator_size).Select(x => x.Trim()).ToArray();
                                        dst.Cells[dst_row, 10].Value2 = get_numeric(s[0]);
                                        dst.Cells[dst_row, 11].Value2 = get_numeric(s[1]);
                                        sz = true;
                                        v = false;
                                    }
                                    else //sort
                                    {
                                        dst.Cells[dst_row, 8].Value = str;
                                        sort = str;
                                        srt = true;
                                        v = false;
                                    }
                                }

                                if (!v)
                                {
                                    dst.Cells[dst_row, 13].Value2 = src.Cells[src_row, GetColumnNumber(src, "G31_7")].Value2.Replace(',', '.'); ;
                                }
                            }
                            break;
                        case "МИНАКОВА":
                        case "ЕФИМОВА":
                            {
                                var temps = nonPars.Split(separator, StringSplitOptions.RemoveEmptyEntries).ToList();
                                temps.RemoveAt(0);
                                dst.Cells[dst_row, 8].Value = nonPars;
                                bool bfirst = true;


                                dst_row--;
                                foreach (var tmp in temps)
                                {
                                    dst_row++;
                                    if (!bfirst)
                                        dst.Cells[dst_row - 1, 1].EntireRow.Copy(dst.Cells[dst_row, 1].EntireRow);
                                    dst.Cells[dst_row, 8].Value = get_value(tmp, "Сорт");
                                    dst.Cells[dst_row, 9].Value = get_value(tmp, "Марка");
                                    var string_size = get_value(tmp, "Размер");
                                    if (string_size != "ОТСУТСТВУЕТ")
                                    {
                                        if (string_size.IndexOf(' ') != -1)
                                            string_size = string_size.Substring(0, string_size.IndexOf(' '));
                                        while (string_size.Length > 0 && !char.IsDigit(string_size.Last()))
                                            string_size = string_size.Remove(string_size.Length - 1);

                                    }
                                    var ssize = string_size.Split(separator_size);
                                    dst.Cells[dst_row, 10].Value2 = ssize[0].Replace(',', '.');
                                    dst.Cells[dst_row, 11].Value2 = ssize[1].Replace(',', '.');
                                    dst.Cells[dst_row, 12].Value2 = get_numeric(nonPars, nonPars.IndexOf("ТОЛЩИНА")); ;

                                    string volume = get_value(tmp, "Кол-во");
                                    volume = volume.Substring(0, volume.LastIndexOf(' ')).Replace(',', '.');
                                    if (volume == "ОТСУТСТВУЕТ")
                                        volume = src.Cells[src_row, GetColumnNumber(src, "G31_7")].Value2.Replace(',', '.');
                                    dst.Cells[dst_row, 13].Value2 = volume;
                                    bfirst = false;
                                }
                            }
                            break;
                        case "КИСЕЛЕВА":
                            {
                                throw new Exception("КИСЕЛЕВА");
                            }
                            break;
                        case "КОНСТАНТИНОВ":
                        case "ЛАЗУТКИНА":
                        case "МОХОВА":
                        case "ПОПОВ":
                            {
                                try
                                {
                                    var temps = nonPars.Split(separator, StringSplitOptions.RemoveEmptyEntries).ToList();
                                    temps.RemoveAt(0);
                                    dst.Cells[dst_row, 8].Value = nonPars;
                                    bool bfirst = true;


                                    dst_row--;
                                    foreach (var tmp in temps)
                                    {
                                        dst_row++;
                                        if (!bfirst)
                                            dst.Cells[dst_row - 1, 1].EntireRow.Copy(dst.Cells[dst_row, 1].EntireRow);
                                        string srt = get_value(tmp, "Сорт");
                                        if (srt == "ОТСУТСТВУЕТ")
                                            throw new Exception("Нет сорта");
                                        dst.Cells[dst_row, 8].Value = get_value(tmp, "Сорт");
                                        dst.Cells[dst_row, 9].Value = get_value(tmp, "Марка");
                                        var string_size = get_value(tmp, "Размер");
                                        if (string_size != "ОТСУТСТВУЕТ")
                                        {
                                            if (string_size.IndexOf(' ') != -1)
                                                string_size = string_size.Substring(0, string_size.IndexOf(' '));
                                            while (string_size.Length > 0 && !char.IsDigit(string_size.Last()))
                                                string_size = string_size.Remove(string_size.Length - 1);

                                        }
                                        var ssize = string_size.Split(separator_size);
                                        dst.Cells[dst_row, 10].Value2 = ssize[0].Replace(',', '.');
                                        dst.Cells[dst_row, 11].Value2 = ssize[1].Replace(',', '.');
                                        for (int i = 0; i < ssize[2].Length; ++i)
                                        {
                                            if (char.IsDigit(ssize[2][i]) || ssize[2][i] == ',' || ssize[2][i] == '.') continue;
                                            ssize[2] = ssize[2].Substring(0, i);
                                            break;
                                        }
                                        dst.Cells[dst_row, 12].Value2 = ssize[2].Replace(',', '.');

                                        string volume = get_value(tmp, "Кол-во");
                                        volume = volume.Substring(0, volume.LastIndexOf(' ')).Replace(',', '.');
                                        if (volume == "ОТСУТСТВУЕТ")
                                            volume = src.Cells[src_row, GetColumnNumber(src, "G31_7")].Value2.Replace(',', '.');
                                        dst.Cells[dst_row, 13].Value2 = volume;
                                        bfirst = false;
                                    }
                                }
                                catch (Exception)
                                {
                                    try
                                    {
                                        var divs = nonPars.Split(new char[] { '_' }).Select(x => x.Trim()).ToList();
                                        if (char.IsDigit(divs[0].Last()))
                                            dst.Cells[dst_row, 9].Value = "ОТСУТСТВУЕТ";
                                        else dst.Cells[dst_row, 9].Value = divs[0].Substring(divs[0].LastIndexOf(' ')).Trim();

                                        var sorts = divs[1].Split(sep_sort.ToArray(), StringSplitOptions.RemoveEmptyEntries);
                                        var szw = divs[2].Split(izg_sep.ToArray(), StringSplitOptions.RemoveEmptyEntries);
                                        for (int i = 0; i < sorts.Length; ++i)
                                        {
                                            if (i != 0)
                                                dst.Cells[dst_row - 1, 1].EntireRow.Copy(dst.Cells[dst_row, 1].EntireRow);
                                            string w = get_numeric(sorts[i]);
                                            dst.Cells[dst_row, 12].Value2 = w;
                                            dst.Cells[dst_row, 8].Value = sorts[i].Substring(w.Length);

                                            string sz_str = get_value(szw[i], "Размер");
                                            string dl = get_numeric(sz_str);
                                            dst.Cells[dst_row, 10].Value2 = dl;
                                            dst.Cells[dst_row, 11].Value2 = get_numeric(sz_str, dl.Length);

                                            string volume = get_value(szw[i], "Кол-во");
                                            volume = volume.Substring(0, volume.LastIndexOf(' ')).Replace(',', '.');
                                            if (volume == "ОТСУТСТВУЕТ")
                                                volume = src.Cells[src_row, GetColumnNumber(src, "G31_7")].Value2.Replace(',', '.');
                                            dst.Cells[dst_row, 13].Value2 = volume;
                                            if (i + 1 != sorts.Length)
                                                dst_row++;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Logger.LogMessage(ex);
                                        string v = get_weight(nonPars);
                                        int i = nonPars.IndexOf("СТО");
                                        if (i == -1)
                                            i = nonPars.IndexOf("CTO");
                                        i += 3;

                                        while (i < nonPars.Length && nonPars[i] == ' ') i++;
                                        while (i < nonPars.Length && char.IsDigit(nonPars[i]) || nonPars[i] == '-') i++;
                                        while (i < nonPars.Length && nonPars[i] == ' ') i++;
                                        string srt = "";
                                        string mrk = "ОТСУТСТВУЕТ";
                                        int ind = nonPars.IndexOf("СОРТ");
                                        if (ind == -1)
                                            ind = nonPars.IndexOf("COPT");
                                        if (i == ind) //сорт
                                            srt = nonPars.Substring(i + 4, nonPars.IndexOf(v) - i - 5);
                                        else if (nonPars[i] == 'М' || nonPars[i] == 'M') //марка
                                        {
                                            mrk = nonPars.Substring(i + 6, nonPars.IndexOf(' ', i + 6) - i - 6);
                                            srt = nonPars.Substring(i + 6 + mrk.Length + 1, nonPars.IndexOf(v) - (i + 6 + mrk.Length + 1) - 1);
                                        }
                                        else if (ind != -1)
                                        {
                                            mrk = nonPars.Substring(i, nonPars.IndexOf(' ', i) - i);
                                            srt = nonPars.Substring(ind + 5, nonPars.IndexOf(v) - ind - 5);
                                        }
                                        else
                                        {
                                            srt = nonPars.Substring(i, nonPars.IndexOf(v) - i);
                                        }
                                        string w = get_numeric(srt);
                                        srt = srt.Substring(srt.IndexOf(w) + w.Length).Trim();
                                        string sz;
                                        string[] vsz;
                                        int l = nonPars.ToLower().IndexOf("размер");
                                        int r = nonPars.ToLower().IndexOf("мм");
                                        if (r == -1)
                                            r = nonPars.ToLower().IndexOf("mm");
                                        if (r - l - 6 < 5)
                                        {
                                            r = nonPars.ToLower().IndexOf("размер");
                                            l = nonPars.ToLower().IndexOf("m3") + 2;
                                            if (l == -1)
                                                l = nonPars.ToLower().IndexOf("м3");
                                        }
                                        sz = nonPars.Substring(l, r - l - 1);
                                        vsz = sz.Split(separator_size, StringSplitOptions.RemoveEmptyEntries);
                                        dst.Cells[dst_row, 8].Value = srt;
                                        dst.Cells[dst_row, 9].Value = mrk.Trim();
                                        dst.Cells[dst_row, 10].Value2 = get_numeric(vsz[0]);
                                        dst.Cells[dst_row, 11].Value2 = get_numeric(vsz[1]);
                                        dst.Cells[dst_row, 12].Value2 = w;
                                        dst.Cells[dst_row, 13].Value2 = v;
                                    }
                                }
                            }
                            break;
                        case "ГОЛИКОВА":
                        case "ЗАТЕЙ":
                        case "ПОКАЗИЙ":
                        case "АСТАПОВ":
                            {
                                throw new Exception("СЛОЖНО");
                            }
                            break;
                        case "МАЛЫШЕВА":
                        case "ОВЧИННИКОВА":
                        case "РОЩЕКТАЕВА":
                            {
                                throw new Exception("НАО СВЕЗА ВЕРХНЯЯ СИНЯЧИХА");
                            }
                            break;
                        case "СИТДИКОВА":
                        case "ФИЛЮШИНА":
                            {
                                var temps = nonPars.Split(separator, StringSplitOptions.RemoveEmptyEntries).ToList();
                                temps.RemoveAt(0);
                                dst.Cells[dst_row, 8].Value = nonPars;
                                bool bfirst = true;
                                List<List<string>> toex = new List<List<string>>();
                                dst_row--;
                                foreach (var tmp in temps)
                                {
                                    dst_row++;
                                    if (!bfirst)
                                        dst.Cells[dst_row - 1, 1].EntireRow.Copy(dst.Cells[dst_row, 1].EntireRow);
                                    dst.Cells[dst_row, 8].Value = get_value(tmp, "Сорт");
                                    dst.Cells[dst_row, 9].Value = get_value(tmp, "Марка");
                                    var string_size = get_value(tmp, "Размер");
                                    if (string_size != "ОТСУТСТВУЕТ")
                                    {
                                        if (string_size.IndexOf(' ') != -1)
                                            string_size = string_size.Substring(0, string_size.IndexOf(' '));
                                        while (string_size.Length > 0 && !char.IsDigit(string_size.Last()))
                                            string_size = string_size.Remove(string_size.Length - 1);

                                    }
                                    dst.Cells[dst_row, 10].Value2 = get_numeric(tmp, tmp.IndexOf("ДЛИНА")); ;
                                    dst.Cells[dst_row, 11].Value2 = get_numeric(tmp, tmp.IndexOf("ШИРИНА")); ;
                                    dst.Cells[dst_row, 12].Value2 = get_numeric(tmp, tmp.IndexOf("ТОЛЩИНА")); ;

                                    string volume = get_value(tmp, "Кол-во");
                                    volume = volume.Substring(0, volume.LastIndexOf(' ')).Replace(',', '.');
                                    if (volume == "ОТСУТСТВУЕТ")
                                        volume = src.Cells[src_row, GetColumnNumber(src, "G31_7")].Value2.Replace(',', '.');
                                    dst.Cells[dst_row, 13].Value2 = volume;
                                    bfirst = false;
                                }
                            }
                            break;
                        case "СЕРГЕЕВА":
                            {
                                var temps = nonPars.Split(separator, StringSplitOptions.RemoveEmptyEntries).ToList();
                                temps.RemoveAt(0);
                                dst.Cells[dst_row, 8].Value = nonPars;
                                bool bfirst = true;


                                dst_row--;
                                foreach (var tmp in temps)
                                {
                                    dst_row++;
                                    if (!bfirst)
                                        dst.Cells[dst_row - 1, 1].EntireRow.Copy(dst.Cells[dst_row, 1].EntireRow);
                                    dst.Cells[dst_row, 8].Value = get_value(tmp, "Сорт");
                                    dst.Cells[dst_row, 9].Value = get_value(tmp, "Марка");
                                    dst.Cells[dst_row, 10].Value2 = get_numeric(tmp, tmp.IndexOf("ДЛИНА")); ;
                                    dst.Cells[dst_row, 11].Value2 = get_numeric(tmp, tmp.IndexOf("ШИРИНА")); ;
                                    dst.Cells[dst_row, 12].Value2 = get_numeric(tmp, tmp.IndexOf("ТОЛЩИНА")); ;

                                    string volume = get_value(tmp, "Кол-во");
                                    volume = volume.Substring(0, volume.LastIndexOf(' ')).Replace(',', '.');
                                    if (volume == "ОТСУТСТВУЕТ")
                                        volume = src.Cells[src_row, GetColumnNumber(src, "G31_7")].Value2.Replace(',', '.');
                                    dst.Cells[dst_row, 13].Value2 = volume;
                                    bfirst = false;
                                }
                            }
                            break;
                        default:
                            {
                                try
                                {
                                    var temps = nonPars.Split(separator, StringSplitOptions.RemoveEmptyEntries).ToList();
                                    temps.RemoveAt(0);
                                    dst.Cells[dst_row, 8].Value = nonPars;
                                    bool bfirst = true;


                                    dst_row--;
                                    foreach (var tmp in temps)
                                    {
                                        dst_row++;
                                        if (!bfirst)
                                            dst.Cells[dst_row - 1, 1].EntireRow.Copy(dst.Cells[dst_row, 1].EntireRow);
                                        dst.Cells[dst_row, 8].Value = get_value(tmp, "Сорт");
                                        dst.Cells[dst_row, 9].Value = get_value(tmp, "Марка");
                                        var string_size = get_value(tmp, "Размер");
                                        if (string_size != "ОТСУТСТВУЕТ")
                                        {
                                            if (string_size.IndexOf(' ') != -1)
                                                string_size = string_size.Substring(0, string_size.IndexOf(' '));
                                            while (string_size.Length > 0 && !char.IsDigit(string_size.Last()))
                                                string_size = string_size.Remove(string_size.Length - 1);

                                        }
                                        var ssize = string_size.Split(separator_size);
                                        dst.Cells[dst_row, 10].Value2 = ssize[0].Replace(',', '.');
                                        dst.Cells[dst_row, 11].Value2 = ssize[1].Replace(',', '.');
                                        for (int i = 0; i < ssize[2].Length; ++i)
                                        {
                                            if (char.IsDigit(ssize[2][i]) || ssize[2][i] == ',' || ssize[2][i] == '.') continue;
                                            ssize[2] = ssize[2].Substring(0, i);
                                            break;
                                        }
                                        dst.Cells[dst_row, 12].Value2 = ssize[2].Replace(',', '.');

                                        string volume = get_value(tmp, "Кол-во");
                                        volume = volume.Substring(0, volume.LastIndexOf(' ')).Replace(',', '.');
                                        if (volume == "ОТСУТСТВУЕТ")
                                            volume = src.Cells[src_row, GetColumnNumber(src, "G31_7")].Value2.Replace(',', '.');
                                        dst.Cells[dst_row, 13].Value2 = volume;
                                        bfirst = false;
                                    }
                                }
                                catch (Exception)
                                {
                                    /*
                                     * ДЕТАЛЬ МЕБЕЛЬНАЯ ИЗ КЛЕЕНОЙ БЕРЕЗОВОЙ ФАНЕРЫ СОРТ А/В, 9-ТИ СЛОЙНАЯ ТОЛЩИНОЙ 10ММ:700Х50Х10ММ-29,400М3(84000ШТ)._
                                     * [=1=] :_[=1.1=]  Изготовитель: ОООТЕХНОФЛЕКС; Тов.знак: ОТСУТСТВУЕТ;
                                    */
                                    try
                                    {
                                        dst.Cells[dst_row, 8].Value2 = get_value2(nonPars, "СОРТ");
                                        dst.Cells[dst_row, 9].Value = "ОТСУТСТВУЕТ";
                                        var ssize = get_size_string(nonPars).Split(separator_size);
                                        dst.Cells[dst_row, 10].Value2 = ssize[0].Replace(',', '.');
                                        dst.Cells[dst_row, 11].Value2 = ssize[1].Replace(',', '.');
                                        dst.Cells[dst_row, 12].Value2 = ssize[2].Replace(',', '.');
                                        string volume = get_weight(nonPars);
                                        if (volume == "ОТСУТСТВУЕТ")
                                            volume = src.Cells[src_row, GetColumnNumber(src, "G31_7")].Value2.Replace(',', '.');
                                        dst.Cells[dst_row, 13].Value2 = volume;
                                    }
                                    catch (Exception)
                                    {
                                        try
                                        {
                                            /*
                                             * 1-ФАНЕРА КЛЕЕНАЯ, СОСТОЯЩАЯ ИСКЛЮЧИТЕЛЬНО ИЗ БЕРЕЗОВОГО ШПОНА МАРКИ ФК, ГОСТ 3916.1-96, РАЗМЕР 1525Х1525ММ, 
                                             * НЕШЛИФОВАННАЯ,СОРТ 4/4, КЛАС ЭМИССИИ Е-1, КРОМКИ И ТОРЦЫ НЕ ИМЕЮТ ПАЗОВ И ГРЕБНЕЙ,: ТОЛЩИНА- 6ММ, 32 ПАКЕТА-32.15КУБ.М, 
                                             * КОЛИЧЕСТВО СЛОЕВ-5, ТО_[=1=] ЛЩИНА КАЖДОГО СЛОЯ 1,2ММ, СПЕЦИФИКАЦИЯ №56,ЦЕНА 6ММ- 230 ЕВРО ДЛЯ СТРОИТЕЛЬНЫХ РАБОТ:_[=1.1=]  
                                             * Изготовитель: ПК МАКСАТИХИНСКИЙ ЛЕСОПРОМЫШЛЕННЫЙ КОМБИНАТ; Тов.знак: ОТСУТСТВУЕТ;
                                            */
                                            dst.Cells[dst_row, 8].Value2 = get_value2(nonPars, "СОРТ");
                                            dst.Cells[dst_row, 9].Value = get_value2(nonPars, "МАРКИ");
                                            int r = nonPars.IndexOf("РАЗМЕР");
                                            string a = get_numeric(nonPars, r);
                                            dst.Cells[dst_row, 10].Value2 = a;
                                            r += a.Length;
                                            dst.Cells[dst_row, 11].Value2 = get_numeric(nonPars, r);
                                            dst.Cells[dst_row, 12].Value2 = get_numeric(nonPars, nonPars.IndexOf("ТОЛЩИНА"));
                                            string volume = get_numeric(nonPars, nonPars.IndexOf("ПАКЕТА"));
                                            if (volume == "ОТСУТСТВУЕТ")
                                                volume = src.Cells[src_row, GetColumnNumber(src, "G31_7")].Value2.Replace(',', '.');
                                            dst.Cells[dst_row, 13].Value2 = volume;
                                        }
                                        catch (Exception)
                                        {
                                            try
                                            {
                                                /*
                                                1 - ФАНЕРА КЛЕЕНАЯ, СОСТОЯЩАЯ ИСКЛЮЧИТЕЛЬНО ИЗ БЕРЕЗОВОГО ШПОНА МАРКИ ФК, ГОСТ 3916.1 - 96, РАЗМЕР 1525Х1525ММ, 
                                                НЕШЛИФОВАННАЯ,СОРТ 4 / 4, КЛАС ЭМИССИИ Е - 1, КРОМКИ И ТОРЦЫ НЕ ИМЕЮТ ПАЗОВ И ГРЕБНЕЙ,: ТОЛЩИНА - 9ММ, 
                                                17ПАКЕТОВ - 17,08КУБ.М, КОЛИЧЕСТВО СЛОЕВ-7, ТОЛЩИНА - 15 ММ, 16 ПАКЕТОВ - 15,94 КУБ.М, КОЛИЧЕСТВО СЛОЕВ-11, 
                                                ТОЛЩИНА КАЖДОГО СЛОЯ 1,33ММ, СПЕЦИФИКАЦИЯ № 10, ЦЕНА 9ММ - 214ЕВРО,ЦЕНА 15 ММ - 207ЕВРО, ДЛЯ СТРОИТЕЛЬНЫХ РАБОТ: 
                                                Изготовитель: ПК МАКСАТИХИНСКИЙ ЛЕСОПРОМЫШЛЕННЫЙ КОМБИНАТ; Тов.знак: ОТСУТСТВУЕТ;
                                                */
                                                if (nonPars.IndexOf("СОРТ") == -1 ||
                                                    nonPars.IndexOf("МАРКИ") == -1 ||
                                                    nonPars.IndexOf("РАЗМЕР") == -1 ||
                                                    nonPars.IndexOf("ТОЛЩИНА - ") == -1 ||
                                                    nonPars.IndexOf("ПАКЕТОВ") == -1)
                                                    throw new Exception("Не соответствие шаблону");
                                                dst.Cells[dst_row, 8].Value = get_value2(nonPars, "СОРТ");
                                                dst.Cells[dst_row, 9].Value = get_value2(nonPars, "МАРКИ");
                                                int r = nonPars.IndexOf("РАЗМЕР");
                                                string a = get_numeric(nonPars, r);
                                                dst.Cells[dst_row, 10].Value2 = a;
                                                r += a.Length;
                                                dst_row--;
                                                dst.Cells[dst_row, 11].Value2 = get_numeric(nonPars, r);
                                                bool bfirst = true;

                                                var temps = nonPars.Split(new string[] { "ТОЛЩИНА - " }, StringSplitOptions.RemoveEmptyEntries).ToList();
                                                temps.RemoveAt(0);
                                                foreach (var tmp in temps)
                                                {
                                                    dst_row++;

                                                    if (!bfirst)
                                                        dst.Cells[dst_row - 1, 1].EntireRow.Copy(dst.Cells[dst_row, 1].EntireRow);
                                                    dst.Cells[dst_row, 12].Value2 = get_numeric(tmp, 0);
                                                    string volume = get_numeric(tmp, nonPars.IndexOf("ПАКЕТОВ"));
                                                    if (volume == "ОТСУТСТВУЕТ")
                                                        volume = src.Cells[src_row, GetColumnNumber(src, "G31_7")].Value2.Replace(',', '.');
                                                    dst.Cells[dst_row, 13].Value2 = volume;
                                                    bfirst = false;
                                                }
                                            }
                                            catch (Exception)
                                            {
                                                //!!!!! НЕПРАВИЛЬНО
                                                /*
                                                 * ФАHЕРА КЛЕЕНАЯ БЕРЕЗОВАЯ (С ТОЛЩИНОЙ ШПОНА НЕ БОЛЕЕ 2,5ММ, НАРУЖН.СЛОИ ИЗ ЛИСТОВ БЕРЕЗ.ШПОНА) 
                                                 * ВВ 1250*2500
                                                 *  (
                                                 *      6,5ММ-2,966М3;
                                                 *      9ММ-6,076М3;
                                                 *      12ММ-15М3;
                                                 *      15ММ-27М3;
                                                 *      18ММ-12,152М3;
                                                 *      21ММ-13,581М3;
                                                 *      24ММ-3М3;
                                                 *      27ММ-3,038М3;
                                                 *      30ММ-6М3
                                                 *  ).: 
                                                 *  Изготовитель: ООО СЫКТЫВКАРСКИЙ ФАНЕРНЫЙ ЗАВОД; 
                                                 *  Тов.знак: SYPLY; 
                                                 *  Марка :ФСФ; 
                                                 *  Модель: ОТСУТСТВУЕТ; 
                                                 *  Артикул: ОТСУТСТВУЕТ; 
                                                 *  Стандарт: ТУ5512-001-44769167-11; 
                                                 *  Кол-во: 88,813 М3
                                                */
                                                dst.Cells[dst_row, 9].Value = get_value(nonPars, "Марка");
                                                int ttl = nonPars.IndexOf(')') + 1;
                                                int ttr = ttl;
                                                while (!char.IsDigit(nonPars[ttr]))
                                                    ttr++;
                                                dst.Cells[dst_row, 8].Value = nonPars.Substring(ttl, ttr - ttl - 1);

                                                string a = get_numeric(nonPars, ttr);
                                                dst.Cells[dst_row, 10].Value2 = a;
                                                ttr += a.Length;
                                                a = get_numeric(nonPars, ttr);
                                                dst.Cells[dst_row, 11].Value2 = a;

                                                var tmppars = nonPars.Substring(nonPars.IndexOf('(', ttl) + 1);
                                                tmppars = tmppars.Remove(tmppars.IndexOf(')'));
                                                var temps = tmppars.Split(new char[] { ';' }).ToList();
                                                dst_row--;
                                                bool bfirst = true;
                                                foreach (var tmp in temps)
                                                {
                                                    dst_row++;

                                                    if (!bfirst)
                                                        dst.Cells[dst_row - 1, 1].EntireRow.Copy(dst.Cells[dst_row, 1].EntireRow);

                                                    var b = get_numeric(tmp, 0);
                                                    dst.Cells[dst_row, 12].Value2 = b;
                                                    dst.Cells[dst_row, 13].Value2 = get_numeric(tmp, b.Length);
                                                    bfirst = false;
                                                }

                                            }
                                        }
                                    }
                                }
                            }
                            break;
                    }

                    dst_row++;
                    lproc++;
                }
                catch (ThreadAbortException)
                {
                    return;
                }
                catch (Exception ex)
                {
                    Logger.LogMessage(ex);
                    try
                    {
                        dst.Cells[dst_row, 1].EntireRow.Clear();
                        src.Cells[src_row, 1].EntireRow.Copy(err.Cells[err_row++, 1].EntireRow);
                    }
                    catch (Exception ex2)
                    {
                        Logger.LogMessage(ex2);
                    }
                }
                pg_bar.Invoke((MethodInvoker)delegate
                            {
                                pg_bar.Value = src_row - 1;
                            });
                l_proc.Invoke((MethodInvoker)delegate
                {
                    l_proc.Text = lproc.ToString();
                });
                l_nproc.Invoke((MethodInvoker)delegate
                {
                    l_nproc.Text = (srcMaxRows - 1 - lproc).ToString();
                });

                swatch.Stop();
                sumTicks += swatch.Elapsed.Ticks;
                if (src_row % 25 == 0)
                {

                    l_lost.Invoke((MethodInvoker)delegate
                    {
                        l_lost.Text = TimeSpan.FromTicks((sumTicks * (srcMaxRows - src_row) / 25)).ToString(@"hh\:mm\:ss");
                        sumTicks = 0;
                    });

                }
            }
            Logger.LogMessage("INFO", "Finish parsing");
            ex_serv.Workbooks[1].Close();
            Logger.LogMessage("INFO", "Closed first wokbook");
            /*if (l_nproc.Text == "0")
                ex_serv.Workbooks[3].Close();*/
            ex_serv.Visible = true;
            timer.Enabled = false;
            start.Invoke((MethodInvoker)delegate
            {
                start.Enabled = true;
                MessageBox.Show("Выполнено!", "Парсинг", MessageBoxButtons.OK, MessageBoxIcon.Information);
            });
        }

        private bool check_size_string(string str)
        {
            char[] seps = new char[] { 'x', 'X', 'х', 'Х', '*' };
            int i = 0;
            while (i < str.Length && str[i] == ' ') i++;
            while (i < str.Length && (str[i] != ' ' && !seps.Contains(str[i])))
            {
                if (!char.IsDigit(str[i]))
                    return false;
                i++;
            }
            while (i < str.Length && str[i] == ' ') i++;
            if (!seps.Contains(str[i]))
                return false;
            i++;
            while (i < str.Length && str[i] != ' ')
            {
                if (!char.IsDigit(str[i]))
                    return false;
                i++;
            }
            while (i < str.Length && str[i] == ' ') i++;
            if (i < str.Length)
                return false;
            return true;
        }

        private string get_value(string src, string valname)
        {
            int j = 0;
            for (int i = 0; i < src.Length; ++i)
            {
                if (char.ToLower(src[i]) == char.ToLower(valname[j]))
                    j++;
                else
                    j = 0;
                if (j == valname.Length)
                {
                    int k = i;
                    while (k < src.Length && src[k] != ':') k++;
                    int r = i;
                    while (r < src.Length && src[r] != ';') r++;
                    return src.Substring(k + 1, r - k - 1).Trim(' ');
                }
            }
            return "ОТСУТСТВУЕТ";
        }

        private string get_value2(string src, string valname)
        {
            int j = 0;
            for (int i = 0; i < src.Length; ++i)
            {
                if (char.ToLower(src[i]) == char.ToLower(valname[j]))
                    j++;
                else
                    j = 0;
                if (j == valname.Length)
                {
                    i++;
                    while (i < src.Length && src[i] == ' ') i++;
                    int k = i;
                    while (k < src.Length && !(src[k] == ',' || src[k] == ' ')) k++;
                    return src.Substring(i, k - i).Trim(' ');
                }
            }
            return "ОТСУТСТВУЕТ";
        }

        private string get_size_string(string src)
        {
            int st = 0;
            int fn = 0;
            int sep_count = 0;
            for (int i = 0; i < src.Length; ++i)
            {
                if (char.IsDigit(src[i]) || src[i] == ',' || src[i] == '.' || separator_size.Contains(src[i]))
                {
                    if (sep_count == 2)
                        break;
                    fn = i;
                    sep_count += separator_size.Contains(src[i]) ? 1 : 0;
                }
                else
                    st = i;
                if (sep_count == 2)
                    break;
            }
            if (sep_count == 2)
            {
                fn++;
                while (char.IsDigit(src[fn]) || src[fn] == ',' || src[fn] == '.') fn++;
                return src.Substring(st + 1, fn - st - 1);
            }
            return "ОТСУТСТВУЕТ";
        }

        private string get_weight(string src)
        {

            int r = src.ToLower().IndexOf("м3");
            if (r == -1)
                r = src.ToLower().IndexOf("m3");
            if (r == -1)
                return "ОТСУТСТВУЕТ";
            r--;

            while (r > 0 && src[r] == ' ') r--;

            int l = r;
            while (l > 0 && (char.IsDigit(src[l]) || src[l] == ',' || src[l] == '.')) l--;
            return get_numeric(src, l);
        }

        private string get_numeric(string src, int pos = 0)
        {
            int l = pos;
            while (l < src.Length && !char.IsDigit(src[l])) l++;
            int r = l;
            bool fl = true;
            while (r < src.Length && (char.IsDigit(src[r]) || fl && (src[r] == ',' || src[r] == '.')))
            {
                if (src[r] == ',' || src[r] == '.')
                    fl = false;
                r++;
            }

            return src.Substring(l, r - l).Trim(new char[] { ' ', ',', '.' }).Replace(',', '.');
        }

        private void start_Click(object sender, EventArgs e)
        {
            if (ex_serv != null)
                ex_serv.Quit();

            start.Enabled = false;
            l_cur_time.Text = l_lost.Text = "00:00:00";
            l_all.Text = l_proc.Text = l_nproc.Text = "0";
            pars_thread = new Thread(Parsing);

            sec = 0;
            timer.Enabled = true;

            Logger.LogMessage("INFO", "Open file " + Path.GetFullPath(fileName));
            ex_serv.Workbooks.Open(Path.GetFullPath(fileName));
            ex_serv.Workbooks.Add();
            ex_serv.Workbooks.Add();
            pars_thread.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            sec++;
            l_cur_time.Text = TimeSpan.FromSeconds(sec).ToString(@"hh\:mm\:ss");
        }

        private void b_file_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                start.Enabled = true;
                fileName = ofd.FileName;
                link_file.Text = Path.GetFileName(fileName);
            }
            else
            {
                start.Enabled = false;
                link_file.Text = "...";
            }

        }
    }
}
