using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using Janus.Windows.GridEX;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Easy1K
{
	public partial class MainForm : Form
    {
        // Тип для подсчета сменяемости
        public struct smen
        {
            public int key;
            public int sluzba_prev;
        }

        // Тип для подсчета выслуги
        public struct vysl
        {
            public string key;
            public int years_c;
            public int years_l;
        }

        public static int[] pos = new int[127];  // Массив расчетных позиций
        public static int[,] oneK = new int[1000, 127]; // Массив всего отчета
        
        // Список используемых ДГ до 2016 г.
        //public List<int> DG = new List<int> { 10, 13, 20, 30, 40, 45, 48, 50, 60, 63, 73, 83, 92, 117, 170, 171, 191, 193, 194, 200, 201, 202, 203, 204, 210, 211, 220, 245, 246, 251, 260, 270, 271, 280, 281, 283, 290, 291, 300, 301, 350, 351, 352, 355, 356, 360, 361, 362, 380, 381, 390, 810, 820, 860, 870, 900, 901, 971, 972, 978, 979, 992, 993, 994 };
        //public List<string> repDG = new List<string> { "010", "013", "020", "030", "971", "972", "978", "979", "092", "117", "170", "171", "210", "211", "200", "202", "191", "073", "083", "290", "283", "270", "281", "280", "271", "260", "220", "350", "380", "381", "390", "291", "245", "246", "351", "352", "251", "193", "194", "993", "994", "992", "355", "356", "810", "820", "860", "870", "360", "361", "362", "040", "048", "045", "050", "060", "063", "203", "201", "204" };
        
        // Список используемых ДГ после 2016 г. + ПКОН, ПВМ, -ОВО, ОМОН, СОБР
        public List<int> DG = new List<int> { 10, 13, 20, 30, 40, 45, 48, 50, 60, 63, 73, 83, 92, 117, 170, 171, 191, 193, 194, 200, 201, 202, 203, 204, 210, 211, 220, 245, 246, 260, 270, 271, 280, 281, 283, 290, 291, 300, 301, 350, 351, 352, 355, 356, 360, 361, 362, 410, 420, 440, 450, 900, 901, 971, 972, 978, 979, 992, 993, 994 };
        public List<string> repDG = new List<string> { "010", "013", "020", "030", "971", "972", "978", "979", "092", "117", "170", "171", "210", "211", "200", "202", "191", "073", "083", "290", "283", "270", "281", "280", "271", "260", "220", "350", "291", "245", "246", "351", "352", "193", "194", "993", "994", "992", "355", "356", "360", "361", "362", "040", "048", "045", "050", "060", "063", "203", "201", "204", "410", "420", "440", "450" };
                
        // Количество Должностных групп (0..n)
        public int CountDG = 55;
        // Массивы ключей для сменяемости и выслуги
        public List<int> PoslKeys1 = new List<int> { };
        public List<smen> PoslKeys2 = new List<smen> { };
        public List<vysl> Vysluga = new List<vysl> { };

        // Текущие показатели из таблицы OneK_Pergroup
        public static string current_dg_code;   // Текущая должностная группа   KEY_OTCH
        public static string current_sql_text;  // Текущий текст запроса ДГ     TEXT_QRY
        public static string current_dg_name;   // Наименование текущей ДГ      NAME_FORM
        public static bool current_dg_lider;    // Текущая ДГ руководящая (Д/Н) LIDER
        public static string current_dg_sluzb;     // Текущий показатель службы    SLUZBA_SQL
        public static string current_dg_dolz;   // Текущий показатель должности DOLZNOST_SQL

        public DateTime calc_date = Convert.ToDateTime("31.12.2022");//DateTime.Now.ToShortDateString());
        public string calc_year = "2022";//calc_date.Year.ToString();          // Год расчета
        public string calc_month = "12";//calc_date.Month.ToString();          // Месяц расчета

        public static string KConn = "";
        public static string RConn = "";

        public MainForm()
        {
            InitializeComponent();
        }

        // Инициализация расчетного массива
        public void InitPosArray(int maxinit)
        {
            for (int i = 0; i < maxinit; i++) pos[i] = 0;
        }

        // Инициализация всего отчета 
        public void InitAllOneK()
        {
            for (int i = 0; i < 1000; i++)
            {
                for (int j = 0; j < 127; j++) oneK[i, j] = 0;
            }
        }

        // Сохранение PosArray в OneK
        public void SavePos2OneK(string code)
        {
            int n = Convert.ToInt16(code);

            for (int i = 0; i < 127; i++) oneK[n, i] = pos[i]; 
        }

        // Сохранение всего отчета в БД
        public void SaveAllOneKToBase()
        {
            for (int i = 0; i < 1000; i++)
            {
                if (DG.Contains(i))
                {
                    if ( i < 100 ) SaveDGtoBase("0" + i.ToString(), calc_date.ToShortDateString());
                    else SaveDGtoBase(i.ToString(), calc_date.ToShortDateString());
                }
            }
            MessageBox.Show("Отчет сохранен!");
        }

        // Загрузка данных отчета из БД по дате
        public void LoadAllOneKFromBase(DateTime date)
        {
            InitAllOneK();

            DataTable dt = DataProvider._getDataSQL(RConn, String.Format("SELECT * FROM OneK_2012 WHERE calc_date = '{0}' ORDER BY code_dg", DateToStr(date)));
            DataRowCollection rc = dt.Rows;
            dt.Dispose();

            if (rc.Count > 0)
            {

                for (int i = 0; i < rc.Count; i++)
                {
                    int code = Convert.ToInt16(rc[i]["code_dg"]);
                    for (int j = 1; j < 127; j++) oneK[code, j] = Convert.ToInt16(rc[i]["pos_" + j.ToString()]);
                }

                rc.Clear();
                MessageBox.Show("Отчет загружен!");
            }
            else MessageBox.Show("Ошибка загрузки данных!");
        }

        // Проверка между должностными группами
        public bool CheckInterDG()
        {
            int cnt = 0;
            int [] sum1 = new int[127];
            int [] sum2 = new int[127];

            debug.Clear();
            
            for (int i = 1; i < 127; i++)
            {
                //020+030=971+972+978+979+360+362
                sum1[i] = oneK[20,i] + oneK[30,i];
                sum2[i] = oneK[971,i] + oneK[972,i] + oneK[978,i] + oneK[979,i] + oneK[360,i] + oneK[362,i];
                if (sum1[i] != sum2[i])
                {
                    cnt++;
                    debug.Text += String.Format("Не бъется сумма: [020]+[030] = [971]+[972]+[978]+[979]+[360]+[362] в позиции {0} | {1} <> {2}\n" +
                                                "---------------------------------------------------------------------------------------------\n", i, sum1[i], sum2[i]);
                    debug.Text += String.Format("Не бъется сумма: ({0})+({1}) = ({2})+({3})+({4})+({5})+({6})+({7}) в позиции {8} | {9} <> {10}\n", oneK[20,i],oneK[30,i], oneK[971,i], oneK[972,i], oneK[978,i], oneK[979,i], oneK[360,i], oneK[362,i], i, sum1[i], sum2[i]);
                }
                
                // 020>=971+978+360
                sum1[i] = oneK[971, i] + oneK[978, i] + oneK[360, i];
                if ( oneK[20,i] < sum1[i] )
                {
                    cnt++;
                    debug.Text += String.Format("Не выполнено условие: [020] >= [971]+[978]+[360] в позиции {0} | {1} < {2}\n", i, oneK[20,i], sum1[i]);
                }
                
                // 030>=972+979+362
                sum1[i] = oneK[972,i] + oneK[979,i] + oneK[362,i];
                if ( oneK[30,i] < sum1[i] )
                {
                    cnt++;
                    debug.Text += String.Format("Не выполнено условие: [030] >= [972]+[979]+[362] в позиции {0} | {1} < {2}\n", i, oneK[30,i], sum1[i]);
                }

                //360+362<=020+030
                sum1[i] = oneK[360, i] + oneK[362, i];
                sum2[i] = oneK[20, i] + oneK[30, i];
                if ( sum1[i] > sum2[i] )
                {
                    cnt++;
                    debug.Text += String.Format("Не выполнено условие: [360]+[362] <= [20]+[30] в позиции {0} | {1} > {2}\n", i, sum1[i], sum2[i]);
                }

                // 048<=040
                if ( oneK[48,i] > oneK[040,i] ) { cnt++; debug.Text += String.Format("Не выполнено условие: [048] <= [040] в позиции {0} | {1} > {2}\n", i, oneK[48,i], oneK[40,i]); }

                // 171<=170
                if (oneK[171, i] > oneK[170, i]) { cnt++; debug.Text += String.Format("Не выполнено условие: [171] <= [170] в позиции {0}\n", i); }

                // 083+280+246+352+194+993+356+420+450<=972
                sum1[i] = oneK[83, i] + oneK[280, i] + oneK[246, i] + oneK[352, i] + oneK[194, i] + oneK[993,i] + oneK[356,i] + oneK[420,i] + oneK[450,i];
                if (sum1[i] > oneK[972,i])
                {
                    cnt++;
                    debug.Text += String.Format("Не выполняется условие: 083+280+246+352+194+993+356+420+450<=972 в позиции {0} | {1} > {2}\n", i, sum1[i], oneK[972,i]);
                }

                // 211 <= 210
                if (oneK[211, i] > oneK[210, i]) { cnt++; debug.Text += String.Format("Не выполнено условие: [211] <= [210] в позиции {0} | {1} > {2}\n", i, oneK[211,i], oneK[210,i]); }
                // 202<=200
                if (oneK[202, i] > oneK[200, i]) { cnt++; debug.Text += String.Format("Не выполнено условие: [202] <= [200] в позиции {0} | {1} > {2}\n", i, oneK[202, i], oneK[200, i]); }
                // 361<=360
                if (oneK[361, i] > oneK[360, i]) { cnt++; debug.Text += String.Format("Не выполнено условие: [361] <= [360] в позиции {0} | {1} > {2}\n", i, oneK[361, i], oneK[360, i]); }
                // 283<=290
                if (oneK[283, i] > oneK[290, i]) { cnt++; debug.Text += String.Format("Не выполнено условие: [283] <= [290] в позиции {0} | {1} > {2}\n", i, oneK[283, i], oneK[290, i]); }
                // 381<=380 исключено в 2016 г.
                //if (oneK[381, i] > oneK[380, i]) { cnt++; debug.Text += String.Format("Не выполнено условие: [381] <= [380] в позиции {0} | {1} > {2}", i, oneK[381, i], oneK[380, i])); }


                // 170+191+210+200+073+290+220+350+291+245+351+193+994+355+410+440<=971

                sum1[i] = oneK[170, i] + oneK[191, i] + oneK[210, i] + oneK[200, i] + oneK[073, i] + oneK[290, i] + oneK[220, i] + oneK[350, i] + oneK[291, i] + oneK[245,i] + oneK[351,i] + oneK[193,i] +  oneK[994,i] +  oneK[355,i] + oneK[410,i] + oneK[440,i];
                if (sum1[i] > oneK[971, i])
                {
                    cnt++;
                    debug.Text += String.Format("Не выполняется условие: 170+191+210+200+073+290+220+350+291+245+351+193+994+355+410+440<=971 в позиции {0} | {1} > {2}\n", i, sum1[i], oneK[971, i]);
                }
                
            }
            
            if (cnt > 0) return false;
            else return true;
        }

        // Проверка должностной группы
        public bool CheckDG(string code)
        {
            int cnt = 0; // счетчик ошибок

            debug.Clear();

            if (pos[1] <= 0) { debug.Text += String.Format("Позиция [01] не может быть меньше или равна нулю.([01] = {0})\n",pos[1]); cnt++; }
            if ((pos[126] > 0) && (pos[2] / pos[126]) < 0) { debug.Text += "Не выполнено условие: [02] / [126] >= 0\n"; cnt++; }
            if (pos[2] > pos[1]) { debug.Text += "Замещение [02] не может быть больше штатной численности [01]\n"; cnt++; }
            if (pos[3] > pos[2]) { debug.Text += "Позиция [03] должна быть меньше или равна [02]\n"; cnt++; }
            if ((pos[4] + pos[5]) > pos[2]) { debug.Text += "Не выполнено условие 04 + 05 <= 02\n"; cnt++; }
            if ((pos[6] + pos[14] + pos[25]) != pos[2]) { debug.Text += "Не выполнено условие 06 + 14 + 25 = 02\n"; cnt++; }
            if ((pos[7] + pos[8] + pos[9] + pos[10]) > pos[6]) { debug.Text += "Не выполнено условие 07+08+09+10 <= 06\n"; cnt++; }
            if ((pos[11] > pos[6]) && code != "025" && code != "035" && code != "010" && code != "013") { debug.Text += "Не выполнено условие 11 <= 06\n"; cnt++; }
            if ((pos[12] > pos[11]) && code != "025" && code != "035" && code != "010" && code != "013") { debug.Text += "Не выполнено условие 12 <= 11\n"; cnt++; }
            if ((pos[13] > pos[6]) && code != "025" && code != "035" && code != "010" && code != "013") { debug.Text += "Не выполнено условие 13 <= 06\n"; cnt++; }
            if ((pos[15] + pos[16] + pos[17] + pos[18]) > pos[14]) { debug.Text += "Не выполнено условие 15+16+17+18 <= 14\n"; cnt++; }
            if ((pos[19] > pos[14]) && code != "025" && code != "035" && code != "010" && code != "013") { debug.Text += "Не выполнено условие 19 <= 14\n"; cnt++; }
            if ((pos[20] > pos[19]) && code != "025" && code != "035" && code != "010" && code != "013") { debug.Text += "Не выполнено условие 20 <= 19\n"; cnt++; }
            if ((pos[21] > pos[14]) && code != "025" && code != "035" && code != "010" && code != "013") { debug.Text += "Не выполнено условие 21 <= 14\n"; cnt++; }
            if ((pos[7] + pos[15]) <= 0) { debug.Text += "Возможная ошибка: 07+15> 0*\n"; cnt++; }
            if (((pos[23] + pos[24]) != pos[22]) && code != "025" && code != "035") { debug.Text += "Не выполнено условие 23+24 = 22\n"; cnt++; }
            if (pos[22] > pos[2]) { debug.Text += "Позиция [22] должна быть меньше или равна [02]\n"; cnt++; }
            if ((pos[26] + pos[27] + pos[28] + pos[29] + pos[30]) != pos[2]) { debug.Text += "Не выполнено условие 02 = 26+27+28+29+30\n"; cnt++; }
            if (((pos[31] + pos[32] + pos[33] + pos[34] + pos[35] + pos[36] + pos[37]) != pos[2]) && code != "025" && code != "035" && code != "010" && code != "013") { debug.Text += "Не выполнено условие 02=31+32+33+34+35+36+37\n"; cnt++; }
            if (((pos[38] + pos[39] + pos[40] + pos[41] + pos[42] + pos[43]) != pos[2]) && code != "010" && code != "013" && code != "020" && code != "030" && code != "063" && code != "083" && code != "194" && code != "246" && code != "280" && code != "352" && code != "356" && code != "362" && code != "390" && code != "790" && code != "820" && code != "870" && code != "971" && code != "972" && code != "978" && code != "979" && code != "993" && code != "025" && code != "035") { debug.Text += "Не выполнено условие 02=38+39+40+41++42+43\n"; cnt++; }
            if ((pos[44] > pos[2]) && code != "025" && code != "035") { debug.Text += "Позиция [44] должна быть меньше или равна [02]\n"; cnt++; }
            if (((pos[36] + pos[37]) > pos[44]) && code != "025" && code != "035" && code != "010" && code != "013") { debug.Text += "Не выполнено условие 36+37 <= 44\n"; cnt++; }
            if (pos[46] > pos[2]) { debug.Text += "Возможная ошибка: 46 <= 02\n"; cnt++; }
            if (pos[45] < pos[46]) { debug.Text += "*Возможная ошибка: 46 <= 45\n"; cnt++; }
            if (((pos[47] + pos[48] + pos[49] + pos[50] + pos[51]) > pos[46]) && code != "025" && code != "035") { debug.Text += "Не выполнено условие 47+48+49+50+51 <= 46\n"; cnt++; }
            if (((pos[52] + pos[53] + pos[54]) != pos[46]) && code != "025" && code != "035") { debug.Text += "Не выполнено условие 52+53+54 = 46\n"; cnt++; }
            if ((pos[55] > pos[2]) && code != "025" && code != "035" && code != "010" && code != "013") { debug.Text += "Не выполнено условие 55 <= 02\n"; cnt++; }
            if ((pos[56] > pos[55]) && code != "025" && code != "035" && code != "010" && code != "013") { debug.Text += "Не выполнено условие 56 <= 55\n"; cnt++; }
            if ((pos[61] > pos[2]) && code != "025" && code != "035" && code != "010" && code != "013") { debug.Text += "Не выполнено условие 61 <= 02\n"; cnt++; }
            if (pos[62] > pos[2]) { debug.Text += "Не выполнено условие 62 <= 02\n"; cnt++; }
            if ((pos[63] > pos[62]) && code != "025" && code != "035" && code != "010" && code != "013") { debug.Text += "Не выполнено условие 63 <= 62\n"; cnt++; }
            if (((pos[63] + pos[64] + pos[65] + pos[66] + pos[67] + pos[68] + pos[69]) != pos[62]) && code != "025" && code != "035" && code != "010" && code != "013") { debug.Text += "Не выполнено условие 63+64+65+66+67+68+69 = 62\n"; cnt++; }
            if ((pos[70] > pos[62]) && code != "025" && code != "035" && code != "010" && code != "013") { debug.Text += "Не выполнено условие 70 <= 62\n"; cnt++; }
            if ((pos[72] > pos[71]) && code != "025" && code != "035" && code != "010" && code != "013") { debug.Text += "Не выполнено условие 72 <= 71\n"; cnt++; }
            if (((pos[73] + pos[74]) > pos[71]) && code != "025" && code != "035" && code != "010" && code != "013") { debug.Text += "Не выполнено условие 71=73+74\n"; cnt++; }
            if ((pos[75] > pos[71]) && code != "025" && code != "035" && code != "010" && code != "013") { debug.Text += "Не выполнено условие 75 <= 71\n"; cnt++; }

            int sum = pos[63];
            for (int i = 76; i < 117; i++) sum += pos[i];
            if (sum != pos[62] && code != "010" && code != "013") { debug.Text += "Не выполнено условие 62=63+76+77+78+……+116\n"; cnt++; }

            if (current_dg_lider == true)
            {
                if ((pos[120] + pos[122] + pos[123] + pos[124]) > pos[118]) { debug.Text += "Не выполнено условие 120+122+123+124 <= 118\n"; cnt++; }
                if (pos[121] > pos[120]) { debug.Text += "Не выполнено условие 121<=120\n"; cnt++; }
                if ((pos[125] + pos[126]) > pos[124]) { debug.Text += "Не выполнено условие 125+126<=124\n"; cnt++; }
            }
 
            if (cnt > 0) return false;
            else return true;
        }

        // Создание строки КЛЮЧЕЙ для запроса (из DataRowCollection)
        public string GenKeys(DataRowCollection rc)
        {
            string keys = "";
            for (int n = 0; n < rc.Count; n++)
            {
              if (n == 0) keys += rc[n]["KEY_1"].ToString();
              else keys += "," + rc[n]["KEY_1"].ToString();
            }
            return keys;
        }

        // Создание строки КЛЮЧЕЙ для запроса (из DataRowCollection)
        public string GenKeysP(DataRowCollection rc)
        {
            string keys = "";
            for (int n = 0; n < rc.Count; n++)
            {
                if (n == 0) keys += rc[n]["key_posl"].ToString();
                else keys += "," + rc[n]["key_posl"].ToString();
            }
            return keys;
        }

        // Проверка на достижение предельного возраста
        public bool IsHiAge(DateTime born, int zv)
        {
            int let = calc_date.Year - born.Year;
            if (zv == 114 && let >= 60) return true;  // генералы до 60
            if ((zv == 53 || zv == 73 || zv == 113) && let >= 55) return true; // полковники до 55
            else if (let >= 50) return true; // остальнын до 50
            else return false;
        }

        // Предварительный расчет выслуги лет (включая льготную)
        public TPeriod CalcStage(string key, bool calendar)
        {
            TPeriod army = new TPeriod();
            TPeriod posl = new TPeriod();
            //TPeriod uch = new TPeriod();
            TPeriod skr_k = new TPeriod();
            TPeriod skr_l = new TPeriod();
            TPeriod itog = new TPeriod();
            DateDifference dd;
            
            DateTime dateX = calc_date;

            #region Армия...
            DataTable dt = DataProvider._getDataSQL(KConn, String.Format("SELECT ARMY_BEGIN, ARMY_STOP FROM VUCHET WHERE KEY_1 = {0}", key));
            DataRowCollection rc = dt.Rows;

            for (int i = 0; i < rc.Count; i++)
            {
                if (rc[i]["ARMY_BEGIN"] != DBNull.Value)
                {
                    dd = new DateDifference(Convert.ToDateTime(rc[i]["ARMY_BEGIN"]), Convert.ToDateTime(rc[i]["ARMY_STOP"]));
                    army.Add(dd.Days, dd.Months, dd.Years);
                }
            }

            dt.Dispose();
            rc.Clear();
            #endregion

            #region Служба по послужному...
            dt = DataProvider._getDataSQL(KConn, String.Format("SELECT DATA_OT, STATUS, NOM_PRIK, SLUZBA FROM POSL_SPI POSL WHERE (DOLZNOST < '800000') AND (KEY_POSL = {0}) ORDER BY DATA_OT, NOM_PRIK", key));
            rc = dt.Rows;

            DateTime date1;
            DateTime date2 = new DateTime();

            // Если в послужном 1 запись о приеме...
            if (rc.Count == 1)
            {
                // Если 1 запись в послужном (принят)
                if (rc[0]["STATUS"].ToString() == "2000")
                {
                    date1 = Convert.ToDateTime(rc[0]["DATA_OT"]);
                    date2 = dateX;
                    /// Вычисляем длительность периода
                    /// проверяем служба льготная ???
                    int sl = Convert.ToInt16(rc[0]["SLUZBA"]);
                    if (sl == 26 || sl == 28 || sl == 44 || sl == 59 || sl == 61) { if (calendar == false) dd = new DateDifference(date1, date2, 1.33); else dd = new DateDifference(date1, date2); }
                    else if (sl == 6 || sl == 39 || sl == 43 || sl == 58 || sl == 72) { if (calendar == false) dd = new DateDifference(date1, date2, 1.50); else dd = new DateDifference(date1, date2); }
                    else dd = new DateDifference(date1, date2);
                    // Добавляем в выслугу...
                    posl.Add(dd.Days, dd.Months, dd.Years);
                }
            }
            else // Если записей в послужном более 1
            {
                // Перебираем записи...
                // Начальные условия
                date1 = Convert.ToDateTime(rc[0]["DATA_OT"]); // дата
                int sl = Convert.ToInt16(rc[0]["SLUZBA"]); // служба
                int pos = 1;
                string st = ""; // статус
                do
                {
                    st = rc[pos]["STATUS"].ToString();
                    if (st == "4000") // Если уволен...
                    {
                        date2 = Convert.ToDateTime(rc[pos]["DATA_OT"]);
                        // Вычисляем длительность периода
                        if (sl == 26 || sl == 28 || sl == 44 || sl == 59 || sl == 61) { if (calendar == false) dd = new DateDifference(date1, date2, 1.33); else dd = new DateDifference(date1, date2); }
                        else if (sl == 6 || sl == 39 || sl == 43 || sl == 58 || sl == 72) { if (calendar == false) dd = new DateDifference(date1, date2, 1.50); else dd = new DateDifference(date1, date2); }
                        else dd = new DateDifference(date1, date2);
                        // Добавляем в выслугу...
                        posl.Add(dd.Days, dd.Months, dd.Years);
                        pos++;
                        if (pos != rc.Count) date1 = Convert.ToDateTime(rc[pos]["DATA_OT"]);
                    }
                    else
                    {
                        pos++;
                    }
                } while (pos < rc.Count);

                if (sl == 26 || sl == 28 || sl == 44 || sl == 59 || sl == 61) { if (calendar == false) dd = new DateDifference(date1, dateX, 1.33); else dd = new DateDifference(date1, dateX); }
                else if (sl == 6 || sl == 39 || sl == 43 || sl == 58 || sl == 72) { if (calendar == false) dd = new DateDifference(date1, dateX, 1.50); else dd = new DateDifference(date1, dateX); }
                else dd = new DateDifference(date1, dateX);
                
                // Добавляем в выслугу...
                posl.Add(dd.Days, dd.Months, dd.Years);

            }

            Application.DoEvents();
            dt.Dispose();
            rc.Clear();
            #endregion

            #region Служебные командировки...
            if (calendar != true)
            {
                dt = DataProvider._getDataSQL(KConn, String.Format("SELECT date_from, date_to, koef_visl from lgottime where koef_visl <> '1' and main_key = {0} order by date_from", key));
                rc = dt.Rows;

                if (rc.Count > 0)
                {
                    for (int i = 0; i < rc.Count; i++)
                    {
                        dd = new DateDifference(Convert.ToDateTime(rc[i]["date_from"]), Convert.ToDateTime(rc[i]["date_to"]));
                        skr_k.Add(dd.Days, dd.Months, dd.Years);
                        dd = new DateDifference(Convert.ToDateTime(rc[i]["date_from"]), Convert.ToDateTime(rc[i]["date_to"]), Convert.ToDouble(rc[i]["koef_visl"]));
                        skr_l.Add(dd.Days, dd.Months, dd.Years);
                    }

                    // Заменяем календарь командировок на льготку и добавляем к послужному
                    posl.Substract(skr_k.d, skr_k.m, skr_k.y);
                    posl.Add(skr_l.d, skr_l.m, skr_l.y);
                }
                rc.Clear();
                dt.Dispose();
            }
            #endregion

            #region Половина срока обучения...

            //// Дата поступления в ОВД
            //dt = DataProvider._getDataSQL(KConn, String.Format("SELECT DATA_POST FROM AAQQ WHERE KEY_1 = {0}", key));
            //rc = dt.Rows;
            //DateTime data_post = Convert.ToDateTime("01.01.1900");
            //if ( rc[0]["DATA_POST"] != DBNull.Value) data_post = Convert.ToDateTime(rc[0]["DATA_POST"].ToString());
            //dt.Dispose();
            //rc.Clear();
            //// Ищем вышку до даты поступления..
            //dt = DataProvider._getDataSQL(KConn, String.Format("SELECT VID, YEAR FROM LEARN WHERE VID='10' AND STATUS = 1 AND YEAR <= {0} AND KEY_1 = {1}", data_post.Year, key));
            //rc = dt.Rows;
            //if (rc.Count > 0) uch.Add(0, 4, 2);  // высшее +2г4м
            //else
            //{
            //    // Если нет, ищем среднее специальное
            //    dt.Dispose();
            //    rc.Clear();
            //    dt = DataProvider._getDataSQL(KConn, String.Format("SELECT VID, YEAR FROM LEARN WHERE VID='20' AND STATUS = 1 AND YEAR <= {0} AND KEY_1 = {1}", data_post.Year, key));
            //    rc = dt.Rows;
            //    if (rc.Count > 0) uch.Add(0, 6, 1);  // среднее профессиональное +1г6м
            //}
            //rc.Clear();
            //dt.Dispose();

            //itog.Add(uch.d, uch.m, uch.y);
            #endregion
            
            // Итого
            
            itog.Add(army.d, army.m, army.y);
            itog.Add(posl.d, posl.m, posl.y);

           return itog;
        }

        // Расчет должностной группы ( код, текс запроса )
        public void CalcDG(string code, string sql_text)
        {
            InitPosArray(127);

            progress.Maximum = 126;
            progress.Value = 0;
            string form_corrector = "";
            string sovmesh = "";

            #region // Должностей по штату...

            if (code != "010") form_corrector = "COUNT(DOLZNOST)";
            else form_corrector = "SUM(STAVKA_DLZ)";

            PrStat("Должностей по штату");
            if (code != "201") pos[1] = DataProvider._getDataSQLs(KConn, String.Format("SELECT {0} FROM AAQQ {1} AND DATA_SOKR IS NULL", form_corrector, sql_text));
            else pos[1] = Convert.ToInt16(UUPSelo_Stat.Text); // Если сельский участковый - пока ручками

            // из них замещено
            if (code != "010") form_corrector = "COUNT(KEY_1)";
            else form_corrector = "SUM(STAVKA_PRS)";

            // Если аттестованные считаем по ключам, вольных по ставкам
            PrStat("из них замещено");
            pos[2] = DataProvider._getDataSQLs(KConn, String.Format("SELECT {0} FROM AAQQ {1} AND FAMILIYA <> ''", form_corrector, sql_text));

            // в том числе женщинами
            PrStat("в том числе женщинами");
            pos[3] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' AND NACIONALN > 50", sql_text));

            // в том числе кандидатами наук
            PrStat("в том числе кандидатами наук");
            pos[4] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' AND UCH_STEP = 1", sql_text));

            // в том числе докторами наук
            PrStat("в том числе докторами наук");
            pos[5] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' AND UCH_STEP = 2", sql_text));

            #endregion


            #region // Сотрудников с высшим профессиональным образованием...

            // Корректировка по 010 форме
            if (code == "010") sovmesh = "AND LICH_NOM_2 <> 'совмещ'";
            else sovmesh = "";

            PrStat("Сотрудников с высшим профессиональным образованием");
            DataTable dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('10') ORDER BY KEY_1", sql_text, sovmesh));
            pos[6] = dt.Rows.Count;

            // в том числе юридическим
            PrStat("в том числе юридическим");
            pos[7] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('10') AND OBRAZ_LIC1 LIKE 'Юридическое%'", sql_text, sovmesh));

            // в том числе техническим
            PrStat("в том числе техническим");
            pos[8] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('10') AND OBRAZ_LIC1 LIKE 'Техническое%'", sql_text, sovmesh));

            // в том числе педагогическим
            PrStat("в том числе педагогическим");
            pos[9] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('10') AND OBRAZ_LIC1 LIKE 'Педагогическое%'", sql_text, sovmesh));

            // в том числе экономическим
            PrStat("в том числе экономическим");
            pos[10] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('10') AND OBRAZ_LIC1 LIKE 'Экономическое%'", sql_text, sovmesh));

            if (code != "010" && code != "013")
            {
                //в том числе выпускников высших ОУ МВД России
                if (dt.Rows.Count > 0)
                {
                    PrStat("в том числе выпускников высших ОУ МВД России");
                    DataTable dt1 = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM LEARN WHERE KEY_1 IN ({0}) AND STATUS = 1 AND VID IN ('10') AND UCH_ZAV IN (SELECT P2 FROM SLVUCZ WHERE P1 LIKE '%МВД%' AND P2 < '300000') ORDER BY KEY_1", GenKeys(dt.Rows)));
                    pos[11] = dt1.Rows.Count;

                    if (pos[11] > 0)
                    {
                        //из них очной формы обучения
                        PrStat("из них очной формы обучения");
                        pos[12] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM PRIEM WHERE KEY_1 IN ({0}) AND KAT_POST IN (103)", GenKeys(dt1.Rows)));
                    }
                    dt1.Dispose();
                }

                // в том числе образование не соответствует направлению деятельности
                PrStat("в том числе образование не соответствует направлению деятельности");
                if (dt.Rows.Count > 0) pos[13] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ WHERE KEY_1 IN ({0}) AND KVAL = 0", GenKeys(dt.Rows)));
                else pos[13] = 0;
            }
            dt.Dispose();
            #endregion


            #region // Сотрудников со средним профессиональным образованием...
            PrStat("Сотрудников со средним профессиональным образованием");
            dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('20','50') ORDER BY KEY_1", sql_text, sovmesh));
            pos[14] = dt.Rows.Count;

            // в том числе юридическим
            PrStat("в том числе юридическим");
            pos[15] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('20','50') AND OBRAZ_LIC1 LIKE 'Юридическое%'", sql_text, sovmesh));

            // в том числе техническим
            PrStat("в том числе техническим");
            pos[16] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('20','50') AND OBRAZ_LIC1 LIKE 'Техническое%'", sql_text, sovmesh));

            // в том числе педагогическим
            PrStat("в том числе педагогическим");
            pos[17] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('20','50') AND OBRAZ_LIC1 LIKE 'Педагогическое%'", sql_text, sovmesh));

            // в том числе экономическим
            PrStat("в том числе экономическим");
            pos[18] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('20','50') AND OBRAZ_LIC1 LIKE 'Экономическое%'", sql_text, sovmesh));

            if (code != "010" && code != "013")
            {
                //в том числе выпускников средних ОУ МВД России
                if (dt.Rows.Count > 0)
                {
                    PrStat("в том числе выпускников средних ОУ МВД России");
                    DataTable dt1 = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM LEARN WHERE KEY_1 IN ({0}) AND STATUS = 1 AND VID = '20' AND UCH_ZAV IN (SELECT P2 FROM SLVUCZ WHERE P1 LIKE '%МВД%' AND P2 BETWEEN '300000' AND '500000')", GenKeys(dt.Rows)));
                    pos[19] = dt1.Rows.Count;

                    if (pos[19] > 0)
                    {
                        //из них очной формы обучения
                        PrStat("из них очной формы обучения");
                        pos[20] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM PRIEM WHERE KEY_1 IN ({0}) AND KAT_POST IN (103)", GenKeys(dt1.Rows)));
                    }
                    dt1.Dispose();
                }

                // в том числе образование не соответствует направлению деятельности
                PrStat("в том числе образование не соответствует направлению деятельности");
                if (dt.Rows.Count > 0) pos[21] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ WHERE KEY_1 IN ({0}) AND KVAL = 0", GenKeys(dt.Rows)));
                else pos[21] = 0;
            }
            dt.Dispose();
            #endregion


            #region //Сотрудников, продолжающих обучение...
            PrStat("Сотрудников, продолжающих обучение");
            pos[22] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(FAMILIYA) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND UCHEBA = 1", sql_text, sovmesh));

            //в том числе	в ОУ МВД России
            PrStat("в том числе	в ОУ МВД России");
            pos[23] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(FAMILIYA) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND UCHEBA = 1 AND ((GDE_UCH BETWEEN '200000' AND '300000') OR (GDE_UCH BETWEEN '400000' AND '500000'))", sql_text, sovmesh));

            // в гражданских ОУ
            PrStat("в гражданских ОУ");
            pos[24] = pos[22] - pos[23];
            #endregion


            #region //Сотрудников со средним (полным) общим образованием
            PrStat("Сотрудников со средним (полным) общим образованием");
            pos[25] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('30','40')", sql_text, sovmesh));
            #endregion


            #region // Распределение сотрудников по возрасту..
            PrStat("Распределение сотрудников по возрасту");
            dt = DataProvider._getDataSQL(KConn, String.Format("SELECT DATA_ROZD FROM AAQQ {0} AND FAMILIYA <> '' {1}", sql_text, sovmesh));

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DateTime dateX = Convert.ToDateTime(dt.Rows[i]["DATA_ROZD"]);
                int y = calc_date.Year - dateX.Year;

                if (y <= 20) pos[26]++;
                else if (y <= 30 && y >= 21) pos[27]++;
                else if (y <= 40 && y >= 31) pos[28]++;
                else if (y <= 55 && y >= 41) pos[29]++;
                else if (y > 55)
                {
                    pos[30]++;
                    if (code == "010" || code == "013") pos[44]++;
                }
            }

            progress.Value += 4;
            dt.Dispose();
            #endregion
            

            #region // Распределение сотрудников по стажу службы в ОВД (выслуга лет) + Сотрудников, имеющих право выхода на пенсию
            if (code != "010" && code != "013")
            {
                PrStat("Распределение сотрудников по стажу службы в ОВД (выслуга лет) + Сотрудников, имеющих право выхода на пенсию");
                dt = DataProvider._getDataSQL(KConn, String.Format("SELECT DATA_POST, SL_RANE_OT, SL_RANE_DO, SLVARM_OT, SLVARM_DO, DATA_ROZD, ZVANIE, KEY_1 FROM AAQQ {0} AND FAMILIYA <> ''", sql_text));

                StatusBar.Panels[1].ProgressBarMaxValue = dt.Rows.Count;
                StatusBar.Panels[1].ProgressBarValue = 0;

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    StatusBar.Panels[1].ProgressBarValue++;
                    string id2 = dt.Rows[i]["KEY_1"].ToString();
                    vysl item = Vysluga.Find(x => x.key == id2);
                    
                    //DateTime date_post = Convert.ToDateTime(dt.Rows[i]["DATA_POST"]);
                    //int arm = 0;
                    //int slrane = 0;
                    //if (dt.Rows[i]["SLVARM_OT"] != DBNull.Value) arm = Convert.ToDateTime(dt.Rows[i]["SLVARM_DO"]).Year - Convert.ToDateTime(dt.Rows[i]["SLVARM_OT"]).Year;
                    //if (dt.Rows[i]["SL_RANE_OT"] != DBNull.Value) slrane = Convert.ToDateTime(dt.Rows[i]["SL_RANE_DO"]).Year - Convert.ToDateTime(dt.Rows[i]["SL_RANE_OT"]).Year;
                    //int y = calc_date.Year - (date_post.Year + arm + slrane);

                    if (item.years_c < 1) pos[31]++;
                    else if (item.years_c < 3 && item.years_c >= 1) pos[32]++;
                    else if (item.years_c < 5 && item.years_c >= 3) pos[33]++;
                    else if (item.years_c < 10 && item.years_c >= 5) pos[34]++;
                    else if (item.years_c < 20 && item.years_c >= 10) pos[35]++;
                    else if (item.years_c < 25 && item.years_c >= 20) pos[36]++;
                    else if (item.years_c >= 25) pos[37]++;

                    // Если 20 в льготке (пенсионеров++)
                    if (item.years_l >= 20) pos[44]++;
                    //// Смотрим есть ли в льготке, если да пенсионеры++
                    //else if (CalcStage(id2, false).y >= 20) pos[44]++;
                    //// Если достиг предельного возраста (пенсионеров++)
                    //else if (IsHiAge(Convert.ToDateTime(dt.Rows[i]["DATA_ROZD"]), Convert.ToInt16(dt.Rows[i]["ZVANIE"])) == true) pos[44]++;
                }
                StatusBar.Panels[1].ProgressBarValue = 0;
                progress.Value += 6;
                dt.Dispose();
            }
            #endregion


            #region // Распределение сотрудников по стажу работы в данной службе (опыту)
            if (code != "010" && code != "013")
            {
                PrStat("Распределение сотрудников по стажу работы в данной службе (опыту)");
                dt = DataProvider._getDataSQL(KConn, String.Format("SELECT VREMI_V_SL FROM AAQQ {0} AND FAMILIYA <> ''", sql_text));

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DateTime dateX = Convert.ToDateTime(dt.Rows[i]["VREMI_V_SL"]);
                    int y = calc_date.Year - dateX.Year;

                    if (y < 1) pos[38]++;
                    else if (y < 3 && y >= 1) pos[39]++;
                    else if (y < 5 && y >= 3) pos[40]++;
                    else if (y < 10 && y >= 5) pos[41]++;
                    else if (y < 20 && y >= 10) pos[42]++;
                    else if (y >= 20) pos[43]++;
                }

                progress.Value += 5;
                dt.Dispose();
            }
            #endregion


            #region // Принято за отчетный период
            //(без прибывших) в 2016 прибывшие из ФСКН и ФМС добавлять в принятые
            PrStat("Принято за отчетный период");
            dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1, KAT_POST, OBRAZ_LIC2, RTRIM(OTKYDA) AS OTKYDA FROM PRIEM {0} AND ( (KAT_POST NOT IN (101,102,104) OR KAT_POST IS NULL) OR (OTKYDA LIKE 'УФМС%') OR (OTKYDA LIKE 'УФСКН%') ) AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            pos[46] = dt.Rows.Count;
            if (pos[46] > 0) pos[45] = pos[46]; // Оформлялось кандидатов
                        
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i]["KAT_POST"] != DBNull.Value)
                {
                    int kat = 0;
                    if ( dt.Rows[i]["KAT_POST"] != DBNull.Value ) kat = Convert.ToInt16(dt.Rows[i]["KAT_POST"]);
                    string otkyda = dt.Rows[i]["OTKYDA"].ToString();
                    if (otkyda != "УФМС России по Ивановской области" && otkyda != "УФСКН России по Ивановской области")
                    {
                        
                        if (kat == 0 || kat == 201) pos[47]++;  // из гражданских организаций
                        else if (kat == 203) pos[48]++;         // из органов госвласти
                        else if (kat == 202) pos[49]++;         // из ВС и др.силовых структур
                        else if (kat == 103) pos[50]++;         // по окончании ОУ МВД
                        else if (kat == 204) pos[51]++;         // по окончании гражд. ОУ
                    }
                    else pos[49]++;
                }
                else pos[47]++;

                string obr = dt.Rows[i]["OBRAZ_LIC2"].ToString();   // с образованием

                if (obr == "10") pos[52]++;                         // высшим
                
                else if (obr == "20") pos[53]++;                         // ср.спец
                else if (obr == "50") pos[53]++;                         // неп.высшее в ср.спец.
                
                else if (obr == "30") pos[54]++;                         // среднее
                else if (obr == "40") pos[54]++;                         // неп.среднее в среднее.
            }
            progress.Value += 9;
            dt.Dispose();
            #endregion


            #region // Восстановлено сотрудников на службе
            //if (code != "010" && code != "013")
            //{
            //    PrStat("Восстановлено сотрудников на службе");
            //    if (code != "971" && code != "972" && code != "978" && code != "979")
            //    {
            //        // Если сельские участковые, заменяем key_1 на key_posl
            //        if (code != "201") pos[55] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI {0} AND STATUS IN ('1600','4001') AND DATA_OT >= {1} AND DATA_OT <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            //        else pos[55] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI {0} AND STATUS IN ('1600','4001') AND DATA_OT >= {1} AND DATA_OT <= {2}", sql_text.Replace("key_1", "key_posl"), Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            //    }
            //    else
            //    {
            //        string ch_sql_text = sql_text.Replace("potolok", "zvanie");
            //        pos[55] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI {0} AND STATUS IN ('1600','4001') AND DATA_OT >= {1} AND DATA_OT <= {2}", ch_sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            //    }
            //    // из них в судебном порядке
            //    pos[56] = pos[55];
            //}
            //progress.Value++;
            #endregion


            #region // Прибыло сотрудников из других регионов (др.регионы - только из системы МВД !)
            if (code != "010" && code != "013")
            {
                // в 2016 г. (без ФМС и ФСКН)
                PrStat("Прибыло сотрудников из других регионов");
                pos[57] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(FAMILIYA) FROM PRIEM {0} AND (KAT_POST=101 AND OTKYDA NOT LIKE 'УФМС%' AND OTKYDA NOT LIKE 'УФСКН%') AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            }
            #endregion


            #region // Выбыло сотрудников в другие регионы
            // Код выбытия - 104 - только для "хитрых"
            if (code != "010" && code != "013")
            {
                PrStat("Выбыло сотрудников в другие регионы");
                pos[58] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(FAMILIYA) FROM VYEZD {0} AND KOD_VYBYL NOT IN (104) AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            }
            #endregion

            #region // Прибыло сотрудников из других служб
            if (code != "010" && code != "013")
            {
                PrStat("Прибыло сотрудников из других служб");
                //// Расчет с 2015 года
                //// Выбираем ключи всех сотрудников текущей должностной группы у кого было назначение в текущем году...
                //dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM AAQQ {0} AND FAMILIYA <> '' AND DATA_VDOLZ >= {1} AND DATA_VDOLZ <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                //// Пробегаем всех
                //for (int i = 0; i < dt.Rows.Count; i++)
                //{
                //    // ключ
                //    string cid = dt.Rows[i]["KEY_1"].ToString();
                //    // Выбираем параметры предыдущей должности...
                //    DataTable prm = DataProvider._getDataSQL(KConn, String.Format("SELECT SLUZBA, DOLZNOST, ZVANIE FROM POSL_SPI WHERE KEY_POSL = {0} ORDER BY DATA_OT DESC", cid));
                //    // Если запись не единственная берем предыдущую...
                //    if (prm.Rows.Count > 1)
                //    {
                //        // Создаем строку дополнительный условий
                //        string extra_sql = string.Format("AND SLUZBA = {0} AND DOLZNOST = '{1}'", prm.Rows[1]["SLUZBA"], prm.Rows[1]["DOLZNOST"]);
                //        // Смотрим попал ли сотрудник в категорию текущей должностной группы при параметрах предыдущей службы
                //        int res = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(DOLZNOST) FROM AAQQ {0} {1}", sql_text, extra_sql));
                //        // Если попал - то сотрудник НЕ ПРИБЫЛ ИЗ ДРУГОЙ "СЛУЖБЫ"
                //        if (res <= 0) pos[59]++;
                //    }
                //    prm.Dispose();
                //}

                // Выбираем ключи тех у кого в тек.году сменилась должность (DATA_VDOLZ)
                dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM AAQQ {0} AND FAMILIYA <> '' AND DATA_VDOLZ >= {1} AND DATA_VDOLZ <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                // Проверяем была ли предыдущая служба в послужном другая...
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    // Проверяем есть ли такой ключ в PoslKeys1
                    if (PoslKeys1.Contains(Convert.ToInt16(dt.Rows[i]["KEY_1"]))) pos[59]++;
                }
                dt.Dispose();
            }
            #endregion

            #region // Выбыло сотрудников в другие службы
            if (code != "010" && code != "013")
            {
                PrStat("Выбыло сотрудников в другие службы");
                // Выбираем ключи всех у кого в тек.году сменилась должность (DATA_VDOLZ)
                dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM AAQQ WHERE DOLZNOST < '800000' AND FAMILIYA <> '' AND DATA_VDOLZ >= {0} AND DATA_VDOLZ <= {1}", Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                // Проверяем была ли предыдущая служба в послужном другая...
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    // Проверяем есть ли такой ключ в PoslKeys2
                    // Если в текущей форме присутствует признак службы проверяем не была ли она предыдущей...
                    if (current_dg_sluzb != "" && current_dg_sluzb != null)
                    {
                        string[] sl = current_dg_sluzb.Split(Convert.ToChar(","));
                        foreach (string sluzb in sl)
                        {
                            smen item = new smen();
                            item.key = Convert.ToInt16(dt.Rows[i]["KEY_1"]);
                            item.sluzba_prev = Convert.ToInt16(sluzb);

                            if (PoslKeys2.Contains(item)) pos[60]++;
                        }
                    }
                    // если нет просто сравниваем 1 и 2 
                    else
                    {
                        if (PoslKeys1.Contains(Convert.ToInt16(dt.Rows[i]["KEY_1"]))) pos[60]++;
                    }
                }
                dt.Dispose();
            }
            #endregion

            #region Сотрудников, которым приостановлена служба в ОВД
            if (code != "010" && code != "013")
            {
                pos[61] = 0; // Пока не реализовано ????
                progress.Value++;
            }
            #endregion

            #region // Уволено
            PrStat("Уволено");
            pos[62] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM ARCHIVE {0} AND PRICH_UV NOT IN ('1010','1019','1021') AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1, PRICH_UV FROM ARCHIVE {0} AND ZVANIE NOT IN (99) AND PRICH_UV NOT IN ('1010','1019','1021') AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            if (code != "010" && code != "013")
            {
                string uv_keys = GenKeys(dt.Rows);
                DataRowCollection uv = dt.Rows;
                dt.Dispose();

                // Уволено как не прошедшие испытательный срок
                // AND PRICH_UV = '1012' (разобраться - всех стажеров или только по статье не прошедших исп.срок)
                PrStat("Уволено как не прошедшие испытательный срок");
                pos[63] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(FAMILIYA) FROM ARCHIVE {0} AND PRICH_UV NOT IN ('1010','1019','1021') AND ZVANIE IN (99) AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));

                // со стажем службы в ОВД (распределение уволенных)
                PrStat("со стажем службы в ОВД (распределение уволенных)");
                dt = DataProvider._getDataSQL(KConn, String.Format("SELECT DATA_POST, DATA_UVOL, SL_RANE_OT, SL_RANE_DO FROM ARCHIVE {0} AND PRICH_UV NOT IN ('1010','1019','1021') AND ZVANIE NOT IN (99) AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DateTime date_post = Convert.ToDateTime(dt.Rows[i]["DATA_POST"]);
                    DateTime date_uvol = Convert.ToDateTime(dt.Rows[i]["DATA_UVOL"]);
                    DateTime date_r1 = Convert.ToDateTime("01.01.1900");
                    DateTime date_r2 = Convert.ToDateTime("01.01.1900");
                    if (dt.Rows[i]["SL_RANE_OT"] != DBNull.Value) date_r1 = Convert.ToDateTime(dt.Rows[i]["SL_RANE_OT"]);
                    if (dt.Rows[i]["SL_RANE_DO"] != DBNull.Value) date_r2 = Convert.ToDateTime(dt.Rows[i]["SL_RANE_DO"]);
                    int Years = 0;

                    DateDifference dd = new DateDifference(date_uvol, date_post);
                    if (dt.Rows[i]["SL_RANE_OT"] != DBNull.Value)
                    {
                        DateDifference dd2 = new DateDifference(date_r2, date_r1);
                        Years = dd.Years + dd2.Years;
                    }
                    else Years = dd.Years;

                    if (Years < 1)
                    {
                        pos[64]++;
                    }
                    else if (Years < 3 && Years >= 1) pos[65]++;
                    else if (Years < 5 && Years >= 3) pos[66]++;
                    else if (Years < 10 && Years >= 5) pos[67]++;
                    else if (Years < 20 && Years >= 10) pos[68]++;
                    else if (Years >= 20) pos[69]++;
                }
                dt.Dispose();

                // выпускников ОУ, проработавших менее 5-х лет
                PrStat("Уволено выпускников, проработавших менее 5-х лет");
                dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1, DAT_OKUZ, DATA_UVOL FROM ARCHIVE {0} AND PRICH_UV NOT IN ('1010','1019','1021') AND ((UCHZAV BETWEEN '200000' AND '300000') OR (UCHZAV BETWEEN '400000' AND '500000')) AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        int y = Convert.ToDateTime(dt.Rows[i]["DATA_UVOL"]).Year - Convert.ToInt16(dt.Rows[i]["DAT_OKUZ"]);
                        string id = dt.Rows[i]["KEY_1"].ToString();

                        if (y < 5) // Если проработал после окончания менее 5 лет ++
                        {
                            pos[71]++;
                            if (y < 1) pos[72]++; // если менее 1 года ++
                            // Выпускники очных
                            int res = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM PRIEM WHERE KEY_1 IN ({0}) AND KAT_POST IN (103)", id));
                            if (res > 0) pos[73]++;
                        }
                    }
                    // Заочники в остатке...
                    pos[74] = pos[71] - pos[73];
                }
                dt.Dispose();

                // Позиции с 76 по 116 (разбивка по основаниям увольнения)
                PrStat("Позиции с 76 по 116 (разбивка по основаниям увольнения)");
                for (int i = 0; i < uv.Count; i++)
                {
                    string pr = uv[i]["PRICH_UV"].ToString();

                    if (pr == "1030" || pr == "1046") pos[76]++;
                    if (pr == "1038" || pr == "1009" || pr == "1002" || pr == "1025") pos[77]++;
                    if (pr == "1037" || pr == "1011") pos[79]++;
                    if (pr == "1040" || pr == "1008") pos[80]++;
                    if (pr == "1003" || pr == "1026" || pr == "1024") pos[81]++;
                    if (pr == "1034") pos[82]++;
                    if (pr == "1007" || pr == "1020" || pr == "1036") pos[83]++;
                    if (pr == "1023" || pr == "1033") pos[84]++;
                    if (pr == "1044") pos[85]++;
                    if (pr == "1031" || pr == "1017" || pr == "1050") pos[86]++;
                    if (pr == "1014" || pr == "1022" || pr == "1028" || pr == "1047") pos[89]++;
                    if (pr == "1006") pos[90]++;
                    if (pr == "1015") pos[92]++;
                    if (pr == "1027" || pr == "1018") pos[95]++;
                    if (pr == "1049") pos[97]++;
                    if (pr == "1039" || pr == "1005") pos[99]++;
                    if (pr == "1029") pos[100]++;
                    if (pr == "1035") pos[103]++;
                    if (pr == "1013" || pr == "1043") pos[105]++;
                    if (pr == "1016") pos[106]++;
                    if (pr == "1041" || pr == "1004") pos[107]++;
                    if (pr == "1042") pos[109]++;
                }
                uv.Clear();
                progress.Value += 32;
            }
            #endregion


            #region // Исключено из реестра сотрудников по причине смерти, признания судом безвестно отсутствующим
            PrStat("Исключено из реестра сотрудников по причине смерти, признания судом безвестно отсутствующим");
            pos[117] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(FAMILIYA) FROM ARCHIVE {0} AND PRICH_UV IN ('1010','1019','1021') AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            #endregion


            #region // Сменилось руководителей (только руководящие ДГ:092, 117, 171, 202, 211, 281, 283, 361, 381, 992)
            // ! без оргштатных
            if (current_dg_lider)
            {
                // Сменилось всего (все кто ушел с должности)
                // Выбираем ключи всех сотрудников (действующих) у кого в текущем году сменилась должность
                PrStat("Сменилось руководящих работников");
                dt = DataProvider._getDataSQL(KConn, String.Format("select key_posl from posl_spi where data_ot BETWEEN {0} AND {1} and key_posl in (select distinct key_1 from aaqq where key_1 <> 0 and dolznost < '200000') order by key_posl", Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                string allkeys = "";
                // получили коллекцию ключей 
                DataRowCollection rc = dt.Rows;
                
                if (rc.Count > 0)
                {
                    string[] dg_sl = current_dg_sluzb.Split(Convert.ToChar(","));
                    // По каждому сотруднику смотрим 2 "верхние" записи в послужном (отсортированном по датам в порядке убывания)...
                    for (int i = 0; i < rc.Count; i++)
                    {
                        string key = rc[i]["key_posl"].ToString();
                        DataTable tmp = DataProvider._getDataSQL(KConn, String.Format("select top(2) sluzba from posl_spi where key_posl = {0} order by data_ot desc", key));
                        DataRowCollection sl = tmp.Rows;
                        string sl1 = sl[1]["sluzba"].ToString();
                        string sl0 = sl[0]["sluzba"].ToString();
                        if (tmp.Rows.Count == 2)
                        {
                            // если в предыдущей (из 2) записи служба из SLUZBA_SQL (и текущая другая) - значит была сменяемость.
                            if (dg_sl.Contains(sl1) == true && sl0 != sl1)
                            {
                                pos[118]++;
                                if (allkeys == "") allkeys += key;
                                else allkeys += "," + key;
                            }
                        }
                        tmp.Clear();
                        Application.DoEvents();
                    }
                    // добавляем уволенных, зачисленных в распоряжение и откомандированных
                    pos[118] += DataProvider._getDataSQLs(KConn,
                               String.Format("select a = (select count(key_1) from archive where sluzba in ({0}) and DAT_REG BETWEEN {1} AND {2}) + " +
                                             "(select count(key_1) from RESERV where sluzba in ({0}) and DATA_ZACH BETWEEN {1} AND {2}) + " +
                                             "(select count(key_1) from VYEZD where sluzba in ({0}) and DATA_UVOL BETWEEN {1} AND {2})", current_dg_sluzb, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                   
                }
                                
                dt.Dispose();
                if (pos[118] > 0 && allkeys!="")
                {
                    pos[119] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM PRIEM {0} AND DAT_REG >= {1}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate()));
                    PrStat("работали в должности менее года");
                    pos[120] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('9000','9002') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    PrStat("назначено на вышестоящую должность");
                    // Позиция 121 - пока руками...
                    pos[122] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('1100','1500') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    PrStat("назначено на равнозначную должность");
                    pos[123] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM VYEZD {0} AND KOD_VYBYL IN (103) AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    PrStat("направлено на обучение");
                    pos[124] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('1300','1400','1200','9400') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    PrStat("назначено на нижестоящую должность");
                    pos[125] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('1300','9400') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    PrStat("назначено на нижестоящую должность по служебному несоответствию");
                    pos[126] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('1400') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    PrStat("назначено на нижестоящую должность в порядке дисциплинарного взыскания");

                    pos[120] += pos[123];

                }
            }
            #endregion

            #region
            // Старый расчет до 2022 года...
            // Сменилось руководителей (только руководящие ДГ:092, 117, 171, 202, 211, 281, 283, 361, 381, 992)
            //// ! без оргштатных
            //if (current_dg_lider)
            //{
            //    // Сменилось всего
            //    // Выбираем ключи
            //    if (code == "020") sql_text = "where dolznost<'200000'";
            //    //|| code == "971" || code == "978") sql_text.Replace("500000", "200000");
            //    dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM AAQQ {0} AND KEY_1 IN (SELECT DISTINCT KEY_POSL FROM POSL_SPI WHERE STATUS NOT IN ('1701','6000','4000','4001','7000','5000','6001','2000','1000') AND DATA_OT >= {1} AND DATA_OT <= {2}) AND DATA_VDOLZ >= {1} AND DATA_VDOLZ <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            //    PrStat("Сменилось руководящих работников");
            //    pos[118] = dt.Rows.Count;
            //    string allkeys = GenKeys(dt.Rows);
            //    dt.Dispose();
            //    if (pos[118] > 0)
            //    {
            //        pos[119] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM PRIEM {0} AND DAT_REG >= {1}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate()));
            //        PrStat("работали в должности менее года");
            //        pos[120] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('9000','9002') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            //        PrStat("назначено на вышестоящую должность");
            //        // Позиция 121 - пока руками...
            //        pos[122] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('1100','1500') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            //        PrStat("назначено на равнозначную должность");
            //        pos[123] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM VYEZD {0} AND KOD_VYBYL IN (103) AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            //        PrStat("направлено на обучение");
            //        pos[124] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('1300','1400','1200','9400') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            //        PrStat("назначено на нижестоящую должность");
            //        pos[125] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('1300','9400') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            //        PrStat("назначено на нижестоящую должность по служебному несоответствию");
            //        pos[126] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('1400') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            //        PrStat("назначено на нижестоящую должность в порядке дисциплинарного взыскания");

            //        pos[120] += pos[123];

            //    }
            //}
            #endregion


            progress.Value = 0;

            SavePos2OneK(code);
        }

        // Расчет должностной группы ( код, текс запроса, расчитываемая позиция )
        public void CalcDG(string code, string sql_text, string fixed_pos)
        {
            progress.Maximum = 126;
            progress.Value = 0;
            string form_corrector = "";
            string sovmesh = "";
            int fpos = Convert.ToInt16(fixed_pos);

            #region // Должностей по штату...

            if (code != "010") form_corrector = "COUNT(DOLZNOST)";
            else form_corrector = "SUM(STAVKA_DLZ)";

            PrStat("Должностей по штату");
            if (fpos == 1)
            {
                if (code != "201") pos[1] = DataProvider._getDataSQLs(KConn, String.Format("SELECT {0} FROM AAQQ {1} AND DATA_SOKR IS NULL", form_corrector, sql_text));
                else pos[1] = Convert.ToInt16(UUPSelo_Stat.Text); // Если сельский участковый - пока ручками
            }

            // из них замещено
            if (code != "010") form_corrector = "COUNT(KEY_1)";
            else form_corrector = "SUM(STAVKA_PRS)";

            // Если аттестованные считаем по ключам, вольных по ставкам
            PrStat("из них замещено");
            if (fpos == 2)
            {
                pos[2] = DataProvider._getDataSQLs(KConn, String.Format("SELECT {0} FROM AAQQ {1} AND FAMILIYA <> ''", form_corrector, sql_text));
            }

            // в том числе женщинами
            PrStat("в том числе женщинами");
            if (fpos == 3)
            {
                pos[3] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' AND NACIONALN > 50", sql_text));
            }
            // в том числе кандидатами наук
            PrStat("в том числе кандидатами наук");
            if (fpos == 4)
            {
                pos[4] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' AND UCH_STEP = 1", sql_text));
            }

            // в том числе докторами наук
            PrStat("в том числе докторами наук");
            if (fpos == 5)
            {
                pos[5] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' AND UCH_STEP = 2", sql_text));
            }
            #endregion


            #region // Сотрудников с высшим профессиональным образованием...

            // Корректировка по 010 форме
            if (code == "010") sovmesh = "AND LICH_NOM_2 <> 'совмещ'";
            else sovmesh = "";
                        
            PrStat("Сотрудников с высшим профессиональным образованием");
            if (fpos >= 6 && fpos <= 13)
            {
                DataTable dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('10') ORDER BY KEY_1", sql_text, sovmesh));
                pos[6] = dt.Rows.Count;

                // в том числе юридическим
                PrStat("в том числе юридическим");
                pos[7] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('10') AND OBRAZ_LIC1 LIKE 'Юридическое%'", sql_text, sovmesh));

                // в том числе техническим
                PrStat("в том числе техническим");
                pos[8] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('10') AND OBRAZ_LIC1 LIKE 'Техническое%'", sql_text, sovmesh));

                // в том числе педагогическим
                PrStat("в том числе педагогическим");
                pos[9] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('10') AND OBRAZ_LIC1 LIKE 'Педагогическое%'", sql_text, sovmesh));

                // в том числе экономическим
                PrStat("в том числе экономическим");
                pos[10] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('10') AND OBRAZ_LIC1 LIKE 'Экономическое%'", sql_text, sovmesh));

                if (code != "010" && code != "013")
                {
                    //в том числе выпускников высших ОУ МВД России
                    if (dt.Rows.Count > 0)
                    {
                        PrStat("в том числе выпускников высших ОУ МВД России");
                        DataTable dt1 = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM LEARN WHERE KEY_1 IN ({0}) AND STATUS = 1 AND VID IN ('10') AND UCH_ZAV IN (SELECT P2 FROM SLVUCZ WHERE P1 LIKE '%МВД%' AND P2 < '300000') ORDER BY KEY_1", GenKeys(dt.Rows)));
                        pos[11] = dt1.Rows.Count;

                        if (pos[11] > 0)
                        {
                            //из них очной формы обучения
                            PrStat("из них очной формы обучения");
                            pos[12] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM PRIEM WHERE KEY_1 IN ({0}) AND KAT_POST IN (103)", GenKeys(dt1.Rows)));
                        }
                        dt1.Dispose();
                    }

                    // в том числе образование не соответствует направлению деятельности
                    PrStat("в том числе образование не соответствует направлению деятельности");
                    if (dt.Rows.Count > 0) pos[13] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ WHERE KEY_1 IN ({0}) AND KVAL = 0", GenKeys(dt.Rows)));
                    else pos[13] = 0;
                }
                dt.Dispose();
            }
            #endregion


            #region // Сотрудников со средним профессиональным образованием...
            PrStat("Сотрудников со средним профессиональным образованием");

            if (fpos >= 14 && fpos <= 21)
            {
                DataTable dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('20','50') ORDER BY KEY_1", sql_text, sovmesh));
                pos[14] = dt.Rows.Count;

                // в том числе юридическим
                PrStat("в том числе юридическим");
                pos[15] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('20','50') AND OBRAZ_LIC1 LIKE 'Юридическое%'", sql_text, sovmesh));

                // в том числе техническим
                PrStat("в том числе техническим");
                pos[16] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('20','50') AND OBRAZ_LIC1 LIKE 'Техническое%'", sql_text, sovmesh));

                // в том числе педагогическим
                PrStat("в том числе педагогическим");
                pos[17] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('20','50') AND OBRAZ_LIC1 LIKE 'Педагогическое%'", sql_text, sovmesh));

                // в том числе экономическим
                PrStat("в том числе экономическим");
                pos[18] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('20','50') AND OBRAZ_LIC1 LIKE 'Экономическое%'", sql_text, sovmesh));

                if (code != "010" && code != "013")
                {
                    //в том числе выпускников средних ОУ МВД России
                    if (dt.Rows.Count > 0)
                    {
                        PrStat("в том числе выпускников средних ОУ МВД России");
                        DataTable dt1 = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM LEARN WHERE KEY_1 IN ({0}) AND STATUS = 1 AND VID = '20' AND UCH_ZAV IN (SELECT P2 FROM SLVUCZ WHERE P1 LIKE '%МВД%' AND P2 BETWEEN '300000' AND '500000')", GenKeys(dt.Rows)));
                        pos[19] = dt1.Rows.Count;

                        if (pos[19] > 0)
                        {
                            //из них очной формы обучения
                            PrStat("из них очной формы обучения");
                            pos[20] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM PRIEM WHERE KEY_1 IN ({0}) AND KAT_POST IN (103)", GenKeys(dt1.Rows)));
                        }
                        dt1.Dispose();
                    }

                    // в том числе образование не соответствует направлению деятельности
                    PrStat("в том числе образование не соответствует направлению деятельности");
                    if (dt.Rows.Count > 0) pos[21] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ WHERE KEY_1 IN ({0}) AND KVAL = 0", GenKeys(dt.Rows)));
                    else pos[21] = 0;
                }
                dt.Dispose();
            }
            #endregion


            #region //Сотрудников, продолжающих обучение...
            if (fpos >= 22 && fpos <= 24)
            {
                PrStat("Сотрудников, продолжающих обучение");
                pos[22] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(FAMILIYA) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND UCHEBA = 1", sql_text, sovmesh));

                //в том числе	в ОУ МВД России
                PrStat("в том числе	в ОУ МВД России");
                pos[23] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(FAMILIYA) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND UCHEBA = 1 AND ((GDE_UCH BETWEEN '200000' AND '300000') OR (GDE_UCH BETWEEN '400000' AND '500000'))", sql_text, sovmesh));

                // в гражданских ОУ
                PrStat("в гражданских ОУ");
                pos[24] = pos[22] - pos[23];
            }
            #endregion


            #region //Сотрудников со средним (полным) общим образованием
            PrStat("Сотрудников со средним (полным) общим образованием");
            if (fpos == 25)
            {
                pos[25] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' {1} AND OBRAZ_LIC2 IN ('30','40')", sql_text, sovmesh));
            }
            #endregion


            #region // Распределение сотрудников по возрасту..
            PrStat("Распределение сотрудников по возрасту");
            if (fpos >= 26 && fpos <= 30)
            {
                pos[26] = pos[27] = pos[28] = pos[29] = pos[30] = 0;

                DataTable dt = DataProvider._getDataSQL(KConn, String.Format("SELECT DATA_ROZD FROM AAQQ {0} AND FAMILIYA <> '' {1}", sql_text, sovmesh));

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DateTime dateX = Convert.ToDateTime(dt.Rows[i]["DATA_ROZD"]);
                    int y = calc_date.Year - dateX.Year;

                    if (y <= 20) pos[26]++;
                    else if (y <= 30 && y >= 21) pos[27]++;
                    else if (y <= 40 && y >= 31) pos[28]++;
                    else if (y <= 55 && y >= 41) pos[29]++;
                    else if (y > 55)
                    {
                        pos[30]++;
                        if (code == "010" || code == "013") pos[44]++;
                    }
                }
                progress.Value += 4;
                dt.Dispose();
            }
            #endregion


            #region // Распределение сотрудников по стажу службы в ОВД (выслуга лет) + Сотрудников, имеющих право выхода на пенсию
            if (code != "010" && code != "013")
            {
                PrStat("Распределение сотрудников по стажу службы в ОВД (выслуга лет) + Сотрудников, имеющих право выхода на пенсию");
                if ( (fpos >= 31 && fpos <= 37) || (fpos == 44) )
                {
                    pos[31] = pos[32] = pos[33] = pos[34] = pos[35] = pos[36] = pos[37] = pos[44] = 0;

                    DataTable dt = DataProvider._getDataSQL(KConn, String.Format("SELECT DATA_POST, SL_RANE_OT, SL_RANE_DO, SLVARM_OT, SLVARM_DO, DATA_ROZD, ZVANIE, KEY_1 FROM AAQQ {0} AND FAMILIYA <> ''", sql_text));

                    StatusBar.Panels[1].ProgressBarMaxValue = dt.Rows.Count;
                    StatusBar.Panels[1].ProgressBarValue = 0;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        StatusBar.Panels[1].ProgressBarValue++;
                        string id2 = dt.Rows[i]["KEY_1"].ToString();
                        int y = CalcStage(id2, true).y;

                        //DateTime date_post = Convert.ToDateTime(dt.Rows[i]["DATA_POST"]);
                        //int arm = 0;
                        //int slrane = 0;
                        //if (dt.Rows[i]["SLVARM_OT"] != DBNull.Value) arm = Convert.ToDateTime(dt.Rows[i]["SLVARM_DO"]).Year - Convert.ToDateTime(dt.Rows[i]["SLVARM_OT"]).Year;
                        //if (dt.Rows[i]["SL_RANE_OT"] != DBNull.Value) slrane = Convert.ToDateTime(dt.Rows[i]["SL_RANE_DO"]).Year - Convert.ToDateTime(dt.Rows[i]["SL_RANE_OT"]).Year;
                        //int y = calc_date.Year - (date_post.Year + arm + slrane);

                        if (y < 1) pos[31]++;
                        else if (y < 3 && y >= 1) pos[32]++;
                        else if (y < 5 && y >= 3) pos[33]++;
                        else if (y < 10 && y >= 5) pos[34]++;
                        else if (y < 20 && y >= 10) pos[35]++;
                        else if (y < 25 && y >= 20) pos[36]++;
                        else if (y >= 25) pos[37]++;

                        // Если 20 календарей (пенсионеров++)
                        if (y >= 20) pos[44]++;
                        // Смотрим есть ли в льготке, если да пенсионеры++
                        else if (CalcStage(id2, false).y >= 20) pos[44]++;
                        // Если достиг предельного возраста (пенсионеров++)
                        else if (IsHiAge(Convert.ToDateTime(dt.Rows[i]["DATA_ROZD"]), Convert.ToInt16(dt.Rows[i]["ZVANIE"])) == true) pos[44]++;
                    }
                    StatusBar.Panels[1].ProgressBarValue = 0;
                    progress.Value += 6;
                    dt.Dispose();
                }
            }
            #endregion


            #region // Распределение сотрудников по стажу работы в данной службе (опыту)
            if (code != "010" && code != "013")
            {
                PrStat("Распределение сотрудников по стажу работы в данной службе (опыту)");
                if (fpos >= 38 && fpos <= 43)
                {
                    pos[38] = pos[39] = pos[40] = pos[41] = pos[42] = pos[43] = 0;
                    DataTable dt = DataProvider._getDataSQL(KConn, String.Format("SELECT VREMI_V_SL FROM AAQQ {0} AND FAMILIYA <> ''", sql_text));

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        DateTime dateX = Convert.ToDateTime(dt.Rows[i]["VREMI_V_SL"]);
                        int y = calc_date.Year - dateX.Year;

                        if (y < 1) pos[38]++;
                        else if (y < 3 && y >= 1) pos[39]++;
                        else if (y < 5 && y >= 3) pos[40]++;
                        else if (y < 10 && y >= 5) pos[41]++;
                        else if (y < 20 && y >= 10) pos[42]++;
                        else if (y >= 20) pos[43]++;
                    }

                    progress.Value += 5;
                    dt.Dispose();
                }
            }
            #endregion


            #region // Принято за отчетный период
            //(без прибывших) в 2016 прибывшие из ФСКН и ФМС добавлять в принятые
            PrStat("Принято за отчетный период");
            if (fpos >= 45 && fpos <= 54)
            {
                pos[45] = pos[46] = pos[47] = pos[48] = pos[49] = pos[50] = pos[51] = pos[52] = pos[53] = pos[54] = 0;
                DataTable dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1, KAT_POST, OBRAZ_LIC2, RTRIM(OTKYDA) AS OTKYDA FROM PRIEM {0} AND ( (KAT_POST NOT IN (101,102,104) OR KAT_POST IS NULL) OR (OTKYDA LIKE 'УФМС%') OR (OTKYDA LIKE 'УФСКН%') ) AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                pos[46] = dt.Rows.Count;
                if (pos[46] > 0) pos[45] = pos[46]; // Оформлялось кандидатов

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i]["KAT_POST"] != DBNull.Value)
                    {
                        int kat = Convert.ToInt16(dt.Rows[i]["KAT_POST"]);
                        string otkyda = dt.Rows[i]["OTKYDA"].ToString();
                        if (otkyda != "УФМС России по Ивановской области" && otkyda != "УФСКН России по Ивановской области")
                        {

                            if (kat == 0 || kat == 201) pos[47]++;  // из гражданских организаций
                            else if (kat == 203) pos[48]++;         // из органов госвласти
                            else if (kat == 202) pos[49]++;         // из ВС и др.силовых структур
                            else if (kat == 103) pos[50]++;         // по окончании ОУ МВД
                            else if (kat == 204) pos[51]++;         // по окончании гражд. ОУ
                        }
                        else pos[49]++;
                    }
                    else pos[47]++;

                    string obr = dt.Rows[i]["OBRAZ_LIC2"].ToString();   // с образованием
                    if (obr == "10") pos[52]++;                         // высшим
                    if (obr == "20") pos[53]++;                         // ср.спец
                    if (obr == "30") pos[54]++;                         // среднее
                    if (obr == "40") pos[54]++;                         // неп.среднее
                }
                progress.Value += 9;
                dt.Dispose();
            }
            #endregion


            #region // Восстановлено сотрудников на службе
            if (code != "010" && code != "013")
            {
                PrStat("Восстановлено сотрудников на службе");
                if (fpos == 55 || fpos == 56)
                {
                    pos[55] = pos[56] = 0;
                    if (code != "971" && code != "972" && code != "978" && code != "979")
                    {
                        // Если сельские участковые, заменяем key_1 на key_posl
                        if (code != "201") pos[55] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI {0} AND STATUS IN ('1600','4001') AND DATA_OT >= {1} AND DATA_OT <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                        else pos[55] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI {0} AND STATUS IN ('1600','4001') AND DATA_OT >= {1} AND DATA_OT <= {2}", sql_text.Replace("key_1", "key_posl"), Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    }
                    else
                    {
                        string ch_sql_text = sql_text.Replace("potolok", "zvanie");
                        pos[55] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI {0} AND STATUS IN ('1600','4001') AND DATA_OT >= {1} AND DATA_OT <= {2}", ch_sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    }
                    // из них в судебном порядке
                    pos[56] = pos[55];
                }
                progress.Value++;
            }
            #endregion


            #region // Прибыло сотрудников из других регионов (др.регионы - только из системы МВД !)
            if (code != "010" && code != "013")
            {
                // в 2016 г. (без ФМС и ФСКН)
                PrStat("Прибыло сотрудников из других регионов");
                if (fpos == 57)
                {
                    pos[57] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(FAMILIYA) FROM PRIEM {0} AND (KAT_POST=101 AND OTKYDA NOT LIKE 'УФМС%' AND OTKYDA NOT LIKE 'УФСКН%') AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                }
            }
            #endregion


            #region // Выбыло сотрудников в другие регионы
            // Код выбытия - 104 - только для "хитрых"
            if (code != "010" && code != "013")
            {
                PrStat("Выбыло сотрудников в другие регионы");
                if (fpos == 58)
                {
                    pos[58] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(FAMILIYA) FROM VYEZD {0} AND KOD_VYBYL NOT IN (104) AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                }
            }
            #endregion


            #region // Прибыло сотрудников из других служб
            if (code != "010" && code != "013")
            {
                PrStat("Прибыло сотрудников из других служб");
                if (fpos == 59)
                {
                    pos[59] = 0;
                    //// Расчет с 2015 года
                    //// Выбираем ключи всех сотрудников текущей должностной группы у кого было назначение в текущем году...
                    //dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM AAQQ {0} AND FAMILIYA <> '' AND DATA_VDOLZ >= {1} AND DATA_VDOLZ <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    //// Пробегаем всех
                    //for (int i = 0; i < dt.Rows.Count; i++)
                    //{
                    //    // ключ
                    //    string cid = dt.Rows[i]["KEY_1"].ToString();
                    //    // Выбираем параметры предыдущей должности...
                    //    DataTable prm = DataProvider._getDataSQL(KConn, String.Format("SELECT SLUZBA, DOLZNOST, ZVANIE FROM POSL_SPI WHERE KEY_POSL = {0} ORDER BY DATA_OT DESC", cid));
                    //    // Если запись не единственная берем предыдущую...
                    //    if (prm.Rows.Count > 1)
                    //    {
                    //        // Создаем строку дополнительный условий
                    //        string extra_sql = string.Format("AND SLUZBA = {0} AND DOLZNOST = '{1}'", prm.Rows[1]["SLUZBA"], prm.Rows[1]["DOLZNOST"]);
                    //        // Смотрим попал ли сотрудник в категорию текущей должностной группы при параметрах предыдущей службы
                    //        int res = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(DOLZNOST) FROM AAQQ {0} {1}", sql_text, extra_sql));
                    //        // Если попал - то сотрудник НЕ ПРИБЫЛ ИЗ ДРУГОЙ "СЛУЖБЫ"
                    //        if (res <= 0) pos[59]++;
                    //    }
                    //    prm.Dispose();
                    //}

                    // Выбираем ключи тех у кого в тек.году сменилась должность (DATA_VDOLZ)
                    DataTable dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM AAQQ {0} AND FAMILIYA <> '' AND DATA_VDOLZ >= {1} AND DATA_VDOLZ <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    // Проверяем была ли предыдущая служба в послужном другая...
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        // Проверяем есть ли такой ключ в PoslKeys1
                        if (PoslKeys1.Contains(Convert.ToInt16(dt.Rows[i]["KEY_1"]))) pos[59]++;
                    }
                    dt.Dispose();
                }
            }
            #endregion


            #region // Выбыло сотрудников в другие службы
            if (code != "010" && code != "013")
            {
                PrStat("Выбыло сотрудников в другие службы");
                if (fpos == 60)
                {
                    pos[60] = 0;
                    // Выбираем ключи всех у кого в тек.году сменилась должность (DATA_VDOLZ)
                    DataTable dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM AAQQ WHERE DOLZNOST < '800000' AND FAMILIYA <> '' AND DATA_VDOLZ >= {0} AND DATA_VDOLZ <= {1}", Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    // Проверяем была ли предыдущая служба в послужном другая...
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        // Проверяем есть ли такой ключ в PoslKeys2
                        // Если в текущей форме присутствует признак службы проверяем не была ли она предыдущей...
                        if (current_dg_sluzb != "" && current_dg_sluzb != null)
                        {
                            string[] sl = current_dg_sluzb.Split(Convert.ToChar(","));
                            foreach (string sluzb in sl)
                            {
                                smen item = new smen();
                                item.key = Convert.ToInt16(dt.Rows[i]["KEY_1"]);
                                item.sluzba_prev = Convert.ToInt16(sluzb);

                                if (PoslKeys2.Contains(item)) pos[60]++;
                            }
                        }
                        // если нет просто сравниваем 1 и 2 
                        else
                        {
                            if (PoslKeys1.Contains(Convert.ToInt16(dt.Rows[i]["KEY_1"]))) pos[60]++;
                        }
                    }
                    dt.Dispose();
                }
            }
            #endregion


            #region Сотрудников, которым приостановлена служба в ОВД
            if (code != "010" && code != "013")
            {
                pos[61] = 0; // Пока не реализовано ????
                progress.Value++;
            }
            #endregion


            #region // Уволено
            PrStat("Уволено");
            if (fpos == 62)
            {
                pos[62] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM ARCHIVE {0} AND PRICH_UV NOT IN ('1010','1019','1021') AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            }

            if (fpos >= 63 && fpos <= 116)
            {
                int k = 63;
                do
                {
                    pos[k] = 0;
                    k++;
                } while (k != 117);

                DataTable dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1, PRICH_UV FROM ARCHIVE {0} AND ZVANIE NOT IN (99) AND PRICH_UV NOT IN ('1010','1019','1021') AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));

                if (code != "010" && code != "013")
                {
                    string uv_keys = GenKeys(dt.Rows);
                    DataRowCollection uv = dt.Rows;
                    dt.Dispose();

                    // Уволено как не прошедшие испытательный срок
                    // AND PRICH_UV = '1012' (разобраться - всех стажеров или только по статье не прошедших исп.срок)
                    PrStat("Уволено как не прошедшие испытательный срок");
                    pos[63] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(FAMILIYA) FROM ARCHIVE {0} AND PRICH_UV = '1012' AND ZVANIE IN (99) AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));

                    // со стажем службы в ОВД (распределение уволенных)
                    PrStat("со стажем службы в ОВД (распределение уволенных)");
                    dt = DataProvider._getDataSQL(KConn, String.Format("SELECT DATA_POST, DATA_UVOL, SL_RANE_OT, SL_RANE_DO FROM ARCHIVE {0} AND PRICH_UV NOT IN ('1010','1019','1021') AND ZVANIE NOT IN (99) AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        DateTime date_post = Convert.ToDateTime(dt.Rows[i]["DATA_POST"]);
                        DateTime date_uvol = Convert.ToDateTime(dt.Rows[i]["DATA_UVOL"]);
                        DateTime date_r1 = Convert.ToDateTime("01.01.1900");
                        DateTime date_r2 = Convert.ToDateTime("01.01.1900");
                        if (dt.Rows[i]["SL_RANE_OT"] != DBNull.Value) date_r1 = Convert.ToDateTime(dt.Rows[i]["SL_RANE_OT"]);
                        if (dt.Rows[i]["SL_RANE_DO"] != DBNull.Value) date_r2 = Convert.ToDateTime(dt.Rows[i]["SL_RANE_DO"]);
                        int Years = 0;

                        DateDifference dd = new DateDifference(date_uvol, date_post);
                        if (dt.Rows[i]["SL_RANE_OT"] != DBNull.Value)
                        {
                            DateDifference dd2 = new DateDifference(date_r2, date_r1);
                            Years = dd.Years + dd2.Years;
                        }
                        else Years = dd.Years;

                        if (Years < 1)
                        {
                            pos[64]++;
                        }
                        else if (Years < 3 && Years >= 1) pos[65]++;
                        else if (Years < 5 && Years >= 3) pos[66]++;
                        else if (Years < 10 && Years >= 5) pos[67]++;
                        else if (Years < 20 && Years >= 10) pos[68]++;
                        else if (Years >= 20) pos[69]++;
                    }
                    dt.Dispose();

                    // выпускников ОУ, проработавших менее 5-х лет
                    PrStat("Уволено выпускников, проработавших менее 5-х лет");
                    dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1, DAT_OKUZ, DATA_UVOL FROM ARCHIVE {0} AND PRICH_UV NOT IN ('1010','1019','1021') AND ((UCHZAV BETWEEN '200000' AND '300000') OR (UCHZAV BETWEEN '400000' AND '500000')) AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    if (dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            int y = Convert.ToDateTime(dt.Rows[i]["DATA_UVOL"]).Year - Convert.ToInt16(dt.Rows[i]["DAT_OKUZ"]);
                            string id = dt.Rows[i]["KEY_1"].ToString();

                            if (y < 5) // Если проработал после окончания менее 5 лет ++
                            {
                                pos[71]++;
                                if (y < 1) pos[72]++; // если менее 1 года ++
                                // Выпускники очных
                                int res = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM PRIEM WHERE KEY_1 IN ({0}) AND KAT_POST IN (103)", id));
                                if (res > 0) pos[73]++;
                            }
                        }
                        // Заочники в остатке...
                        pos[74] = pos[71] - pos[73];
                    }
                    dt.Dispose();

                    // Позиции с 76 по 116 (разбивка по основаниям увольнения)
                    PrStat("Позиции с 76 по 116 (разбивка по основаниям увольнения)");
                    for (int i = 0; i < uv.Count; i++)
                    {
                        string pr = uv[i]["PRICH_UV"].ToString();

                        if (pr == "1030" || pr == "1046") pos[76]++;
                        if (pr == "1038" || pr == "1009" || pr == "1002" || pr == "1025") pos[77]++;
                        if (pr == "1037" || pr == "1011") pos[79]++;
                        if (pr == "1040" || pr == "1008") pos[80]++;
                        if (pr == "1003" || pr == "1026" || pr == "1024") pos[81]++;
                        if (pr == "1034") pos[82]++;
                        if (pr == "1007" || pr == "1020" || pr == "1036") pos[83]++;
                        if (pr == "1023" || pr == "1033") pos[84]++;
                        if (pr == "1044") pos[85]++;
                        if (pr == "1031" || pr == "1017" || pr == "1050") pos[86]++;
                        if (pr == "1014" || pr == "1022" || pr == "1028" || pr == "1047") pos[89]++;
                        if (pr == "1006") pos[90]++;
                        if (pr == "1015") pos[92]++;
                        if (pr == "1027" || pr == "1018") pos[95]++;
                        if (pr == "1049") pos[97]++;
                        if (pr == "1039" || pr == "1005") pos[99]++;
                        if (pr == "1029") pos[100]++;
                        if (pr == "1035") pos[103]++;
                        if (pr == "1013" || pr == "1043") pos[105]++;
                        if (pr == "1016") pos[106]++;
                        if (pr == "1041" || pr == "1004") pos[107]++;
                        if (pr == "1042") pos[109]++;
                    }
                    uv.Clear();
                    progress.Value += 32;
                }
            }
            #endregion


            #region // Исключено из реестра сотрудников по причине смерти, признания судом безвестно отсутствующим
            PrStat("Исключено из реестра сотрудников по причине смерти, признания судом безвестно отсутствующим");
            if (fpos == 117)
            {
                pos[117] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(FAMILIYA) FROM ARCHIVE {0} AND PRICH_UV IN ('1010','1019','1021') AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            }
            #endregion


            #region // Сменилось руководителей (только руководящие ДГ:092, 117, 171, 202, 211, 281, 283, 361, 381, 992)
            // ! без оргштатных
            if (current_dg_lider)
            {
                // Сменилось всего (все кто ушел с должности)
                // Выбираем ключи всех сотрудников (действующих) у кого в текущем году сменилась должность
                PrStat("Сменилось руководящих работников");
                DataTable dt = DataProvider._getDataSQL(KConn, String.Format("select key_posl from posl_spi where data_ot BETWEEN {0} AND {1} and key_posl in (select distinct key_1 from aaqq where key_1 <> 0 and dolznost < '200000') order by key_posl", Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                string allkeys = "";
                // получили коллекцию ключей 
                DataRowCollection rc = dt.Rows;

                if (rc.Count > 0)
                {
                    string[] dg_sl = current_dg_sluzb.Split(Convert.ToChar(","));
                    // По каждому сотруднику смотрим 2 "верхние" записи в послужном (отсортированном по датам в порядке убывания)...
                    for (int i = 0; i < rc.Count; i++)
                    {
                        string key = rc[i]["key_posl"].ToString();
                        DataTable tmp = DataProvider._getDataSQL(KConn, String.Format("select top(2) sluzba from posl_spi where key_posl = {0} order by data_ot desc", key));
                        DataRowCollection sl = tmp.Rows;
                        string sl1 = sl[1]["sluzba"].ToString();
                        string sl0 = sl[0]["sluzba"].ToString();
                        if (tmp.Rows.Count == 2)
                        {
                            // если в предыдущей (из 2) записи служба из SLUZBA_SQL (и текущая другая) - значит была сменяемость.
                            if (dg_sl.Contains(sl1) == true && sl0 != sl1)
                            {
                                pos[118]++;
                                if (allkeys == "") allkeys += key;
                                else allkeys += "," + key;
                            }
                        }
                        tmp.Clear();
                        Application.DoEvents();
                    }
                    // добавляем уволенных, зачисленных в распоряжение и откомандированных
                    pos[118] += DataProvider._getDataSQLs(KConn,
                               String.Format("select a = (select count(key_1) from archive where sluzba in ({0}) and DAT_REG BETWEEN {1} AND {2}) + " +
                                             "(select count(key_1) from RESERV where sluzba in ({0}) and DATA_ZACH BETWEEN {1} AND {2}) + " +
                                             "(select count(key_1) from VYEZD where sluzba in ({0}) and DATA_UVOL BETWEEN {1} AND {2})", current_dg_sluzb, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));

                }

                dt.Dispose();
                if (pos[118] > 0 && allkeys != "")
                {
                    pos[119] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM PRIEM {0} AND DAT_REG >= {1}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate()));
                    PrStat("работали в должности менее года");
                    pos[120] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('9000','9002') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    PrStat("назначено на вышестоящую должность");
                    // Позиция 121 - пока руками...
                    pos[122] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('1100','1500') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    PrStat("назначено на равнозначную должность");
                    pos[123] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM VYEZD {0} AND KOD_VYBYL IN (103) AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    PrStat("направлено на обучение");
                    pos[124] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('1300','1400','1200','9400') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    PrStat("назначено на нижестоящую должность");
                    pos[125] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('1300','9400') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    PrStat("назначено на нижестоящую должность по служебному несоответствию");
                    pos[126] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('1400') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    PrStat("назначено на нижестоящую должность в порядке дисциплинарного взыскания");

                    pos[120] += pos[123];

                }
            }
            #endregion

            #region
            // Старый расчет до 2022 года
            // Сменилось руководителей (только руководящие ДГ:092, 117, 171, 202, 211, 281, 283, 361, 381, 992)
            //// ! без оргштатных
            //if (current_dg_lider)
            //{
            //    // Сменилось всего
            //    // Выбираем ключи
            //    if (code == "020") sql_text = "where dolznost<'200000'";
            //    //|| code == "971" || code == "978") sql_text.Replace("500000", "200000");
            //    if (fpos >= 118 && fpos <= 126)
            //    {
            //        DataTable dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM AAQQ {0} AND KEY_1 IN (SELECT DISTINCT KEY_POSL FROM POSL_SPI WHERE STATUS NOT IN ('1701','6000','4000','4001','7000','5000','6001','2000','1000') AND DATA_OT >= {1} AND DATA_OT <= {2}) AND DATA_VDOLZ >= {1} AND DATA_VDOLZ <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            //        PrStat("Сменилось руководящих работников");
            //        pos[118] = dt.Rows.Count;
            //        string allkeys = GenKeys(dt.Rows);
            //        dt.Dispose();
            //        if (pos[118] > 0)
            //        {
            //            pos[119] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM PRIEM {0} AND DAT_REG >= {1}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate()));
            //            PrStat("работали в должности менее года");
            //            pos[120] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('9000','9002') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            //            PrStat("назначено на вышестоящую должность");
            //            // Позиция 121 - пока руками...
            //            pos[122] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('1100','1500') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            //            PrStat("назначено на равнозначную должность");
            //            pos[123] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM VYEZD {0} AND KOD_VYBYL IN (103) AND DAT_REG >= {1} AND DAT_REG <= {2}", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            //            PrStat("направлено на обучение");
            //            pos[124] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('1300','1400','1200','9400') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            //            PrStat("назначено на нижестоящую должность");
            //            pos[125] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('1300','9400') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            //            PrStat("назначено на нижестоящую должность по служебному несоответствию");
            //            pos[126] = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_POSL) FROM POSL_SPI WHERE KEY_POSL IN ({0}) AND STATUS IN ('1400') AND DATA_OT >= {1} AND DATA_OT <= {2}", allkeys, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            //            PrStat("назначено на нижестоящую должность в порядке дисциплинарного взыскания");

            //            pos[120] += pos[123];

            //        }
            //    }
            //}
            #endregion

            progress.Value = 0;

            SavePos2OneK(code);
        }


        // Информирование о рассчитываемой ДГ
        public void PrStat(string txt)
        {
            StatusBar.Panels[0].Text = String.Format("[{0}] - {1} Расчитываю: {2}", current_dg_code, current_dg_name, txt);
            progress.Value++;
            Application.DoEvents();
        }

        // Расчет некомплекта по ДГ
        private double CalcNek(int code)
        {
            double s = Convert.ToDouble(oneK[code, 1]);
            double z = Convert.ToDouble(oneK[code, 2]);
            return Math.Round((s - z) / s * 100, 1);
        }

        // Заполнение формы редактирования
        public void FillEditPanel(string code, int swich)
        {
            if ( swich == 1 && mainTabControl.SelectedTabPageIndex != 1) mainTabControl.SelectedTabPageIndex = 1;

            int icode = Convert.ToInt16(code);

            Edit_groupBox.Text = String.Format("[{0}] - {1}",current_dg_code, current_dg_name);
            nek_info.Text = String.Format("Некомплект: {0} ед. ({1}) ", oneK[icode,1]-oneK[icode,2], CalcNek(icode));

            edit_pos1.Text = oneK[icode,1].ToString();
            edit_pos2.Text = oneK[icode,2].ToString();
            edit_pos3.Text = oneK[icode,3].ToString();
            edit_pos4.Text = oneK[icode,4].ToString();
            edit_pos5.Text = oneK[icode,5].ToString();
            edit_pos6.Text = oneK[icode,6].ToString();
            edit_pos7.Text = oneK[icode,7].ToString();
            edit_pos8.Text = oneK[icode,8].ToString();
            edit_pos9.Text = oneK[icode,9].ToString();
            edit_pos10.Text = oneK[icode,10].ToString();
            edit_pos11.Text = oneK[icode,11].ToString();
            edit_pos12.Text = oneK[icode,12].ToString();
            edit_pos13.Text = oneK[icode,13].ToString();
            edit_pos14.Text = oneK[icode,14].ToString();
            edit_pos15.Text = oneK[icode,15].ToString();
            edit_pos16.Text = oneK[icode, 16].ToString();
            edit_pos17.Text = oneK[icode, 17].ToString();
            edit_pos18.Text = oneK[icode, 18].ToString();
            edit_pos19.Text = oneK[icode, 19].ToString();
            edit_pos20.Text = oneK[icode, 20].ToString();
            edit_pos21.Text = oneK[icode, 21].ToString();
            edit_pos22.Text = oneK[icode, 22].ToString();
            edit_pos23.Text = oneK[icode, 23].ToString();
            edit_pos24.Text = oneK[icode, 24].ToString();
            edit_pos25.Text = oneK[icode, 25].ToString();
            edit_pos26.Text = oneK[icode, 26].ToString();
            edit_pos27.Text = oneK[icode, 27].ToString();
            edit_pos28.Text = oneK[icode, 28].ToString();
            edit_pos29.Text = oneK[icode, 29].ToString();
            edit_pos30.Text = oneK[icode, 30].ToString();
            edit_pos31.Text = oneK[icode, 31].ToString();
            edit_pos32.Text = oneK[icode, 32].ToString();
            edit_pos33.Text = oneK[icode, 33].ToString();
            edit_pos34.Text = oneK[icode, 34].ToString();
            edit_pos35.Text = oneK[icode, 35].ToString();
            edit_pos36.Text = oneK[icode, 36].ToString();
            edit_pos37.Text = oneK[icode, 37].ToString();
            edit_pos38.Text = oneK[icode, 38].ToString();
            edit_pos39.Text = oneK[icode, 39].ToString();
            edit_pos40.Text = oneK[icode, 40].ToString();
            edit_pos41.Text = oneK[icode, 41].ToString();
            edit_pos42.Text = oneK[icode, 42].ToString();
            edit_pos43.Text = oneK[icode, 43].ToString();
            edit_pos44.Text = oneK[icode, 44].ToString();
            edit_pos45.Text = oneK[icode, 45].ToString();
            edit_pos46.Text = oneK[icode, 46].ToString();
            edit_pos47.Text = oneK[icode, 47].ToString();
            edit_pos48.Text = oneK[icode, 48].ToString();
            edit_pos49.Text = oneK[icode, 49].ToString();
            edit_pos50.Text = oneK[icode, 50].ToString();
            edit_pos51.Text = oneK[icode, 51].ToString();
            edit_pos52.Text = oneK[icode, 52].ToString();
            edit_pos53.Text = oneK[icode, 53].ToString();
            edit_pos54.Text = oneK[icode, 54].ToString();
            edit_pos55.Text = oneK[icode, 55].ToString();
            edit_pos56.Text = oneK[icode, 56].ToString();
            edit_pos57.Text = oneK[icode, 57].ToString();
            edit_pos58.Text = oneK[icode, 58].ToString();
            edit_pos59.Text = oneK[icode, 59].ToString();
            edit_pos60.Text = oneK[icode, 60].ToString();
            edit_pos61.Text = oneK[icode, 61].ToString();

            if (swich == 1 && mainTabControl.SelectedTabPageIndex != 1) mainTabControl.SelectedTabPageIndex = 2;

            edit_pos62.Text = oneK[icode, 62].ToString();
            edit_pos63.Text = oneK[icode, 63].ToString();
            edit_pos64.Text = oneK[icode, 64].ToString();
            edit_pos65.Text = oneK[icode, 65].ToString();
            edit_pos66.Text = oneK[icode, 66].ToString();
            edit_pos67.Text = oneK[icode, 67].ToString();
            edit_pos68.Text = oneK[icode, 68].ToString();
            edit_pos69.Text = oneK[icode, 69].ToString();
            edit_pos70.Text = oneK[icode, 70].ToString();
            edit_pos71.Text = oneK[icode, 71].ToString();
            edit_pos72.Text = oneK[icode, 72].ToString();
            edit_pos73.Text = oneK[icode, 73].ToString();
            edit_pos74.Text = oneK[icode, 74].ToString();
            edit_pos75.Text = oneK[icode, 75].ToString();
            edit_pos76.Text = oneK[icode, 76].ToString();
            edit_pos77.Text = oneK[icode, 77].ToString();
            edit_pos78.Text = oneK[icode, 78].ToString();
            edit_pos79.Text = oneK[icode, 79].ToString();
            edit_pos80.Text = oneK[icode, 80].ToString();
            edit_pos81.Text = oneK[icode, 81].ToString();
            edit_pos82.Text = oneK[icode, 82].ToString();
            edit_pos83.Text = oneK[icode, 83].ToString();
            edit_pos84.Text = oneK[icode, 84].ToString();
            edit_pos85.Text = oneK[icode, 85].ToString();
            edit_pos86.Text = oneK[icode, 86].ToString();
            edit_pos87.Text = oneK[icode, 87].ToString();
            edit_pos88.Text = oneK[icode, 88].ToString();
            edit_pos89.Text = oneK[icode, 89].ToString();
            edit_pos90.Text = oneK[icode, 90].ToString();
            edit_pos91.Text = oneK[icode, 91].ToString();
            edit_pos92.Text = oneK[icode, 92].ToString();
            edit_pos93.Text = oneK[icode, 93].ToString();
            edit_pos94.Text = oneK[icode, 94].ToString();
            edit_pos95.Text = oneK[icode, 95].ToString();
            edit_pos96.Text = oneK[icode, 96].ToString();
            edit_pos97.Text = oneK[icode, 97].ToString();
            edit_pos98.Text = oneK[icode, 98].ToString();
            edit_pos99.Text = oneK[icode, 99].ToString();
            edit_pos100.Text = oneK[icode, 100].ToString();
            edit_pos101.Text = oneK[icode, 101].ToString();
            edit_pos102.Text = oneK[icode, 102].ToString();
            edit_pos103.Text = oneK[icode, 103].ToString();
            edit_pos104.Text = oneK[icode, 104].ToString();

            if (swich == 1 && mainTabControl.SelectedTabPageIndex != 1) mainTabControl.SelectedTabPageIndex = 3;

            edit_pos105.Text = oneK[icode, 105].ToString();
            edit_pos106.Text = oneK[icode, 106].ToString();
            edit_pos107.Text = oneK[icode, 107].ToString();
            edit_pos108.Text = oneK[icode, 108].ToString();
            edit_pos109.Text = oneK[icode, 109].ToString();
            edit_pos110.Text = oneK[icode, 110].ToString();
            edit_pos111.Text = oneK[icode, 111].ToString();
            edit_pos112.Text = oneK[icode, 112].ToString();
            edit_pos113.Text = oneK[icode, 113].ToString();
            edit_pos114.Text = oneK[icode, 114].ToString();
            edit_pos115.Text = oneK[icode, 115].ToString();
            edit_pos116.Text = oneK[icode, 116].ToString();
            edit_pos117.Text = oneK[icode, 117].ToString();
            edit_pos118.Text = oneK[icode, 118].ToString();
            edit_pos119.Text = oneK[icode, 119].ToString();
            edit_pos120.Text = oneK[icode, 120].ToString();
            edit_pos121.Text = oneK[icode, 121].ToString();
            edit_pos122.Text = oneK[icode, 122].ToString();
            edit_pos123.Text = oneK[icode, 123].ToString();
            edit_pos124.Text = oneK[icode, 124].ToString();
            edit_pos125.Text = oneK[icode, 125].ToString();
            edit_pos126.Text = oneK[icode, 126].ToString();

            int sum = oneK[icode,63];
            for (int i = 76; i < 117; i++) sum += oneK[icode,i];
            uvol_sum.Text = "Сумма:" + sum.ToString();
        }

        // Обновление данных в Pos из панели редактирования
        public void RefreshPosFromEditPanel()
        {
            pos[1] = Convert.ToInt16(edit_pos1.Text);
            pos[2] = Convert.ToInt16(edit_pos2.Text);
            pos[3] = Convert.ToInt16(edit_pos3.Text);
            pos[4] = Convert.ToInt16(edit_pos4.Text);
            pos[5] = Convert.ToInt16(edit_pos5.Text);
            pos[6] = Convert.ToInt16(edit_pos6.Text);
            pos[7] = Convert.ToInt16(edit_pos7.Text);
            pos[8] = Convert.ToInt16(edit_pos8.Text);
            pos[9] = Convert.ToInt16(edit_pos9.Text);
            pos[10] = Convert.ToInt16(edit_pos10.Text);
            pos[11] = Convert.ToInt16(edit_pos11.Text);
            pos[12] = Convert.ToInt16(edit_pos12.Text);
            pos[13] = Convert.ToInt16(edit_pos13.Text);
            pos[14] = Convert.ToInt16(edit_pos14.Text);
            pos[15] = Convert.ToInt16(edit_pos15.Text);
            pos[16] = Convert.ToInt16(edit_pos16.Text);
            pos[17] = Convert.ToInt16(edit_pos17.Text);
            pos[18] = Convert.ToInt16(edit_pos18.Text);
            pos[19] = Convert.ToInt16(edit_pos19.Text);
            pos[20] = Convert.ToInt16(edit_pos20.Text);
            pos[21] = Convert.ToInt16(edit_pos21.Text);
            pos[22] = Convert.ToInt16(edit_pos22.Text);
            pos[23] = Convert.ToInt16(edit_pos23.Text);
            pos[24] = Convert.ToInt16(edit_pos24.Text);
            pos[25] = Convert.ToInt16(edit_pos25.Text);
            pos[26] = Convert.ToInt16(edit_pos26.Text);
            pos[27] = Convert.ToInt16(edit_pos27.Text);
            pos[28] = Convert.ToInt16(edit_pos28.Text);
            pos[29] = Convert.ToInt16(edit_pos29.Text);
            pos[30] = Convert.ToInt16(edit_pos30.Text);
            pos[31] = Convert.ToInt16(edit_pos31.Text);
            pos[32] = Convert.ToInt16(edit_pos32.Text);
            pos[33] = Convert.ToInt16(edit_pos33.Text);
            pos[34] = Convert.ToInt16(edit_pos34.Text);
            pos[35] = Convert.ToInt16(edit_pos35.Text);
            pos[36] = Convert.ToInt16(edit_pos36.Text);
            pos[37] = Convert.ToInt16(edit_pos37.Text);
            pos[38] = Convert.ToInt16(edit_pos38.Text);
            pos[39] = Convert.ToInt16(edit_pos39.Text);
            pos[40] = Convert.ToInt16(edit_pos40.Text);
            pos[41] = Convert.ToInt16(edit_pos41.Text);
            pos[42] = Convert.ToInt16(edit_pos42.Text);
            pos[43] = Convert.ToInt16(edit_pos43.Text);
            pos[44] = Convert.ToInt16(edit_pos44.Text);
            pos[45] = Convert.ToInt16(edit_pos45.Text);
            pos[46] = Convert.ToInt16(edit_pos46.Text);
            pos[47] = Convert.ToInt16(edit_pos47.Text);
            pos[48] = Convert.ToInt16(edit_pos48.Text);
            pos[49] = Convert.ToInt16(edit_pos49.Text);
            pos[50] = Convert.ToInt16(edit_pos50.Text);
            pos[51] = Convert.ToInt16(edit_pos51.Text);
            pos[52] = Convert.ToInt16(edit_pos52.Text);
            pos[53] = Convert.ToInt16(edit_pos53.Text);
            pos[54] = Convert.ToInt16(edit_pos54.Text);
            pos[55] = Convert.ToInt16(edit_pos55.Text);
            pos[56] = Convert.ToInt16(edit_pos56.Text);
            pos[57] = Convert.ToInt16(edit_pos57.Text);
            pos[58] = Convert.ToInt16(edit_pos58.Text);
            pos[59] = Convert.ToInt16(edit_pos59.Text);
            pos[60] = Convert.ToInt16(edit_pos60.Text);
            pos[61] = Convert.ToInt16(edit_pos61.Text);
            pos[62] = Convert.ToInt16(edit_pos62.Text);
            pos[63] = Convert.ToInt16(edit_pos63.Text);
            pos[64] = Convert.ToInt16(edit_pos64.Text);
            pos[65] = Convert.ToInt16(edit_pos65.Text);
            pos[66] = Convert.ToInt16(edit_pos66.Text);
            pos[67] = Convert.ToInt16(edit_pos67.Text);
            pos[68] = Convert.ToInt16(edit_pos68.Text);
            pos[69] = Convert.ToInt16(edit_pos69.Text);
            pos[70] = Convert.ToInt16(edit_pos70.Text);
            pos[71] = Convert.ToInt16(edit_pos71.Text);
            pos[72] = Convert.ToInt16(edit_pos72.Text);
            pos[73] = Convert.ToInt16(edit_pos73.Text);
            pos[74] = Convert.ToInt16(edit_pos74.Text);
            pos[75] = Convert.ToInt16(edit_pos75.Text);
            pos[76] = Convert.ToInt16(edit_pos76.Text);
            pos[77] = Convert.ToInt16(edit_pos77.Text);
            pos[78] = Convert.ToInt16(edit_pos78.Text);
            pos[79] = Convert.ToInt16(edit_pos79.Text);
            pos[80] = Convert.ToInt16(edit_pos80.Text);
            pos[81] = Convert.ToInt16(edit_pos81.Text);
            pos[82] = Convert.ToInt16(edit_pos82.Text);
            pos[83] = Convert.ToInt16(edit_pos83.Text);
            pos[84] = Convert.ToInt16(edit_pos84.Text);
            pos[85] = Convert.ToInt16(edit_pos85.Text);
            pos[86] = Convert.ToInt16(edit_pos86.Text);
            pos[87] = Convert.ToInt16(edit_pos87.Text);
            pos[88] = Convert.ToInt16(edit_pos88.Text);
            pos[89] = Convert.ToInt16(edit_pos89.Text);
            pos[90] = Convert.ToInt16(edit_pos90.Text);
            pos[91] = Convert.ToInt16(edit_pos91.Text);
            pos[92] = Convert.ToInt16(edit_pos92.Text);
            pos[93] = Convert.ToInt16(edit_pos93.Text);
            pos[94] = Convert.ToInt16(edit_pos94.Text);
            pos[95] = Convert.ToInt16(edit_pos95.Text);
            pos[96] = Convert.ToInt16(edit_pos96.Text);
            pos[97] = Convert.ToInt16(edit_pos97.Text);
            pos[98] = Convert.ToInt16(edit_pos98.Text);
            pos[99] = Convert.ToInt16(edit_pos99.Text);
            pos[100] = Convert.ToInt16(edit_pos100.Text);
            pos[101] = Convert.ToInt16(edit_pos101.Text);
            pos[102] = Convert.ToInt16(edit_pos102.Text);
            pos[103] = Convert.ToInt16(edit_pos103.Text);
            pos[104] = Convert.ToInt16(edit_pos104.Text);
            pos[105] = Convert.ToInt16(edit_pos105.Text);
            pos[106] = Convert.ToInt16(edit_pos106.Text);
            pos[107] = Convert.ToInt16(edit_pos107.Text);
            pos[108] = Convert.ToInt16(edit_pos108.Text);
            pos[109] = Convert.ToInt16(edit_pos109.Text);
            pos[110] = Convert.ToInt16(edit_pos110.Text);
            pos[111] = Convert.ToInt16(edit_pos111.Text);
            pos[112] = Convert.ToInt16(edit_pos112.Text);
            pos[113] = Convert.ToInt16(edit_pos113.Text);
            pos[114] = Convert.ToInt16(edit_pos114.Text);
            pos[115] = Convert.ToInt16(edit_pos115.Text);
            pos[116] = Convert.ToInt16(edit_pos116.Text);
            pos[117] = Convert.ToInt16(edit_pos117.Text);
            pos[118] = Convert.ToInt16(edit_pos118.Text);
            pos[119] = Convert.ToInt16(edit_pos119.Text);
            pos[120] = Convert.ToInt16(edit_pos120.Text);
            pos[121] = Convert.ToInt16(edit_pos121.Text);
            pos[122] = Convert.ToInt16(edit_pos122.Text);
            pos[123] = Convert.ToInt16(edit_pos123.Text);
            pos[124] = Convert.ToInt16(edit_pos124.Text);
            pos[125] = Convert.ToInt16(edit_pos125.Text);
            pos[126] = Convert.ToInt16(edit_pos126.Text);
        }

        // Печать в отчет 
        public void Write2Report(string pos, string val)
        {
            DocumentPosition docpos = report.Document.Bookmarks[pos].Range.Start;
            report.Document.CaretPosition = docpos;
            report.ScrollToCaret();
            report.Document.InsertText(docpos, val);
        }

        // Преобразование числа месяца в строку
        private string Month2String(int m)
        {
            switch(m)
            {
                case 1: return "январь";
                case 2: return "февраль";
                case 3: return "март";
                case 4: return "апрель";
                case 5: return "май";
                case 6: return "июнь";
                case 7: return "июль";
                case 8: return "август";
                case 9: return "сентябрь";
                case 10: return "октябрь";
                case 11: return "ноябрь";
                case 12: return "декабрь";
                default: return "";
            }
        }

        // Преобразование даты в строку типа: yyyy-mm-dd
        private string DateToStr(DateTime date)
        {
            string res = date.Year.ToString() + "-" + date.Month.ToString() + "-";
            if (date.Day < 10) res += "0" + date.Day.ToString();
            else res += date.Day.ToString();

            return res;
        }

        // Заброс данных в шаблон отчета
        private void PutDataToReport()
        {
            int icode = Convert.ToInt16(current_dg_code);
                        
            report.LoadDocument(rep_filename.Text);

            mainTabControl.SelectedTabPageIndex = 5;

            Write2Report("month", Month2String(Convert.ToInt16(calc_month)));
            Write2Report("yy", calc_year);
            Write2Report("y1", calc_year[2].ToString());
            Write2Report("y2", calc_year[3].ToString());
            Write2Report("m1", calc_month[0].ToString());
            Write2Report("m2", calc_month[1].ToString());
            Write2Report("DG_name", current_dg_name);
            Write2Report("dgc1", current_dg_code[0].ToString());
            Write2Report("dgc2", current_dg_code[1].ToString());
            Write2Report("dgc3", current_dg_code[2].ToString());

            for (int i = 1; i < 127; i++) Write2Report( "item" + i.ToString(), oneK[icode,i].ToString() );

            Write2Report("ruk1_dol", viz_ruk1_dol.Text);
            Write2Report("ruk1_name", viz_ruk1_name.Text);
            Write2Report("ruk1_zvan", viz_ruk1_zvan.Text);

            Write2Report("ruk2_dol", "Начальник УРЛС");
            Write2Report("ruk2_name", viz_ruk2_name.Text);
            Write2Report("ruk2_zvan", viz_ruk2_zvan.Text);

            Write2Report("isp_name", viz_isp_name.Text);
            Write2Report("isp_zvan", viz_isp_zvan.Text);
        }

        // Сохранение данных ДГ в базу...
        private void SaveDGtoBase(string code, string date)
        {
            DateTime cdate = Convert.ToDateTime(date);
            string cmd = "";
            int form_code = Convert.ToInt16(code);

            // Проверяем есть ли такие данные в БД
            int res = DataProvider._getDataSQLs(RConn, String.Format("SELECT COUNT(code_dg) AS CNT FROM OneK_2012 WHERE (code_dg = {0}) AND (calc_date = '{1}-{2}-{3}')", code, cdate.Year, cdate.Month, cdate.Day));
            if (res > 0)
            {
                if (MessageBox.Show("Перезаписать существующие данные?", "Сохранение...", MessageBoxButtons.YesNo) != System.Windows.Forms.DialogResult.Yes) return;
                else
                {
                    cmd = "UPDATE OneK_2012 SET ";
                    for (int i = 1; i < 127; i++)
                    {
                        if (i == 1) cmd += "pos_" + i.ToString() + " = " + oneK[form_code,i];
                        else cmd += ", pos_" + i.ToString() + " = " + oneK[form_code, i];
                    }
                    cmd += String.Format(" WHERE (code_dg = {0}) AND (calc_date = '{1}-{2}-{3}')", code, cdate.Year, cdate.Month, cdate.Day);
                    res = DataProvider._insDataSQL(RConn, cmd);

                    //if (res == 1) MessageBox.Show("Данные успешно сохранены!");
                    //else MessageBox.Show("Ошибка при сохранении данных!");
                }
            }
            else
            {
                cmd = String.Format("INSERT INTO OneK_2012 (code_dg, calc_date, pos_1, pos_2, pos_3, pos_4, pos_5, pos_6, pos_7, pos_8, pos_9, pos_10, pos_11, pos_12, pos_13, " +
                    "pos_14, pos_15, pos_16, pos_17, pos_18, pos_19, pos_20, pos_21, pos_22, pos_23, pos_24, pos_25, pos_26, pos_27, pos_28, pos_29, pos_30, pos_31, pos_32, pos_33, " +
                    "pos_34, pos_35, pos_36, pos_37, pos_38, pos_39, pos_40, pos_41, pos_42, pos_43, pos_44, pos_45, pos_46, pos_47, pos_48, pos_49, pos_50, pos_51, pos_52, pos_53, " +
                    "pos_54, pos_55, pos_56, pos_57, pos_58, pos_59, pos_60, pos_61, pos_62, pos_63, pos_64, pos_65, pos_66, pos_67, pos_68, pos_69, pos_70, pos_71, pos_72, pos_73, " +
                    "pos_74, pos_75, pos_76, pos_77, pos_78, pos_79, pos_80, pos_81, pos_82, pos_83, pos_84, pos_85, pos_86, pos_87, pos_88, pos_89, pos_90, pos_91, pos_92, pos_93, " +
                    "pos_94, pos_95, pos_96, pos_97, pos_98, pos_99, pos_100, pos_101, pos_102, pos_103, pos_104, pos_105, pos_106, pos_107, pos_108, pos_109, pos_110, pos_111, pos_112, pos_113, " +
                    "pos_114, pos_115, pos_116, pos_117, pos_118, pos_119, pos_120, pos_121, pos_122, pos_123, pos_124, pos_125, pos_126) VALUES ('{0}','{1}-{2}-{3}',", code, cdate.Year, cdate.Month, cdate.Day);
                for (int i = 1; i < 127; i++)
                {
                    if (i == 1) cmd += oneK[form_code, i];
                    else cmd += ", " + oneK[form_code, i];
                }

                cmd += ")";

                res = DataProvider._insDataSQL(RConn, cmd);

               // if (res == 1) MessageBox.Show("Данные успешно сохранены!");
               // else MessageBox.Show("Ошибка при сохранении данных!");
            }
        }

        // Загрузка основной формы
        private void MainForm_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "pergroupDataSet.OneK_Pergroup". При необходимости она может быть перемещена или удалена.

            read_cfg();

            this.oneK_PergroupTableAdapter.Fill(this.pergroupDataSet.OneK_Pergroup);

            StatusBar.Panels[0].Text = String.Format("Загружено {0} должностных групп.", pergroupDataSet.Tables[0].Rows.Count);
            // Параметры соединения с БД
            RConn = r_connection_edit.Text;
            KConn = k_connection_edit.Text;
            // Меняем раскладку клавиатуры...
            //InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(new CultureInfo("ru-RU"));
            
            // Инициализируем расчетный массив
            InitAllOneK();

            // Проверяем доступные отчеты
            oneK_dates.Items.Clear();
            DataTable dt = DataProvider._getDataSQL(KConn, "SELECT distinct calc_date FROM OneK_2012 ORDER BY calc_date desc");
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string edate = Convert.ToDateTime(dt.Rows[i]["calc_date"]).ToShortDateString();
                    oneK_dates.Items.Add(edate);
                }
                oneK_dates.SelectedIndex = 0;
            }

            dt.Dispose();
            calc_date = calc_on_date.Value;
            calc_year = calc_date.Year.ToString();
            calc_date_label.Text = string.Format("Текущая дата расчета: {0}",calc_date.ToShortDateString());

            // Загружаем шаблон отчета
            if (File.Exists(rep_filename.Text))
            {
                report.LoadDocument(rep_filename.Text);
            }
            else MessageBox.Show("Файл шаблона отчета:" + rep_filename.Text + " не найден!\nЗайдите во вкладку 'Настройки и Подготовка', далее выберите путь к файлу шаблона, далее нажмите 'Загрузить шаблон отчета'\nНе забудьте сохранить настройки!");

            filename_exp_edit.Text = String.Format("OneK_{0}.txt", System.DateTime.Now.Year);
            kval_check.Checked = false;
            smen_check.Checked = false;
        }

        // Расчет одной выбранной ДГ
        private void calc_one_formButton_Click(object sender, EventArgs e)
        {
            CalcDG(current_dg_code, current_sql_text);
            FillEditPanel(current_dg_code, 1);
        }

        // Выход
        private void exit_Button_Click(object sender, EventArgs e)
        {
            // Прибираемся
            GC.Collect();
            Close();
        }

        // Обработка нажатия на клавишу мыши в таблице должностных групп...
        private void grid_MouseClick(object sender, MouseEventArgs e)
        {
            GridArea area = grid.HitTest(e.X, e.Y);

            if (area == GridArea.Cell)
            {
                GridEXRow row = grid.CurrentRow;

                TitleDG.Text = "Текущая должностная группа: " + row.Cells["KEY_OTCH"].Text;
                current_dg_code = row.Cells["KEY_OTCH"].Text;
                current_sql_text = row.Cells["TEXT_QRY"].Text;
                current_dg_name = row.Cells["NAME_FORM"].Text;
                current_dg_lider = (bool)row.Cells["LIDER"].Value;
                current_dg_sluzb = row.Cells["SLUZBA_SQL"].Text;
                FillEditPanel(current_dg_code, 0);
            }
        }
    
        // Сохранение текущей должностной группы в БД
        private void save_DG_button_Click(object sender, EventArgs e)
        {
            report.SaveDocumentAs();
        }

        // Проверка логики текущей должностной группы
        private void check_logic_dg_button_Click(object sender, EventArgs e)
        {
            if (CheckDG(current_dg_code) == false)
            {
                mainTabControl.SelectedTabPageIndex = 4;
                MessageBox.Show("Форма проверена, есть ошибки...");
            }
            else
            {
                MessageBox.Show("Форма проверена, ошибок не обнаружено!");
            }
        }

        // Расчет выделенной должностной группы
        private void расчитатьДолжностнуюГруппуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StatusBar.Panels[0].Text = String.Format("[{0}] - {1}", current_dg_code, current_dg_name);
            CalcDG(current_dg_code, current_sql_text);
            FillEditPanel(current_dg_code,1);
            RefreshPosFromEditPanel();  //?
            //PutDataToReport();
        }

        // Редактирование текущей ДГ
        private void редактироватьДанныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (current_dg_code != "")
            {
                mainTabControl.SelectedTabPageIndex = 1;
                FillEditPanel(current_dg_code, 1);
                RefreshPosFromEditPanel(); //?
            }
        }

        // Расчет всех должностных групп
        private void calc_all_button_Click(object sender, EventArgs e)
        {
            // Проверка начальных условий и настроек...
            if (kval_check.Checked)
            {
                if (smen_check.Checked)
                {
                    // Если все готово
                    DataRowCollection rc = pergroupDataSet.Tables[0].Rows;

                    //StatusBar.Panels[1].ProgressBarMaxValue = rc.Count;
                    // Проходим все ДГ
                    for (int i = 0; i < rc.Count; i++)
                    {
                        string code = rc[i]["KEY_OTCH"].ToString();
                        // Если не "хитрые"
                        if (code != "300" && code != "301" && code != "900" && code != "901")
                        {
                            StatusBar.Panels[0].Text = String.Format("[{0}] - {1}", code, rc[i]["NAME_FORM"].ToString());
                            // Вычисляем ДГ
                            current_dg_code = rc[i]["KEY_OTCH"].ToString();
                            current_sql_text = rc[i]["TEXT_QRY"].ToString();
                            current_dg_name = rc[i]["NAME_FORM"].ToString();
                            current_dg_lider = (bool)rc[i]["LIDER"];
                            current_dg_sluzb = rc[i]["SLUZBA_SQL"].ToString();
                            current_dg_dolz = rc[i]["DOLZNOST_SQL"].ToString();
                            CalcDG(code, rc[i]["TEXT_QRY"].ToString());
                            //StatusBar.Panels[1].ProgressBarValue++;
                        }
                        // Не засыпаем...
                        Application.DoEvents();
                    }

                    //StatusBar.Panels[1].ProgressBarValue = 0;
                    StatusBar.Panels[0].Text = "Расчет окончен, не забудьте сохраниться...";
                    save_all_express_button.Enabled = true;

                }
                else MessageBox.Show("Не выполнена подготовка массива ключей для сменяемости!");
            }
            else MessageBox.Show("Не выполнена предварительная проверка соответствия квалификационным требованиям!");
        }

        // кнопка сохранения всего отчета
        private void save_all_express_button_Click(object sender, EventArgs e)
        {
            SaveAllOneKToBase();
        }

        // Проверка всего отчета
        private void check_logic_all_button_Click(object sender, EventArgs e)
        {
            if (!CheckInterDG())
            {
                mainTabControl.SelectedTabPageIndex = 4;
                MessageBox.Show("Отчет проверен, есть ошибки...");
            }
            else MessageBox.Show("Отчет проверен, ошибок НЕТ !!!");
        }

        // Загрузка данных всего отчета
        private void load_all_express_button_Click(object sender, EventArgs e)
        {
            DateTime edate = Convert.ToDateTime(oneK_dates.SelectedItem);
            LoadAllOneKFromBase(edate); 
        }

        // Загрузка выбранного отчета из БД
        private void load_express_date_button_Click(object sender, EventArgs e)
        {
            DateTime edate = Convert.ToDateTime(oneK_dates.SelectedItem);
            LoadAllOneKFromBase(edate);
        }

        // Сохранение текущей ДГ в окне редактирования [I]
        private void save_dg_inedit1_button_Click(object sender, EventArgs e)
        {
            RefreshPosFromEditPanel();
            SavePos2OneK(current_dg_code);
            SaveDGtoBase(current_dg_code, calc_date.ToShortDateString());
            MessageBox.Show("Данные успешно сохранены!");
        }

        // Проверка ДГ в форме редактиования [I]
        private void check_dg_inedit1_button_Click(object sender, EventArgs e)
        {
            if (!CheckDG(current_dg_code)) mainTabControl.SelectedTabPageIndex = 4;
        }

        // Печать ДГ из формы редактирования [I]
        private void print_dg_inedit_button_Click(object sender, EventArgs e)
        {
            PutDataToReport();
        }

        // Сохранение текущей ДГ в окне редактирования [III]
        private void save_dg_inedit3_button_Click(object sender, EventArgs e)
        {
            RefreshPosFromEditPanel();
            SavePos2OneK(current_dg_code);
            SaveDGtoBase(current_dg_code, calc_date.ToShortDateString());
            MessageBox.Show("Данные успешно сохранены!");
        }

        // Проверка ДГ в форме редактиования [III]
        private void check_dg_inedit3_button_Click(object sender, EventArgs e)
        {
            if (!CheckDG(current_dg_code)) mainTabControl.SelectedTabPageIndex = 4;
        }

        // Печать ДГ из формы редактиования [III]
        private void print_dg_inedit3_button_Click(object sender, EventArgs e)
        {
            PutDataToReport();
        }
        
        // Сложение 2-х форм: [src] прибавляют к [dst]
        private void SumDG(int src, int dst)
        {
            for (int i = 1; i < 127; i++)
                oneK[dst, i] += oneK[src, i];
        }

        // Добавить в текущую должностную группу
        private void sum_dg_button_Click(object sender, EventArgs e)
        {
            int dst_code = Convert.ToInt16(current_dg_code);
            
            if (edit_plus1.Text != "" && MessageBox.Show(String.Format("Добавить к данным [{0}] данные [{1}] ?", dst_code, edit_plus1.Text),"Сложение ДГ",MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                int src_code = Convert.ToInt16(edit_plus1.Text);
                SumDG(src_code, dst_code);
            }
            if (edit_plus2.Text != ""  && MessageBox.Show(String.Format("Добавить к данным [{0}] данные [{1}] ?", dst_code, edit_plus2.Text),"Сложение ДГ",MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                int src_code = Convert.ToInt16(edit_plus2.Text);
                SumDG(src_code, dst_code);
            }
            if (edit_plus3.Text != ""  && MessageBox.Show(String.Format("Добавить к данным [{0}] данные [{1}] ?", dst_code, edit_plus3.Text),"Сложение ДГ",MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                int src_code = Convert.ToInt16(edit_plus3.Text);
                SumDG(src_code, dst_code);
            }
            
            FillEditPanel(current_dg_code, 1);
        }

        // Очистка отчета и удаление его из БД
        private void clear_all_express_button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите обнулить весь отчет?", "ВНИМАНИЕ", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                InitAllOneK();
                int res = DataProvider._insDataSQL(RConn, String.Format("DELETE FROM OneK_2012 WHERE calc_date = '{0}'", calc_date.ToShortDateString()));
                MessageBox.Show("Отчет текущией даты расчета удален!");
            }
        }

        // Экспорт данных отчета в формат Статистика-К
        private void export_button_Click(object sender, EventArgs e)
        {
            // Имя файла отчета
            saveExportFileDialog.FileName = String.Format("845{0}{1}.rep", calc_year.Substring(2, 2), calc_month); 
            // Диалог сохранения
            System.Windows.Forms.DialogResult res = saveExportFileDialog.ShowDialog();
            if (res != System.Windows.Forms.DialogResult.Cancel && saveExportFileDialog.FileName != "")
            {
                System.IO.StreamWriter rep = new StreamWriter(saveExportFileDialog.FileName);

                // Заголовок файла
                rep.WriteLine(String.Format("/h {0}{1}845 \\h",calc_year.Substring(2, 2), calc_month));
                // Начало первого раздела
                rep.WriteLine("/d 01 \\d");
                // Все формы в разделе...
                for(int i = 0; i < CountDG+1; i++)
                {
                // Замечание по справочнику ПО "СТАТИСТКА-К" st_spr.mdb, присланному из ДГСК МВД России в регионы 29.11.2016
                // В таблице справочника spr_Headers разработчик допущена досадная ошибка:
                // в строке с ID_Header=1558 в поле KG_code помещено значение 250 (ошибочно взятое, очевидно из поля KG_ID таблицы spr_CadrGroups), вместо значения 410 - видимый код кадровой группы
                // в строке с ID_Header=1560 в поле KG_code помещено значение 251 вместо 420
                // в строке с ID_Header=1562 в поле KG_code помещено значение 252 вместо 440
                // в строке с ID_Header=1564 в поле KG_code помещено значение 253 вместо 450
                // В результате - при загрузке данных 1-К из rep-файла новые кадровые группы с номерами в заголовке 410, 420, 440 и 450 не загружается. То есть, процесс загрузки от имени администратора проходит нормально А при попытке пользователя открыть загруженный отчёт программа выдаёт последовательно две ошибки:
                // 1) "Обнаружено несоответствие идентификаторов ячеек для формы 1-К за декабрь 2016!"
                // 2) "Невозможно загрузить данные для формы 1-К за декабрь 2016!"
                // Однако, если в rep-файле в строке заголовков кадровых групп указать вместо этих номеров 250, 251, 252, 253, соответственно - с загруженным отчётом можно нормально работать.
                    
                // ----- Исправлено в СТАТИСТИКЕ-К 28.12.2022
                // if (repDG[i] == "410") rep.Write("250");
                // else if (repDG[i] == "420") rep.Write("251");
                // else if (repDG[i] == "440") rep.Write("252");
                // else if (repDG[i] == "450") rep.Write("253");
                rep.Write(repDG[i]);
                if ( i < CountDG) rep.Write(" ");
                }
                rep.WriteLine("");
                // Пишем данные...
                for (int i = 1; i < 127; i++)
                {
                    for(int j = 0; j < CountDG+1; j++)
                    {
                        int code = Convert.ToInt16(repDG[j]);
                        rep.Write(oneK[code, i]);
                        if (j < CountDG) rep.Write(" ");
                    }
                    rep.WriteLine("");
                }
                // Закрываем                
                rep.WriteLine("\\d");
                rep.WriteLine("\\End");
                rep.Flush();
                rep.Close();
                MessageBox.Show("Экспорт завершен!");
            }

        }
        
        #region Переходы в форме редактирования..
        private void edit_pos1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos2.Focus();
        }

        private void edit_pos2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                CheckPosition();
                edit_pos3.Focus();
            }
        }

        private void edit_pos3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos4.Focus();
        }

        private void edit_pos4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos5.Focus();
        }

        private void edit_pos5_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos6.Focus();
        }

        // Проверка сумм по: образованию, возрасту
        private void CheckPosition()
        {
            //int sum1 = Convert.ToInt16(edit_pos7.Text) + Convert.ToInt16(edit_pos8.Text) + Convert.ToInt16(edit_pos9.Text) + Convert.ToInt16(edit_pos10.Text);
            //if (Convert.ToInt16(edit_pos6.Text) != sum1) obr1_sum.Text = String.Format("Сумма высшего:{0}", sum1);
            //else obr1_sum.Text = "";
            //int sum2 = Convert.ToInt16(edit_pos15.Text) + Convert.ToInt16(edit_pos16.Text) + Convert.ToInt16(edit_pos17.Text) + Convert.ToInt16(edit_pos18.Text);
            //if (Convert.ToInt16(edit_pos14.Text) != sum2) obr2_sum.Text = String.Format("Сумма ср.проф.:{0}", sum2);
            //else obr2_sum.Text = "";
            int sum = Convert.ToInt16(edit_pos6.Text) + Convert.ToInt16(edit_pos14.Text) +Convert.ToInt16(edit_pos25.Text);
            if (sum != Convert.ToInt16(edit_pos2.Text)) obr_sum.Text = String.Format("Сумма всех образований:{0}", sum);
            else obr_sum.Text = "";
        }

        private void edit_pos6_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                CheckPosition();
                edit_pos7.Focus();
            }
        }

        private void edit_pos7_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                CheckPosition();
                edit_pos8.Focus();
            }
        }

        private void edit_pos8_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                CheckPosition();
                edit_pos9.Focus();
            }
        }

        private void edit_pos9_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                CheckPosition();
                edit_pos10.Focus();
            }
        }

        private void edit_pos10_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                CheckPosition();
                edit_pos11.Focus();
            }
        }

        private void edit_pos11_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos12.Focus();
        }

        private void edit_pos12_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos13.Focus();
        }

        private void edit_pos13_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos14.Focus();
        }

        private void edit_pos14_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos15.Focus();
        }

        private void edit_pos15_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos16.Focus();
        }

        private void edit_pos16_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos17.Focus();
        }

        private void edit_pos17_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos18.Focus();
        }

        private void edit_pos18_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos19.Focus();
        }

        private void edit_pos19_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos20.Focus();
        }

        private void edit_pos20_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos21.Focus();
        }

        private void edit_pos21_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos22.Focus();
        }

        private void edit_pos22_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos23.Focus();
        }

        private void edit_pos23_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos24.Focus();
        }

        private void edit_pos24_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos25.Focus();
        }

        private void edit_pos25_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                CheckPosition();
                edit_pos26.Focus();
            }
        }

        private void edit_pos26_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos27.Focus();
        }

        private void edit_pos27_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos28.Focus();
        }

        private void edit_pos28_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos29.Focus();
        }

        private void edit_pos29_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos30.Focus();
        }

        private void edit_pos30_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos31.Focus();
        }

        private void edit_pos31_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos32.Focus();
        }

        private void edit_pos32_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos33.Focus();
        }

        private void edit_pos33_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos34.Focus();
        }

        private void edit_pos34_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos35.Focus();
        }

        private void edit_pos35_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos36.Focus();
        }

        private void edit_pos36_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos37.Focus();
        }

        private void edit_pos37_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos38.Focus();
        }

        private void edit_pos38_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos39.Focus();
        }

        private void edit_pos39_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos40.Focus();
        }

        private void edit_pos40_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos41.Focus();
        }

        private void edit_pos41_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos42.Focus();
        }
                
        private void edit_pos42_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos43.Focus();
        }

        private void edit_pos43_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos44.Focus();
        }

        private void edit_pos44_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos45.Focus();
        }

        private void edit_pos45_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos46.Focus();
        }

        private void edit_pos46_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos47.Focus();
        }

        private void edit_pos47_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos48.Focus();
        }

        private void edit_pos48_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos49.Focus();
        }

        private void edit_pos49_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos50.Focus();
        }

        private void edit_pos50_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos51.Focus();
        }

        private void edit_pos51_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos52.Focus();
        }

        private void edit_pos52_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos53.Focus();
        }

        private void edit_pos53_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos54.Focus();
        }
 
        private void edit_pos54_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos55.Focus();
        }

        private void edit_pos55_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos56.Focus();
        }

        private void edit_pos56_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos57.Focus();
        }

        private void edit_pos57_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos58.Focus();
        }

        private void edit_pos58_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos59.Focus();
        }

        private void edit_pos59_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos60.Focus();
        }

        private void edit_pos60_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos61.Focus();
        }
            
        
        //Control ctrl = (Control) sender;
        //int num = Convert.ToInt16(ctrl.Name.TrimStart("edit_pos".ToCharArray()));
        //string name = String.Format("edit_pos{0}", ++num);
        //this.Controls[name].Focus();
        private void edit_pos61_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                mainTabControl.SelectedTabPageIndex++;
                edit_pos62.Focus();
            }
        }

        private void edit_pos62_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos63.Focus();
        }

        private void edit_pos63_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos64.Focus();
        }

        private void edit_pos64_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos65.Focus();
        }

        private void edit_pos65_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos66.Focus();
        }

        private void edit_pos66_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos67.Focus();
        }

        private void edit_pos67_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos68.Focus();
        }

        private void edit_pos68_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos69.Focus();
        }

        private void edit_pos69_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos70.Focus();
        }

        private void edit_pos70_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos71.Focus();
        }

        private void edit_pos71_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos72.Focus();
        }

        private void edit_pos72_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos73.Focus();
        }

        private void edit_pos73_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos74.Focus();
        }

        private void edit_pos74_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos75.Focus();
        }

        private void edit_pos75_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos76.Focus();
        }

        private void edit_pos76_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos77.Focus();
        }

        private void edit_pos77_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos78.Focus();
        }

        private void edit_pos78_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos79.Focus();
        }

        private void edit_pos79_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos80.Focus();
        }

        private void edit_pos80_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos81.Focus();
        }

        private void edit_pos81_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos82.Focus();
        }

        private void edit_pos82_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos83.Focus();
        }

        private void edit_pos83_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos84.Focus();
        }

        private void edit_pos84_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos85.Focus();
        }

        private void edit_pos85_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos86.Focus();
        }

        private void edit_pos86_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos87.Focus();
        }

        private void edit_pos87_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos88.Focus();
        }

        private void edit_pos88_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos89.Focus();
        }

        private void edit_pos89_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos90.Focus();
        }

        private void edit_pos90_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos91.Focus();
        }

        private void edit_pos91_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos92.Focus();
        }

        private void edit_pos92_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos93.Focus();
        }

        private void edit_pos93_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos94.Focus();
        }

        private void edit_pos94_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos95.Focus();
        }

        private void edit_pos95_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos96.Focus();
        }

        private void edit_pos96_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos97.Focus();
        }

        private void edit_pos97_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos98.Focus();
        }

        private void edit_pos98_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos99.Focus();
        }

        private void edit_pos99_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos100.Focus();
        }

        private void edit_pos100_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos101.Focus();
        }

        private void edit_pos101_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos102.Focus();
        }

        private void edit_pos102_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos103.Focus();
        }

        private void edit_pos103_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos104.Focus();
        }

        private void edit_pos104_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                mainTabControl.SelectedTabPageIndex++;
                edit_pos105.Focus();
            }
        }

        private void edit_pos105_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos106.Focus();
        }

        private void edit_pos106_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos107.Focus();
        }

        private void edit_pos107_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos108.Focus();
        }

        private void edit_pos108_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos109.Focus();
        }

        private void edit_pos109_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos110.Focus();
        }

        private void edit_pos110_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos111.Focus();
        }

        private void edit_pos111_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos112.Focus();
        }

        private void edit_pos112_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos113.Focus();
        }

        private void edit_pos113_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos114.Focus();
        }

        private void edit_pos114_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos115.Focus();
        }

        private void edit_pos115_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos116.Focus();
        }

        private void edit_pos116_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos117.Focus();
        }

        private void edit_pos117_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos118.Focus();
        }

        private void edit_pos118_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos119.Focus();
        }

        private void edit_pos119_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos120.Focus();
        }

        private void edit_pos120_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos121.Focus();
        }

        private void edit_pos121_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos122.Focus();
        }

        private void edit_pos122_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos123.Focus();
        }

        private void edit_pos123_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos124.Focus();
        }

        private void edit_pos124_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos125.Focus();
        }

        private void edit_pos125_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) edit_pos126.Focus();
        }

        private void edit_pos126_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                mainTabControl.SelectedTabPageIndex = 1;
                edit_pos1.Focus();
            }
        }

        #endregion

        // Обновление сведений по соответствию образования
        private void update_kvalify_button_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("ВНИМАНИЕ !\n Будет проведена проверка соответствия имеющегося образования по замещаемым должностям у всего среднего и старшего начсостава.\n Изменения вносятся в БД автоматически!\n (Процесс может занять некоторое время...)", "Проверка...", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                DataTable dt = DataProvider._getDataSQL(KConn,"SELECT KEY_1, KVAL, SLUZBA, DOLZNOST, POTOLOK, OBRAZ_DOL1, OBRAZ_DOL2, KOD_OBRAZ, RTRIM(OBRAZ_LIC1) AS OBRAZ_LIC1, OBRAZ_LIC2, OBRAZ_LIC3 FROM AAQQ WHERE DOLZNOST < '500000' AND FAMILIYA <> '' ORDER BY KEY_1");
                DataRowCollection rc = dt.Rows;
                dt.Dispose();
                StatusBar.Panels[1].ProgressBarMinValue = 0;
                StatusBar.Panels[1].ProgressBarMaxValue = rc.Count;
                StatusBar.Panels[1].ProgressBarValue = 0;
                int kvalYES = 0;
                int kvalNO = 0;

                for (int i = 0; i < rc.Count; i++)
                {
                    int kval = 0; // Квалификация: 0 - не соответствует, 1 - соответствует
                    int sostav = Convert.ToInt16(rc[i]["POTOLOK"]); // Предельное звание
                    int vid = Convert.ToInt16(rc[i]["OBRAZ_LIC2"]); // Высшее, среднее-спец., среднее
                    string type = rc[i]["OBRAZ_LIC1"].ToString();   // Профиль: юридическое,...
                    int sl = Convert.ToInt16(rc[i]["SLUZBA"]);      // Служба
                    string id = rc[i]["KEY_1"].ToString();          // Ключ

                    vysl v = Vysluga.Find(x => x.key == id);
                                        
                    StatusBar.Panels[0].Text = String.Format("Смотрим: ключ - {0}", id);

                    // старший начальствующий состав...
                    if ((sostav >= 51 && sostav <= 55) || (sostav >= 71 && sostav <= 75) || (sostav >= 111 && sostav <= 114))
                    {
                        // Если не высшее - не соответствует
                        if (vid != 10) kval = 0;
                        else
                            // Проверка обязательности юридического (по приказу № 541, с 2016 года)
                            // 2.1. Обучение в образовательной организации высшего образования, имеющей государственную аккредитацию, по направлениям подготовки (специальностям) 40.03.01 Юриспруденция,
                            //      40.04.01 Юриспруденция, 40.05.01 Правовое обеспечение национальной безопасности, 40.05.02 Правоохранительная деятельность при наличии высшего образования,
                            //      не являющегося юридическим, и стажа службы в органах внутренних дел Российской Федерации не менее одного года.
                            // 2.2. Наличие среднего профессионального юридического образования, подтвержденного соответствующим документом об образовании и о квалификации или иным документом,
                            //      приравненным к нему в соответствии с законодательством Российской Федерации об образовании, и высшего образования, не являющегося юридическим,
                            //      а также стажа службы в органах внутренних дел не менее одного года.
                            // 2.3. Наличие ученой степени по юридической специальности, высшего образования, не являющегося юридическим, и опыта работы по юридической специальности не менее одного года.

                            // Если необходимо юридическое
                            if (sl == 10 || sl == 55 || sl == 41 || sl == 86)
                            {
                                // Если нет юридического
                                if (type != "Юридическое")
                                {
                                    // Есть ли обучение по вышке
                                    int ino = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM LEARN WHERE VID = '10' AND STATUS IN (2,3) AND TYPE LIKE 'Юрид%' AND KEY_1 = {0}", id));
                                    if (ino > 0)
                                    {
                                        // Проверяем есть ли иное высшее
                                        ino = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM LEARN WHERE STATUS = 1 AND VID = '10' AND TYPE NOT LIKE 'Юрид%' AND KEY_1 = {0}", id));
                                        if (ino > 0) // Если есть - соответств.
                                        {
                                            if ( v.years_c >= 1 ) kval = 1;
                                        }
                                        else kval = 0;
                                    }
                                    else // Проверяем есть ли среднее профессиональное юридическое
                                    {
                                        ino = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM LEARN WHERE STATUS = 1 AND VID = '20' AND TYPE LIKE 'Юрид%' AND KEY_1 = {0}", id));
                                        if (ino > 0) // Если есть ищем иное высшее
                                        {
                                            // Проверяем есть ли иное высшее
                                            ino = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM LEARN WHERE STATUS = 1 AND VID = '10' AND TYPE NOT LIKE 'Юрид%' AND KEY_1 = {0}", id));
                                            if (ino > 0) // Если есть - соответств.
                                            {
                                                if ( v.years_c >= 1 ) kval = 1;
                                            }
                                            else kval = 0;
                                        }
                                        else kval = 0;
                                    }
                                }
                            }
                            else kval = 1;
                    }
                    // средний начальствующий состав...
                    if ((sostav >= 47 && sostav <= 50) || (sostav >= 67 && sostav <= 70) || (sostav >= 107 && sostav <= 110))
                    {
                        // Проверка обязательности юридического
                        // Если необходимо юридическое
                        if (sl == 10 || sl == 55 || sl == 41 || sl == 86)
                        {
                            // Если нет юридического
                            if (type != "Юридическое")
                            {

                                // Есть ли обучение по вышке
                                int ino = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM LEARN WHERE VID = '10' AND STATUS IN (2,3) AND TYPE LIKE 'Юрид%' AND KEY_1 = {0}", id));
                                if (ino > 0)
                                {
                                    // Проверяем есть ли иное высшее
                                    ino = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM LEARN WHERE STATUS = 1 AND VID = '10' AND TYPE NOT LIKE 'Юрид%' AND KEY_1 = {0}", id));
                                    if (ino > 0) // Если есть - соответств.
                                    {
                                        if (v.years_c >= 1) kval = 1;
                                    }
                                    else kval = 0;
                                }
                                else // Проверяем есть ли среднее профессиональное юридическое
                                {
                                    ino = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM LEARN WHERE STATUS = 1 AND VID = '20' AND TYPE LIKE 'Юрид%' AND KEY_1 = {0}", id));
                                    if (ino > 0) // Если есть ищем иное высшее
                                    {
                                        // Проверяем есть ли иное высшее
                                        ino = DataProvider._getDataSQLs(KConn, String.Format("SELECT COUNT(KEY_1) FROM LEARN WHERE STATUS = 1 AND VID = '10' AND TYPE NOT LIKE 'Юрид%' AND KEY_1 = {0}", id));
                                        if (ino > 0) // Если есть - соответств.
                                        {
                                            if (v.years_c >= 1) kval = 1;
                                        }
                                        else kval = 0;
                                    }
                                    else kval = 0;
                                }
                            }
                        }
                        // Если юридическое не обязательно
                        else
                        {
                           // Если есть высшее, неп.высшее, сре.спец - соответствует
                            if (vid == 10 || vid == 50 || vid == 20) kval = 1;
                            else // Если среднее - несоответств.
                                if (vid == 30) kval = 0;
                        }
                    }

                    // Обновляем данные
                    int res = DataProvider._insDataSQL(KConn, String.Format("UPDATE AAQQ SET KVAL={0} WHERE KEY_1 = {1}", kval, id));

                    if (kval == 0) kvalNO++;
                    else kvalYES++;

                    StatusBar.Panels[1].ProgressBarValue++;
                    StatusBar.Panels[1].Text = i.ToString();
                }
                StatusBar.Panels[0].Text = "";
                StatusBar.Panels[1].ProgressBarValue = 0;
                MessageBox.Show(String.Format("Соответствие образования проверено, данные обновлены!\nСотрудников соответствует квалификационным требованиям - {0}, не соответствует - {1}", kvalYES, kvalNO));
                kval_check.Checked = true;
            }
            
        }

        // Подготовка массивов ключей для расчета сменяемости служб...
        private void prepare_poslkey_button_Click(object sender, EventArgs e)
        {
            // Выбираем ключи всех у кого в тек.году сменилась должность (DATA_VDOLZ)
            DataTable dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM AAQQ WHERE DOLZNOST < '800000' AND FAMILIYA <> '' AND DATA_VDOLZ >= {0} AND DATA_VDOLZ <= {1}", Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
            // Проверяем была ли предыдущая служба в послужном другая...
            StatusBar.Panels[1].ProgressBarMinValue = 0;
            StatusBar.Panels[1].ProgressBarValue = 0;
            StatusBar.Panels[1].ProgressBarMaxValue = dt.Rows.Count;
            PoslKeys1.Clear();
            PoslKeys2.Clear();
            smen item = new smen();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                //STATUS IN ('1100','1200','1300','1400','1500','1700','1701','3000','9000','9002','9003','9004','9006')
                
                // Выбираем послужной, смотрим 2-е последние записи...
                DataTable dt1 = DataProvider._getDataSQL(KConn, String.Format("SELECT SLUZBA, DATA_OT FROM POSL_SPI WHERE STATUS NOT IN ('1000','6000', '4000', '4001', '7000', '5000', '6001', '2000') AND KEY_POSL = {0} ORDER BY DATA_OT DESC", dt.Rows[i]["KEY_1"].ToString()));
                if (dt1.Rows.Count > 1)
                {
                   if (dt1.Rows[0]["SLUZBA"].ToString() != dt1.Rows[1]["SLUZBA"].ToString()) 
                   {
                       PoslKeys1.Add(Convert.ToInt16(dt.Rows[i]["KEY_1"]));
                       item.key = Convert.ToInt16(dt.Rows[i]["KEY_1"]);
                       item.sluzba_prev = Convert.ToInt16(dt1.Rows[1]["SLUZBA"]);
                       PoslKeys2.Add(item);
                   }
                }
                dt1.Dispose();
                StatusBar.Panels[0].Text = String.Format("Смотрим: ключ {0}",item.key);
                StatusBar.Panels[1].ProgressBarValue++;
                Application.DoEvents();
            }
            dt.Dispose();
            StatusBar.Panels[0].Text = "";
            StatusBar.Panels[1].ProgressBarValue = 0;
            MessageBox.Show(String.Format("Готово!"));
            smen_check.Checked = true;
        }

        // Сохранение массивов ключей в файлы...
        private void save_poslkeys_button_Click(object sender, EventArgs e)
        {
            if (PoslKeys1.Count != 0 && PoslKeys2.Count != 0)
            {
                System.IO.StreamWriter pk1 = new StreamWriter("poslkeys1.txt");
                System.IO.StreamWriter pk2 = new StreamWriter("poslkeys2.txt");

                foreach (int key in PoslKeys1) pk1.WriteLine(String.Format("{0}", key));
                pk1.Flush();
                pk1.Close();
                foreach (smen item in PoslKeys2)
                {
                    pk2.WriteLine(String.Format("{0}", item.key));
                    pk2.WriteLine(String.Format("{0}", item.sluzba_prev));
                }
                pk2.Flush();
                pk2.Close();
                MessageBox.Show("Данные сохранены!");
            }
        }

        // Загрузка данных в массивы ключей...
        private void load_poslkeys_button_Click(object sender, EventArgs e)
        {
            if (PoslKeys1.Count != 0 && PoslKeys2.Count != 0)
            {
                PoslKeys1.Clear();
                PoslKeys2.Clear();
            }
                System.IO.StreamReader pk1 = new StreamReader("poslkeys1.txt");
                System.IO.StreamReader pk2 = new StreamReader("poslkeys2.txt");
                            
                do
                {
                  PoslKeys1.Add(Convert.ToInt16(pk1.ReadLine()));
                } while (!pk1.EndOfStream);
                pk1.Close();

                smen item = new smen();
                do
                {
                    item.key = Convert.ToInt16(pk2.ReadLine());
                    item.sluzba_prev = Convert.ToInt16(pk2.ReadLine());
                    PoslKeys2.Add(item);
                } while (!pk2.EndOfStream);
                pk2.Close();

                MessageBox.Show(String.Format("Данные загружены!\nPoslKeys1 - {0}\nPoslKeys2 - {1}", PoslKeys1.Count, PoslKeys2.Count));
                smen_check.Checked = true;
        }

        // Экспорт данных отчета в текстовый файл
        private void exp_txt_button_Click(object sender, EventArgs e)
        {
            if (filename_exp_edit.Text != "")
            {
                System.IO.StreamWriter rep = new StreamWriter(filename_exp_edit.Text);

                for (int i = 0; i < 1000; i++)
                {
                    for (int j = 0; j < 127; j++) rep.WriteLine(oneK[i, j]);
                }

                rep.Flush();
                rep.Close();

                MessageBox.Show(String.Format("Отчет сохранен:{0}!",filename_exp_edit.Text));
            }
            MessageBox.Show("Неверное имя файла!");
        }

        // Импорт отчета из текстового файла...
        private void imp_txt_button_Click(object sender, EventArgs e)
        {
            if (filename_imp_edit.Text != "")
            {
                System.IO.StreamReader rep = new StreamReader(filename_imp_edit.Text);

                for (int i = 0; i < 1000; i++)
                {
                    for (int j = 0; j < 127; j++) oneK[i, j] = Convert.ToInt16(rep.ReadLine());
                }

                rep.Close();

                MessageBox.Show(String.Format("Отчет импортирован из:{0}", filename_imp_edit.Text));
            }
            MessageBox.Show("Неправильное имя файла!");
        }

        // Проверка стажа по ключу...
        private void check_stag_button_Click(object sender, EventArgs e)
        {
            if (key_debbug_edit.Text != "")
            {
                TPeriod p = CalcStage(key_debbug_edit.Text, check_stag_calendar.Checked);
                debug.Text += p.ToShortString() + "\n";
            }
        }

        // Сохранение отчета по ДГ в файл Word
        private void save_report_to_file_button_Click(object sender, EventArgs e)
        {
            report.SaveDocumentAs();
        }

        // Выбор файла для импорта текстового отчета
        private void select_imp_txt_button_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() != System.Windows.Forms.DialogResult.Cancel)
            {
                filename_imp_edit.Text = openFileDialog.FileName;
            }
        }

        // Очистка окна проверки
        private void clear_debbug_button_Click(object sender, EventArgs e)
        {
            debug.Clear();
        }

        // Проверка выделенной должностной группы
        private void проверитьДолжностнуюГруппуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (current_dg_code != "")
            {
                check_logic_dg_button_Click(sender, e);
            }
        }

        private void OnePosMenuStrip_Opening(object sender, CancelEventArgs e)
        {

        }

        // Расчет текущей позиции для данной ДГ
        private void расчетТекущейПозицииДляДаннойКадровойГруппыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (current_dg_code != "")
            {
                #region // Прибыло сотрудников из других служб
                if (current_dg_code != "010" && current_dg_code != "013")
                {
                    PrStat("Прибыло сотрудников из других служб");
                    //pos[59] = DataProvider._getDataODBCs(String.Format("SELECT COUNT(KEY_1) FROM AAQQ {0} AND FAMILIYA <> '' AND VREMI_V_SL >= {1} AND VREMI_V_SL <= {2} AND KEY_1 IN (SELECT KEY_POSL FROM POSL_SPI WHERE STATUS IN ('1100','1200','1300','1400','1500','1700','1701','3000','9000','9002','9003','9004','9006') AND DATA_OT BETWEEN {1} AND {2})", sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    // Выбираем ключи тех у кого в тек.году сменилась служба (VREMI_V_SL)
                    DataTable dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM AAQQ {0} AND FAMILIYA <> '' AND VREMI_V_SL >= {1} AND VREMI_V_SL <= {2}", current_sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    // Проверяем была ли предыдущая служба в послужном другая...
                    pos[59] = 0;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        // Проверяем есть ли такой ключ в PoslKeys1
                        if (PoslKeys1.Contains(Convert.ToInt16(dt.Rows[i]["KEY_1"]))) pos[59]++;
                    }
                    dt.Dispose();
                }
                #endregion

                #region // Выбыло сотрудников в другие службы
                if (current_dg_code != "010" && current_dg_code != "013")
                {
                    PrStat("Выбыло сотрудников в другие службы");
                    // Выбираем ключи всех у кого в тек.году сменилась служба (VREMI_V_SL)
                    DataTable dt = DataProvider._getDataSQL(KConn, String.Format("SELECT KEY_1 FROM AAQQ {0} AND DOLZNOST < '800000' AND FAMILIYA <> '' AND VREMI_V_SL >= {1} AND VREMI_V_SL <= {2}", current_sql_text, Convert.ToDateTime("01.01." + calc_year).ToOADate(), calc_date.ToOADate()));
                    // Проверяем была ли предыдущая служба в послужном другая...
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        // Проверяем есть ли такой ключ в PoslKeys2
                        // Если в текущей форме присутствует признак службы проверяем не была ли она предыдущей...
                        pos[60] = 0;
                        if (current_dg_sluzb != "" && current_dg_sluzb != null)
                        {
                            string[] sl = current_dg_sluzb.Split(Convert.ToChar(","));
                            foreach (string sluzb in sl)
                            {
                                smen item = new smen();
                                item.key = Convert.ToInt16(dt.Rows[i]["KEY_1"]);
                                item.sluzba_prev = Convert.ToInt16(sluzb);

                                if (PoslKeys2.Contains(item)) pos[60]++;
                            }
                        }
                        // если нет просто сравниваем 1 и 2 
                        else
                        {
                            if (PoslKeys1.Contains(Convert.ToInt16(dt.Rows[i]["KEY_1"]))) pos[60]++;
                        }
                    }
                    dt.Dispose();
                }
                #endregion 
 
                edit_pos59.Text = pos[59].ToString();
                edit_pos60.Text = pos[60].ToString();

                SavePos2OneK(current_dg_code);
            }
        }

        // Печать всего отчета (всех ДГ в отдельные файлы)
        private void print_all_report_button_Click(object sender, EventArgs e)
        {
            StatusBar.Panels[1].ProgressBarValue = 0;
            StatusBar.Panels[1].ProgressBarMinValue = 0;
            StatusBar.Panels[1].ProgressBarMaxValue = repDG.Count;

            foreach (string code in repDG)
            {
                StatusBar.Panels[0].Text = String.Format("Распечатка ДГ: {0}", code);

                int icode = Convert.ToInt16(code);

                if (icode != 300 && icode != 301 && icode != 900 && icode != 901)
                {
                    GridEXColumn cl;
                    cl = grid.RootTable.Columns["KEY_OTCH"];
                    GridEXFilterCondition cond = new GridEXFilterCondition(cl, ConditionOperator.Equal, code);
                    grid.RootTable.FilterCondition = cond;
                    report.LoadDocument(rep_filename.Text);
                    GridEXRow row = grid.CurrentRow;

                    Write2Report("month", Month2String(Convert.ToInt16(calc_month)));
                    Write2Report("yy", calc_year);
                    Write2Report("y1", calc_year[2].ToString());
                    Write2Report("y2", calc_year[3].ToString());
                    Write2Report("m1", calc_month[0].ToString());
                    Write2Report("m2", calc_month[1].ToString());
                    Write2Report("DG_name", row.Cells["NAME_FORM"].Text);
                    Write2Report("dgc1", code[0].ToString());
                    Write2Report("dgc2", code[1].ToString());
                    Write2Report("dgc3", code[2].ToString());

                    for (int i = 1; i < 127; i++) Write2Report("item" + i.ToString(), oneK[icode, i].ToString());

                    Write2Report("ruk1_dol", viz_ruk1_dol.Text);
                    Write2Report("ruk1_name", viz_ruk1_name.Text);
                    Write2Report("ruk1_zvan", viz_ruk1_zvan.Text);

                    Write2Report("ruk2_dol", "Начальник УРЛС УМВД");
                    Write2Report("ruk2_name", viz_ruk2_name.Text);
                    Write2Report("ruk2_zvan", viz_ruk2_zvan.Text);

                    Write2Report("isp_name", viz_isp_name.Text);
                    Write2Report("isp_zvan", viz_isp_zvan.Text);

                    report.SaveDocument(path_save_report_edit.Text + code + ".docx", DocumentFormat.OpenXml);

                    StatusBar.Panels[1].ProgressBarValue++;
                    Application.DoEvents();
                }
            }
            
            MessageBox.Show("Готово!");
            grid.RootTable.RemoveFilter();
            StatusBar.Panels[1].ProgressBarValue = 0;
        }

        // Перегруппировка образования...
        private void learn_prepare_button_Click(object sender, EventArgs e)
        {
            // Выбираем всех у кого в основной БД образование не высшее
            DataTable dt = DataProvider._getDataSQL(KConn, "SELECT KEY_1, OBRAZ_LIC2 FROM AAQQ WHERE DOLZNOST < '800000' AND FAMILIYA <> '' AND OBRAZ_LIC2 <> '10' ORDER BY KEY_1");
            StatusBar.Panels[1].ProgressBarMaxValue = dt.Rows.Count;
            StatusBar.Panels[1].ProgressBarValue = 0;
            int cnt = 0;
            int err = 0;
            // проходим всех смотрим что еще есть в LEARN
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string id = dt.Rows[i]["KEY_1"].ToString();
                string vid = dt.Rows[i]["OBRAZ_LIC2"].ToString();
                // Выбираем по ключу всё высшее или ср.спец.
                DataTable ln = DataProvider._getDataSQL(KConn,
                    String.Format("SELECT VID, RTRIM(TYPE) AS TYPE, UCH_ZAV, SPECALNOST, YEAR FROM LEARN WHERE KEY_1 = {0} AND VID IN ('10','20') ORDER BY YEAR DESC", id));

                if (ln.Rows.Count > 0)
                {
                    for (int j = 0; j < ln.Rows.Count; j++)
                    {
                        // Если есть высшее - меняем
                        if (ln.Rows[j]["VID"].ToString() == "10")
                        {
                            // а в основной ср.спец. или среднее
                            if (vid == "30" || vid == "20" || vid == "40" || vid == "50")
                            {
                            // Обновляем на высшее
                            int res = DataProvider._insDataSQL(KConn, String.Format("UPDATE AAQQ SET OBRAZ_LIC2='10', OBRAZ_LIC1='{0}', OBRAZ_LIC3='{1}', UCHZAV='{2}', DAT_OKUZ='{3}' WHERE KEY_1 = {4}",
                                ln.Rows[j]["TYPE"], ln.Rows[j]["SPECALNOST"], ln.Rows[j]["UCH_ZAV"], ln.Rows[j]["YEAR"], id));
                            if (res > 0) cnt++;
                            else err++;
                            }
                        }
                        // Если есть ср.спец., а в основной среднее - меняем
                        else if (ln.Rows[j]["VID"].ToString() == "20" && (vid == "30" || vid == "40" || vid == "50"))
                        {
                            // Обновляем на среднее спец.
                            int res = DataProvider._insDataSQL(KConn, String.Format("UPDATE AAQQ SET OBRAZ_LIC2='20', OBRAZ_LIC1='{0}', OBRAZ_LIC3='{1}', UCHZAV='{2}', DAT_OKUZ='{3}' WHERE KEY_1 = {4}",
                                ln.Rows[j]["TYPE"], ln.Rows[j]["SPECALNOST"], ln.Rows[j]["UCH_ZAV"], ln.Rows[j]["YEAR"], id));
                            if (res > 0) cnt++;
                            else err++;
                        }
                    }
                }
                ln.Dispose();
                StatusBar.Panels[1].ProgressBarValue++;
            }

            MessageBox.Show(String.Format("Готово!\nОбновлено {0} записей, ошибок - {1}", cnt, err));
            StatusBar.Panels[1].ProgressBarValue = 0;
            dt.Dispose();
        }

        // Проверка суммы по определенной позиции ДГ
        private void checksum_button_Click(object sender, EventArgs e)
        {
            // Проверяемая ДГ
            int DG = Convert.ToInt16(checksum_DG.Text);
            // Проверяемая позиция
            int Pos = Convert.ToInt16(checksum_pos.Text);

            switch (DG)
            {
                case 20: { checksum_text.Text = (oneK[971, Pos] + oneK[978, Pos] + oneK[360, Pos]).ToString(); break; }
                case 30: { checksum_text.Text = (oneK[972, Pos] + oneK[979, Pos] + oneK[362, Pos]).ToString(); break; }
                case 971: { checksum_text.Text = (oneK[900, Pos] + oneK[300, Pos] + oneK[92, Pos] + oneK[117, Pos] + oneK[170, Pos] + oneK[210, Pos] + oneK[200, Pos] + oneK[204, Pos] + oneK[191, Pos] + oneK[73, Pos] + oneK[290, Pos] + oneK[220, Pos] + oneK[350, Pos] + oneK[291, Pos] + oneK[245, Pos] + oneK[351, Pos] + oneK[193, Pos] + oneK[994, Pos] + oneK[355, Pos] + oneK[410, Pos] + oneK[440, Pos]).ToString(); break; }
                case 972: { checksum_text.Text = (oneK[301, Pos] + oneK[203, Pos] + oneK[83, Pos] + oneK[280, Pos] + oneK[246, Pos] + oneK[352, Pos] + oneK[194, Pos] + oneK[993, Pos] + oneK[356, Pos] + oneK[420, Pos] + oneK[450, Pos]).ToString(); break; }
                case 978: { checksum_text.Text = (oneK[40, Pos] + oneK[45, Pos] + oneK[50, Pos] + oneK[60, Pos]).ToString(); break; }
                case 979: { checksum_text.Text = (oneK[63, Pos]).ToString(); break; }
                case 290: { checksum_text.Text = (oneK[283, Pos] + oneK[270, Pos] + oneK[271, Pos] + oneK[260, Pos]).ToString(); break; }
            }
        }
        
        // Переключение вкладок
        private void mainTabControl_Click(object sender, EventArgs e)
        {
            // Если вкладка "проверка"
            if (mainTabControl.SelectedTabPageIndex == 4)
            {
                checksum_DG.Text = current_dg_code;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                rep_filename.Text = openFileDialog.FileName;
                report.LoadDocument(rep_filename.Text); 
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                path_save_report_edit.Text = folderBrowser.SelectedPath;
            }
        }

        // Сохранение настроек...
        private void save_cfg()
        {
            try
            {
                Configuration cfg = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                cfg.AppSettings.Settings["rep_file_path"].Value = rep_filename.Text;
                cfg.AppSettings.Settings["rep_word_save_path"].Value = path_save_report_edit.Text;
                cfg.AppSettings.Settings["UMVDboss_dol"].Value = viz_ruk1_dol.Text;
                cfg.AppSettings.Settings["UMVDboss_zv"].Value = viz_ruk1_zvan.Text;
                cfg.AppSettings.Settings["UMVDboss_name"].Value = viz_ruk1_name.Text;
                cfg.AppSettings.Settings["URLSboss_zv"].Value = viz_ruk2_zvan.Text;
                cfg.AppSettings.Settings["URLSboss_name"].Value = viz_ruk2_name.Text;
                cfg.AppSettings.Settings["ISP_zv"].Value = viz_isp_zvan.Text;
                cfg.AppSettings.Settings["ISP_name"].Value = viz_isp_name.Text;
                cfg.AppSettings.Settings["UUPSelo"].Value = UUPSelo_Stat.Text;
                cfg.AppSettings.Settings["calc_date"].Value = calc_on_date.Value.ToShortDateString();
                cfg.AppSettings.Settings["script_date"].Value = script_date.Value.ToShortDateString();
                cfg.Save();
                MessageBox.Show("Настройки успешно сохранены!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // Чтение настроек...
        private void read_cfg()
        {
            try
            {
                Configuration cfg = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                rep_filename.Text = cfg.AppSettings.Settings["rep_file_path"].Value;
                path_save_report_edit.Text = cfg.AppSettings.Settings["rep_word_save_path"].Value;
                viz_ruk1_dol.Text = cfg.AppSettings.Settings["UMVDboss_dol"].Value;
                viz_ruk1_zvan.Text = cfg.AppSettings.Settings["UMVDboss_zv"].Value;
                viz_ruk1_name.Text = cfg.AppSettings.Settings["UMVDboss_name"].Value;
                viz_ruk2_zvan.Text = cfg.AppSettings.Settings["URLSboss_zv"].Value;
                viz_ruk2_name.Text = cfg.AppSettings.Settings["URLSboss_name"].Value;
                viz_isp_zvan.Text = cfg.AppSettings.Settings["ISP_zv"].Value;
                viz_isp_name.Text = cfg.AppSettings.Settings["ISP_name"].Value;
                UUPSelo_Stat.Text = cfg.AppSettings.Settings["UUPSelo"].Value;
                calc_on_date.Value = Convert.ToDateTime(cfg.AppSettings.Settings["calc_date"].Value);
                script_date.Value = Convert.ToDateTime(cfg.AppSettings.Settings["script_date"].Value);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // Сохранение настроек (кнопка)
        private void save_config_button_Click(object sender, EventArgs e)
        {
            save_cfg();
        }

        // Расчет только выбранных позиций
        private void fixed_pos_calc_button_Click(object sender, EventArgs e)
        {
            // Проверка начальных условий и настроек...
            if (kval_check.Checked)
            {
                if (smen_check.Checked)
                {
                    // Если все готово
                    DataRowCollection rc = pergroupDataSet.Tables[0].Rows;

                    StatusBar.Panels[1].ProgressBarMaxValue = rc.Count;
                                        
                    // Проходим все ДГ
                    for (int i = 0; i < rc.Count; i++)
                    {
                        string code = rc[i]["KEY_OTCH"].ToString();
                        // Если не "хитрые" и не вольные
                        if (code != "010" && code != "013" && code != "300" && code != "301" && code != "900" && code != "901")
                        {
                            StatusBar.Panels[0].Text = String.Format("[{0}] - {1}", code, rc[i]["NAME_FORM"].ToString());
                            // Вычисляем ДГ
                            CalcDG(code, rc[i]["TEXT_QRY"].ToString(), fixed_pos_text.Text);
                            StatusBar.Panels[1].ProgressBarValue++;
                        }
                        // Не засыпаем...
                        Application.DoEvents();
                    }

                    StatusBar.Panels[1].ProgressBarValue = 0;
                    StatusBar.Panels[0].Text = "Расчет окончен, не забудьте сохраниться...";
                    save_all_express_button.Enabled = true;

                }
                else MessageBox.Show("Не выполнена подготовка массива ключей для сменяемости!");
            }
            else MessageBox.Show("Не выполнена предварительная проверка соответствия квалификационным требованиям!");
        }

        // Предварительный расчет стажа для всех сотрудников... (для ускорения расчета)
        private void prepare_stage_button_Click(object sender, EventArgs e)
        {
            // Выбираем все возможные ключи
            DataTable dt = DataProvider._getDataSQL(KConn, "SELECT DISTINCT KEY_1 FROM AAQQ WHERE (KEY_1 <> 0) AND (FAMILIYA <> '') AND (DOLZNOST < '800000') ORDER BY KEY_1");
                                                            //"UNION " +
                                                            //"SELECT DISTINCT KEY_1 FROM ARCHIVE WHERE (KEY_1 <> 0) AND (DAT_REG >= '2016-01-01') UNION "+
                                                            //"SELECT DISTINCT KEY_1 FROM PRIEM WHERE (KEY_1 <> 0) AND (DAT_REG >= '2016-01-01') UNION " +
                                                            //"SELECT DISTINCT KEY_1 FROM RESERV WHERE (KEY_1 <> 0) AND (DATA_ZACH >= '2015-01-01') UNION " +
                                                            //"SELECT DISTINCT KEY_1 FROM VYEZD WHERE (KEY_1 <> 0) AND (DAT_REG >= '2016-01-01')");
            DataRowCollection rc;
            rc = dt.Rows;
            dt.Dispose();

            StatusBar.Panels[1].ProgressBarMaxValue = rc.Count;
            StatusBar.Panels[1].ProgressBarValue = 0;

            if (Vysluga.Count != 0) Vysluga.Clear();

            // проходим все ключи, считаем выслугу...
            for (int i = 0; i < rc.Count; i++)
            {
                string id = rc[i]["KEY_1"].ToString();
                vysl item = new vysl();
                TPeriod stage = new TPeriod();
                stage = CalcStage(id, true); // календарь
                item.key = id;
                item.years_c = stage.y;
                stage = CalcStage(id, false); // льготка
                item.years_l = stage.y;
                Vysluga.Add(item);
                Application.DoEvents();
                StatusBar.Panels[0].Text = String.Format("Смотрим: ключ {0}", id);
                StatusBar.Panels[1].ProgressBarValue++;
            }
            StatusBar.Panels[0].Text = "";
            MessageBox.Show("Готово!", "Расчет стажа");
        }

        // Сохранение массива выслуги в файл
        private void prs_filesave_button_Click(object sender, EventArgs e)
        {
            if (Vysluga.Count != 0)
            {
                System.IO.StreamWriter vs = new StreamWriter("Vysluga.txt");
                               
                foreach (vysl item in Vysluga)
                {
                    vs.WriteLine(String.Format("{0}", item.key));
                    vs.WriteLine(String.Format("{0}", item.years_c));
                    vs.WriteLine(String.Format("{0}", item.years_l));
                }
                vs.Flush();
                vs.Close();
                MessageBox.Show("Данные сохранены!");
            }
        }

        // Загрузка массива выслуги из файла
        private void prs_fileload_button_Click(object sender, EventArgs e)
        {
            if (Vysluga.Count != 0) Vysluga.Clear();
                
            System.IO.StreamReader vs = new StreamReader("Vysluga.txt");
            
            vysl item = new vysl();
            do
            {
                item.key = vs.ReadLine();
                item.years_c = Convert.ToInt16(vs.ReadLine());
                item.years_l = Convert.ToInt16(vs.ReadLine());
                Vysluga.Add(item);
            } while (!vs.EndOfStream);
            vs.Close();

            MessageBox.Show(String.Format("Данные загружены!\nVysluga - {0}", Vysluga.Count));
            prs_check.Checked = true;
        }

        // Загрузка шаблона отчета
        private void load_report_template_button_Click(object sender, EventArgs e)
        {
            if (File.Exists(rep_filename.Text))
            {
                report.LoadDocument(rep_filename.Text);
                MessageBox.Show("Файл шаблона отчета успешно загружен!");
            }
            else MessageBox.Show("Файл шаблона отчета:" + rep_filename.Text + " не найден!");
        }
    }
}
