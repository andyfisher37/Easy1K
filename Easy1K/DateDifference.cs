using System;

namespace Easy1K
{
    public struct TPeriod
    {
        public int d;
        public int m;
        public int y;

        // Сложение периодов
        public void Add(int d1, int m1, int y1)
        {
            d = d + d1;
            m = m + m1;
            y = y + y1;
            if (d >= 30) { m++; d = d - 30; }
            if (m >= 12) { y++; m = m - 12; }
        }

        // Вычитание периодов
        public void Substract(int d1, int m1, int y1)
        {
            d = d - d1;
            m = m - m1;
            y = y - y1;
            if (d < 0) { m--; d = 30 - d; }
            if (m < 0) { y--; m = 12 - m; }
        }

        // Очистка периода
        public void Clear()
        {
            d = 0;
            m = 0;
            y = 0;
        }

        // Проверка на пустоту
        public bool isEmpty()
        {
            if (d==0 && m==0 && y==0) return true;
            else return false;
        }

        // Преобразование к виду: (00л00м00д)
        public string ToShortString()
        {
            string tmp = "";
            if (y < 10) tmp += String.Format("0{0}л", y);
            else tmp += y.ToString() + "л";
            if (m < 10) tmp += String.Format("0{0}м", m);
            else tmp += m.ToString() + "м";
            if (d < 10) tmp += String.Format("0{0}д", d);
            else tmp += d.ToString() + "д";
            return tmp;
        }

        // Преобразование к виду: (00 лет 00 месяцев 00 дней)
        public string ToLongString()
        {
            string tmp = "";
            if (y < 10) tmp += String.Format("0{0} лет ", y);
            else tmp += y.ToString() + " лет ";
            if (m < 10) tmp += String.Format("0{0} месяцев ", m);
            else tmp += m.ToString() + " месяцев ";
            if (d < 10) tmp += String.Format("0{0} дней ", d);
            else tmp += d.ToString() + " дней ";
            return tmp;
        }

        // Возвращение результата в виде строки: ## лет ## месяцев ## дней
        public string ToSpecString()
        {
            string tmp = "";
            if (y != 0)
            {
                if (y < 10) tmp += String.Format("0{0} лет ", y);
                else tmp += y.ToString() + " лет ";
            }
            if (m != 0)
            {
                if (m < 10) tmp += String.Format("0{0} месяцев ", m);
                else tmp += m.ToString() + " месяцев ";
            }
            if (d != 0)
            {
                if (d < 10) tmp += String.Format("0{0} дней", d);
                else tmp += d.ToString() + " дней";
            }
            return tmp;
        }

        // Возвращение результата в виде строки: ## (прописью) лет ## (прописью) месяцев ## (прописью) дней
        public string ToSpecLongString()
        {
            string tmp = "";
            if (y >= 0)
            {
                if (y < 10) tmp += String.Format("0{0} (<i>{1}</i>) лет ", y, DigToStr("0" + y.ToString()));
                else tmp += String.Format("{0} (<i>{1}</i>) лет ", y, DigToStr(y.ToString()));
            }
            if (m >= 0)
            {
                if (m < 10) tmp += String.Format("0{0} (<i>{1}</i>) месяцев ", m, DigToStr("0" + m.ToString()));
                else tmp += String.Format("{0} (<i>{1}</i>) месяцев ", m, DigToStr(m.ToString()));
            }
            if (d >= 0)
            {
                if (d < 10) tmp += String.Format("0{0} (<i>{1}</i>) дней", d, DigToStr("0" + d.ToString()));
                else tmp += String.Format("{0} (<i>{1}</i>) дней", d, DigToStr(d.ToString()));
            }
            return tmp;
        }

        // Преобразование числительных в "пропись"
        public static string DigToStr(string dig)
        {
            switch (dig)
            {
                case "00": return "ноль"; 
                case "01": return "один"; 
                case "02": return "два"; 
                case "03": return "три"; 
                case "04": return "четыре";
                case "05": return "пять"; 
                case "06": return "шесть";
                case "07": return "семь"; 
                case "08": return "восемь";
                case "09": return "девять";
                case "10": return "десять";
                case "11": return "одиннадцать";
                case "12": return "двенадцать"; 
                case "13": return "тринадцать"; 
                case "14": return "четырнадцать";
                case "15": return "пятнадцать"; 
                case "16": return "шестнадцать";
                case "17": return "семнадцать"; 
                case "18": return "восемнадцать";
                case "19": return "девятнадцать";
                case "20": return "двадцать"; 
                case "21": return "двадцать один"; 
                case "22": return "двадцать два"; 
                case "23": return "двадцать три"; 
                case "24": return "двадцать четыре"; 
                case "25": return "двадцать пять"; 
                case "26": return "двадцать шесть"; 
                case "27": return "двадцать семь"; 
                case "28": return "двадцать восемь"; 
                case "29": return "двадцать девять"; 
                case "30": return "тридцать"; 
                case "31": return "тридцать один"; 
                case "32": return "тридцать два"; 
                case "33": return "тридцать три"; 
                case "34": return "тридцать четыре"; 
                case "35": return "тридцать пять"; 
                case "36": return "тридцать шесть"; 
                case "37": return "тридцать семь"; 
                case "38": return "тридцать восемь"; 
                case "39": return "тридцать девять"; 
                case "40": return "сорок"; 
                case "41": return "сорок один"; 
                case "42": return "сорок два"; 
                case "43": return "сорок три"; 
                case "44": return "сорок четыре"; 
                case "45": return "сорок пять"; 
                default: return "-"; 
            }

        }
    }

    public class DateDifference
    {
        /// Задаем количество дней в месяцах года; index 0=> январь and 11=> декабрь
        /// февраль может содержать 28 или 29 дней, поэтому его значение -1
        /// точное количество дней в ферале высчитывается позднее.

        private readonly int[] monthDay = new int[12] { 31, -1, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
                
        /// Дата начала периода
        
        private DateTime fromDate;

        /// Дата окончания периода
        
        private DateTime toDate;

        /// Переменные для вывода года, месяца, дней
        
        private int year;
        private int month;
        private int day;
        
        // Период между датами (в календарном исчислении) лет, мес, дней
        public DateDifference(DateTime d1, DateTime d2)
        {
            int increment;
            
           // проверка на то какая из дат больше...
            if (d1 > d2)
            {
                fromDate = d2;
                toDate = d1;
            }
            else
            {
                fromDate = d1;
                toDate = d2;
            }

            
            // Вычисление дней
            increment = 0;

            if (fromDate.Day > toDate.Day)
            {
                increment = monthDay[this.fromDate.Month - 1];
            }
            
            // Если месяц февраль
            if (increment == -1)
            {
                if (DateTime.IsLeapYear(this.fromDate.Year))
                {
                    // проверка на високосный год...
                    increment = 29;
                }
                else
                {
                    increment = 28;
                }
            }

            if (increment != 0)
            {
                day = (toDate.Day + increment) - fromDate.Day;
                increment = 1;
            }
            else
            {
                day = toDate.Day - fromDate.Day;
            }

            
            // Вычисление месяцев
            if ((fromDate.Month + increment) > toDate.Month)
            {
                month = (toDate.Month + 12) - (fromDate.Month + increment);
                increment = 1;
            }
            else
            {
                month = (toDate.Month) - (fromDate.Month + increment);
                increment = 0;
            }

            
            // Вычисление лет
            year = toDate.Year - (fromDate.Year + increment);

        }

        // Период между датами (в льготном исчислении) лет, мес, дней
        public DateDifference(DateTime d1, DateTime d2, double k)
        {
            int increment;

            // проверка на то какая из дат больше...
            if (d1 > d2)
            {
                fromDate = d2;
                toDate = d1;
            }
            else
            {
                fromDate = d1;
                toDate = d2;
            }


            // Вычисление дней
            increment = 0;

            if (fromDate.Day > toDate.Day)
            {
                increment = monthDay[this.fromDate.Month - 1];
            }

            // Если месяц февраль
            if (increment == -1)
            {
                if (DateTime.IsLeapYear(this.fromDate.Year))
                {
                    // проверка на високосный год...
                    increment = 29;
                }
                else
                {
                    increment = 28;
                }
            }

            if (increment != 0)
            {
                day = (int)Math.Truncate(k*((toDate.Day + increment) - fromDate.Day));
                increment = 1;
            }
            else
            {
                day = (int)Math.Truncate(k*(toDate.Day - fromDate.Day));
            }


            // Вычисление месяцев
            if ((fromDate.Month + increment) > toDate.Month)
            {
                month = (int)Math.Truncate(k*((toDate.Month + 12) - (fromDate.Month + increment)));
                increment = 1;
            }
            else
            {
                month = (int)Math.Truncate(k*((toDate.Month) - (fromDate.Month + increment)));
                increment = 0;
            }


            // Вычисление лет
            year = (int)Math.Truncate(k*(toDate.Year - (fromDate.Year + increment)));

        }


        // Возвращение результата в виде строки: ## лет ## месяцев ## дней
        public override string ToString()
        {
            return String.Format("{0} лет, {1} месяцев, {2} дней", year, month, day);
        }

        // Возвращение результата в виде строки: ## (прописью) лет ## (прописью) месяцев ## (прописью) дней
        
       

        public int Years
        {
            get { return year; }
        }

        public int Months
        {
            get { return month; }
        }

        public int Days
        {
            get { return day; }
        }

    }
}
