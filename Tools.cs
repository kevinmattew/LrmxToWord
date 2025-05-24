using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace 干部任免审批表转换
{
    public  class Tools
    {
        public static Boolean isStartWithNumber(String str)
        {
            str = str.Trim();
            Regex regex = new Regex("[0-9]");
            return regex.IsMatch(str[0].ToString());
        }


        public static int  JiSuanNianLing(string chuShengRiQi ,string jiSuanNianLingShiJian=""  ,string tianBiaoShiJian ="" ) {
            int nianLing = 0;
            jiSuanNianLingShiJian = jiSuanNianLingShiJian.Replace("-", "");
            tianBiaoShiJian = tianBiaoShiJian.Replace("-", "");
            if (string.IsNullOrEmpty(jiSuanNianLingShiJian)  &&  string.IsNullOrEmpty(tianBiaoShiJian))
            {
                nianLing = countAage(chuShengRiQi, DateTime.Now.ToString("yyyyMM"));

            }
            else {
                if (string.IsNullOrEmpty(tianBiaoShiJian)  && ( jiSuanNianLingShiJian.Length == 6 || jiSuanNianLingShiJian.Length == 8))
                {

                    nianLing = countAage(chuShengRiQi, jiSuanNianLingShiJian);
                }
                else {

                    if(tianBiaoShiJian.Length == 6 ||  tianBiaoShiJian.Length == 8) { 
                    nianLing = countAage(chuShengRiQi, tianBiaoShiJian);
                    }
                }
            
            }
            

            return nianLing;
        }

        public static  int  countAage(string startDate, string endDate)
        {
            int NianLing = 0;
            NianLing =  int.Parse(endDate.Substring(0,4))  -  int.Parse(startDate.Substring(0,4)) ;
            if (int.Parse(endDate.Substring(4, 2)) < int.Parse(startDate.Substring(4, 2))){
                NianLing--;
            }
            return NianLing;

        }

        public static string  NullToEmpty(object o) { 
                if(o == null) return string.Empty;
            return o.ToString(); 
        
        }


        public static string GetDate(string IDCard)
        {
            string BirthDay = " ";
            string strYear;
            string strMonth;
            string strDay;
            if (IDCard.Length == 15)
            {
                strYear = IDCard.Substring(6, 2);
                strMonth = IDCard.Substring(8, 2);
                strDay = IDCard.Substring(10, 2);
                BirthDay = "19" + strYear + strMonth +   strDay;
            }
            if (IDCard.Length == 18)
            {
                strYear = IDCard.Substring(6, 4);
                strMonth = IDCard.Substring(10, 2);
                strDay = IDCard.Substring(12, 2);
                BirthDay = strYear +  strMonth +  strDay;
            }
            return BirthDay;
        }
 

    }
}
