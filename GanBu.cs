using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 干部任免审批表转换
{
    public  class GanBu
    {
      public   String XingMing;   //姓名  1,0,0   1,0,1
        public String XingBie;  //性别  1,0,2   1,0,3
        public String ChuShengNianYue;  //出生年月  YYYYMM格式    1,0,4  1,0,5
        public String MinZu;  //民族     1,1,0   1,1,1
        public String JiGuan;   //籍贯  1,1,2     1,1,3
        public String ChuShengDi;// 出生地  1,1,4  1,1,5
        public String RuDangShiJian;// 入党时间   1,2,0
        public String CanJiaGongZuoShiJian;//参加工作时间  1,2,2
        public String JianKangZhuangKuang;//健康状况  1,2,4
        public String ZhuanYeJiShuZhiWu;//专业技术职务  1,3,0
        public String ShuXiZhuanYeYouHeZhuanChang;//熟悉专业有何专长   1,3,2
        public String QuanRiZhiJiaoYu_XueLi;//全日制教育学历       1,4,2
        public String QuanRiZhiJiaoYu_XueWei;//全日制教育学位 1,4,2
        public String QuanRiZhiJiaoYu_XueLi_BiYeYuanXiaoXi;//全日制毕业学校   1,4,4
        public String QuanRiZhiJiaoYu_XueWei_BiYeYuanXiaoXi;//全日制毕业院校专业     1,5,4
        public String ZaiZhiJiaoYu_XueLi;//在职教育学历   1,6,2
        public String ZaiZhiJiaoYu_XueWei;//在职教育学位  1,6,2
        public String ZaiZhiJiaoYu_XueLi_BiYeYuanXiaoXi;//在职教育毕业院校    1,6,4
        public String ZaiZhiJiaoYu_XueWei_BiYeYuanXiaoXi;//在职教育毕业院校专业  1,7,4
        public String XianRenZhiWu;//现任职务     1,8,0
        public String NiRenZhiWu;//拟任职务    1,9,0
        public String NiMianZhiWu;//拟免职务   1,10,0
        public List<JianLi> JianLi;  //简历   1,11,0
        public String JiangChengQingKuang;//奖惩情况    2,0,0
        public String NianDuKaoHeJieGuo;//年度考核结果    2,1,0
        public String RenMianLiYou;//任免理由   2,2,0
        public   List<JiaTingChengYuan> JiaTingChengYuan;//家庭成员   2,4,1  称谓    2   姓名      3  年龄  4  政治面貌  5  工作单位及职务
        public String ChengBaoDanWei;//承报单位
        public String JiSuanNianLingShiJian;//计算年龄时间
        public String TianBiaoShiJian;//填报时间
        public String TianBiaoRen;//填报人
        public String ShenFenZheng;//身份证号码
        public String ZhaoPian;//  照片
        public String Version;//版本
        public int NianLing; //年龄

    }
}
