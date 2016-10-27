using System;
using ExcelExportHelper;

namespace MineTest
{
    public class UserManagerTest
    {
        [ExcelInfo("名称")]
        public string Name { get; set; }
        [ExcelInfo("年龄",ExcelStyel=ExcelStyle.left)]
        public int Old { get; set; }
        [ExcelInfo("金额", ExcelStyel =ExcelStyle.right | ExcelStyle.money )]
        public double Money { get; set; }
        [ExcelInfo("时间", ExcelStyel = ExcelStyle.date | ExcelStyle.left)]
        public DateTime CreateDate { get; set; }

    }
}