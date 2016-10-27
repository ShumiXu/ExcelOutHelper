using System;
using ExcelExportHelper;
using System.Collections.Generic;

namespace MineTest
{
    public partial class Test : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            List<UserManagerTest> testList = new List<UserManagerTest>
            {
                new UserManagerTest{
                   CreateDate=DateTime.Now,Name="王小二",Old=20,Money=3.76
                },
                new UserManagerTest{
                    CreateDate=DateTime.Now,Name="李铁妹",Old=30,Money=9.78
                }
            };
            ExcelDownload download=new ExcelDownload("员工信息","年度员工汇总");
            download.ExportExcel<UserManagerTest>(testList);
        }
    }
}