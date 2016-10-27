using System;

namespace ExcelExportHelper
{
    public class ExcelInfoAttribute:Attribute
    {
        public string Name { get; set; }

        public int Width { get; set; }

        public ExcelStyle ExcelStyel { get; set; }

        public ExcelInfoAttribute(string name,int width=2800,ExcelStyle excelStyle=ExcelStyle.left)
        {
            this.Name = name;
            this.Width = width;
            this.ExcelStyel = excelStyle;
        }
    }
}
