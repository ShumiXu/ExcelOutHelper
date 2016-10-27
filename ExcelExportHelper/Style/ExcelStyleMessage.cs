using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace ExcelExportHelper
{
    /// <summary>
    /// Excel样式管理
    /// </summary>
    internal class ExcelStyleMessage
    {
        /// <summary>
        /// 样式集合
        /// </summary>
        private Dictionary<string, ICellStyle> StyleList { get; set; }

        public ExcelStyleMessage()
        {
            StyleList = new Dictionary<string, ICellStyle>();
        }

        /// <summary>
        /// 根据枚举加载对应操作类
        /// </summary>
        /// <param name="excelStyle"></param>
        /// <returns></returns>
        private CellStyleMethod GetStyleMethod(ExcelStyle excelStyle)
        {
            switch (excelStyle)
            {
                case ExcelStyle.title:
                    return new TitleBackgroundMethod();
                case ExcelStyle.right:
                    return new RightAligmentMethod();
                case ExcelStyle.money:
                    return new MoneyFormatMethod();
                case ExcelStyle.left:
                    return new LeftAligmentMethod();
                case ExcelStyle.date:
                    return new DateFormatMethod();
                default:
                    throw new ArgumentException("参数无效");
            }
        }

        internal ICellStyle GetCellStyle<T>(T workbook, ExcelStyle excelStyle) where T:IWorkbook
        {
            if (StyleList.ContainsKey(excelStyle.ToString()))
            {
                return StyleList[excelStyle.ToString()];
            }

            ICellStyle _cellStyle = workbook.CreateCellStyle();
            _cellStyle.BorderBottom = BorderStyle.Thin;
            _cellStyle.BorderLeft = BorderStyle.Thin;
            _cellStyle.BorderRight = BorderStyle.Thin;
            _cellStyle.BorderTop = BorderStyle.Thin;
            CellStyleMethod styleMethod;
            if (excelStyle.ToString().IndexOf(',')>-1)
            {
                foreach (var styleItem in excelStyle.ToString().Replace(" ","").Split(','))
                {
                    if (Enum.IsDefined(typeof(ExcelStyle),styleItem))
                    {
                        ExcelStyle stylemodel = (ExcelStyle)Enum.Parse(typeof(ExcelStyle),styleItem);
                        styleMethod = GetStyleMethod(stylemodel);
                        styleMethod.SetCell(_cellStyle);
                    }
                }
                StyleList.Add(excelStyle.ToString(), _cellStyle);
                return _cellStyle;
            }
            styleMethod = GetStyleMethod(excelStyle);
            styleMethod.SetCell(_cellStyle);
            StyleList.Add(excelStyle.ToString(), _cellStyle);
            return _cellStyle;
        }
    }
}
