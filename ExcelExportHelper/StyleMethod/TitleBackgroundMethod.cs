using NPOI.SS.UserModel;
using NPOI.HSSF.Util;

namespace ExcelExportHelper
{
    internal class TitleBackgroundMethod:CellStyleMethod
    {
        internal override ICellStyle SetCell(ICellStyle cellStyle)
        {
            cellStyle.FillForegroundColor = HSSFColor.White.Index;
            cellStyle.FillBackgroundColor = HSSFColor.Grey50Percent.Index;
            cellStyle.FillPattern = FillPattern.SolidForeground;
            return cellStyle;
        }
    }
}
