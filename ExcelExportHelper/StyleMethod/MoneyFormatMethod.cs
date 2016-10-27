using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;

namespace ExcelExportHelper
{
    internal class MoneyFormatMethod:CellStyleMethod
    {
        internal override ICellStyle SetCell(ICellStyle cellStyle)
        {
            IDataFormat format = Workbook.CreateDataFormat();
            cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("¥#,##0.00");
            return cellStyle;
        }
    }
}
