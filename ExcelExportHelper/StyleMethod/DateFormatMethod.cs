using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;

namespace ExcelExportHelper
{
    internal class DateFormatMethod:CellStyleMethod
    {
        internal override ICellStyle SetCell(ICellStyle cellStyle)
        {
            IDataFormat format = Workbook.CreateDataFormat();
            cellStyle.DataFormat = format.GetFormat("yyyy/mm/dd");
            return cellStyle;
        }
    }
}
