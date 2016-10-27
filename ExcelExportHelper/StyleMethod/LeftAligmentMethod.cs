using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;

namespace ExcelExportHelper
{
    internal class LeftAligmentMethod:CellStyleMethod
    {
        internal override ICellStyle SetCell(ICellStyle cellStyle)
        {
            cellStyle.Alignment = HorizontalAlignment.Left;
            return cellStyle;
        }
    }
}
