using NPOI.SS.UserModel;

namespace ExcelExportHelper
{
    internal class RightAligmentMethod:CellStyleMethod
    {
        internal override ICellStyle SetCell(ICellStyle cellStyle)
        {
            cellStyle.Alignment= HorizontalAlignment.Right;
            return cellStyle;
        }
    }
}
