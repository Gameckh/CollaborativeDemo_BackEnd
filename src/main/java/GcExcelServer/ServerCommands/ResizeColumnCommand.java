package GcExcelServer.ServerCommands;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;

public class ResizeColumnCommand extends SheetCommandBase{
    private ColumnRange[] columns;
    public ColumnRange[] getColumns() {
        return columns;
    }
    public void setColumns(ColumnRange[] columns) {
        this.columns = columns;
    }

    private int size;

    public int getSize() {
        return size;
    }

    public void setSize(int size) {
        this.size = size;
    }

    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = this.getWorksheet(workbook);

        for(ColumnRange columnRange: columns){
            for(int col = columnRange.getFirstCol(); col <= columnRange.getLastCol(); col++) {
                worksheet.getColumns().get(col).setColumnWidthInPixel(this.getSize());
            }
        }
    }
}
