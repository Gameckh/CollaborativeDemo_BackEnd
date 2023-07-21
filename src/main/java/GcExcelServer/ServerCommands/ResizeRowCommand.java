package GcExcelServer.ServerCommands;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;

public class ResizeRowCommand extends SheetCommandBase{

    private RowRange[] rows;
    public RowRange[] getRows() {
        return rows;
    }
    public void setRows(RowRange[] rows) {
        this.rows = rows;
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
        if(worksheet == null){
            return;
        }

        for(RowRange rowRange: rows){
            for(int row = rowRange.getFirstRow(); row <= rowRange.getLastRow(); row++) {
                worksheet.getRows().get(row).setRowHeightInPixel(this.getSize());
            }
        }
    }

}
