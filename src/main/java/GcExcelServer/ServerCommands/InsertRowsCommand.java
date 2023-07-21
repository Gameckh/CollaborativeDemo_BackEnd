package GcExcelServer.ServerCommands;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;

public class InsertRowsCommand extends SelectionCommandBase{
    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = this.getWorksheet(workbook);
        if(worksheet == null){
            return;
        }
        for(SJSRange range: this.getSelections()){
            worksheet.getRange(range.getRow(),range.getCol(),range.getRowCount(), range.getColCount()).getEntireRow().insert();
        }
    }
}
