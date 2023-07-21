package GcExcelServer.ServerCommands;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;

public class SetFontSizeCommand extends SelectionCommandBase{
    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = this.getWorksheet(workbook);
        if(worksheet == null){
            return;
        }
        int fontSize =  Integer.parseInt(this.value.replace("pt",""));
        for(SJSRange range: this.getSelections()){
            worksheet.getRange(range.getRow(),range.getCol(),range.getRowCount(), range.getColCount()).getFont().setSize(fontSize);
        }
    }

    private String value;

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }
}
