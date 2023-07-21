package GcExcelServer.ServerCommands;

import com.grapecity.documents.excel.Color;
import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;

public class SetBackColorCommand extends SelectionCommandBase{
    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = this.getWorksheet(workbook);
        if(worksheet == null){
            return;
        }
        for(SJSRange range: this.getSelections()){
            int argb = Utility.ParseARGBColorFromHtml(this.getValue());
            worksheet.getRange(range.getRow(),range.getCol(),range.getRowCount(), range.getColCount()).getInterior().setColor(Color.FromArgb(argb));
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
