package GcExcelServer.ServerCommands;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;

public class SetFontStyleCommand extends SelectionCommandBase{
	private Object value;
    public Object getValue(){
        return this.value;
    }
    public void setValue(Object value){
        this.value = value;
    }
    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = this.getWorksheet(workbook);
        if(worksheet == null){
            return;
        }
        for(SJSRange range: this.getSelections()){
        	if("italic".equals(value)) {
        		worksheet.getRange(range.getRow(),range.getCol(),range.getRowCount(), range.getColCount()).getFont().setItalic(true);
        	}
        	if("normal".equals(value)) {
        		worksheet.getRange(range.getRow(),range.getCol(),range.getRowCount(), range.getColCount()).getFont().setItalic(false);
        	}
            
        }
    }
}
