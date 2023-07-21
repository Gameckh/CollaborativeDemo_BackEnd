package GcExcelServer.ServerCommands;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.UnderlineType;
import com.grapecity.documents.excel.Workbook;

public class SetDoubleUnderlineCommand extends SelectionCommandBase{
	private Boolean value;
    public Boolean getValue(){
        return this.value;
    }
    public void setValue(Boolean value){
        this.value = value;
    }
    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = this.getWorksheet(workbook);
        if(worksheet == null){
            return;
        }
        for(SJSRange range: this.getSelections()){
        	if(value) {
        		worksheet.getRange(range.getRow(),range.getCol(),range.getRowCount(), range.getColCount()).getFont().setUnderline(UnderlineType.Double);
        	} else {
        		worksheet.getRange(range.getRow(),range.getCol(),range.getRowCount(), range.getColCount()).getFont().setUnderline(UnderlineType.None);
        	}
        	
            
        }
    }
}
