package GcExcelServer.ServerCommands;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;

public abstract class SheetCommandBase extends CommandBase{
    private String sheetName;
    public String getSheetName(){
        return this.sheetName;
    }
    public void setSheetName(String value){
        this.sheetName = value;
    }

    public IWorksheet getWorksheet(Workbook workbook){
        return workbook.getWorksheets().get(this.getSheetName());
    }
}
