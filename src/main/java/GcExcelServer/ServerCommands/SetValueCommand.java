package GcExcelServer.ServerCommands;

import com.grapecity.documents.excel.Workbook;

public class SetValueCommand extends CellCommandBase {
    private Object newValue;
    public Object getNewValue(){
        return this.newValue;
    }
    public void setNewValue(Object newValue){
        this.newValue = newValue;
    }

    @Override
    public void execute(Workbook workbook) {
        workbook.getWorksheets().get(this.getSheetName()).getRange(this.getRow(), this.getColumn()).setValue(this.getNewValue());
    }
}
