package GcExcelServer.ServerCommands;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.IShape;

public class MoveFloatingObjectsCommand extends FloatingObjectsCommandBase {
    @Override
    public void execute(Workbook workbook) {
        IWorksheet worksheet = this.getWorksheet(workbook);
        if(worksheet == null){
            return;
        }

        for (String floatingObject: this.getFloatingObjects()){
            IShape shape = worksheet.getShapes().get(floatingObject);
            if(shape == null){
                continue;
            }
            shape.setLeftInPixel(shape.getLeftInPixel() + this.offsetX);
            shape.setTopInPixel(shape.getTopInPixel() + this.offsetY);
        }
    }

    private int offsetX;

    public int getOffsetX() {
        return offsetX;
    }

    public void setOffsetX(int offsetX) {
        this.offsetX = offsetX;
    }

    private int offsetY;

    public int getOffsetY() {
        return offsetY;
    }

    public void setOffsetY(int offsetY) {
        this.offsetY = offsetY;
    }
}
