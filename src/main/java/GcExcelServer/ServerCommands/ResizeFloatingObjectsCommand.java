package GcExcelServer.ServerCommands;

import com.grapecity.documents.excel.IWorksheet;
import com.grapecity.documents.excel.Workbook;
import com.grapecity.documents.excel.drawing.IShape;

public class ResizeFloatingObjectsCommand extends MoveFloatingObjectsCommand {

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
            shape.setLeftInPixel(shape.getLeftInPixel() + this.getOffsetX());
            shape.setTopInPixel(shape.getTopInPixel() + this.getOffsetY());
            shape.setWidthInPixel(shape.getWidthInPixel() + this.getOffsetWidth());
            shape.setHeightInPixel(shape.getHeightInPixel() + this.getOffsetHeight());
        }
    }

    private int offsetWidth;

    public int getOffsetWidth() {
        return offsetWidth;
    }

    public void setOffsetWidth(int offsetWidth) {
        this.offsetWidth = offsetWidth;
    }

    private int offsetHeight;

    public int getOffsetHeight() {
        return offsetHeight;
    }

    public void setOffsetHeight(int offsetHeight) {
        this.offsetHeight = offsetHeight;
    }
}
