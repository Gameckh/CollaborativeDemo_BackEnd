package GcExcelServer.ServerCommands;

public class RowRange {
    private int firstRow;
    public int getFirstRow() {
        return firstRow;
    }
    public void setFirstRow(int firstRow) {
        this.firstRow = firstRow;
    }

    private int lastRow;
    public int getLastRow() {
        return lastRow;
    }
    public void setLastRow(int lastRow) {
        this.lastRow = lastRow;
    }
}
