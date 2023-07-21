package GcExcelServer.ServerCommands;

public class ColumnRange {
    private int firstCol;
    public int getFirstCol() {
        return firstCol;
    }
    public void setFirstCol(int firstCol) {
        this.firstCol = firstCol;
    }

    private int lastCol;
    public int getLastCol() {
        return lastCol;
    }
    public void setLastCol(int lastCol) {
        this.lastCol = lastCol;
    }
}
