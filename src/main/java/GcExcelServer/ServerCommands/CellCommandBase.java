package GcExcelServer.ServerCommands;

public abstract class CellCommandBase extends SheetCommandBase{
    private String range;
    public String getRange(){
        return this.range;
    }
    public void setRange(String range){
        this.range = range;
    }

    private int row;
    public int getRow(){
        return row;
    }
    public void setRow(int value){
        this.row = value;
    }

    private int column;
    public int getColumn(){
        return this.column;
    }
    public void setColumn(int value){
        this.column = value;
    }
}
