package GcExcelServer.ServerCommands;

public abstract class SelectionCommandBase extends SheetCommandBase {

    private SJSRange[] selections;

    public SJSRange[] getSelections() {
        return selections;
    }

    public void setSelections(SJSRange[] selections) {
        this.selections = selections;
    }
}
