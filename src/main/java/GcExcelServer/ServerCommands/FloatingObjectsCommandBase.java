package GcExcelServer.ServerCommands;

public abstract class FloatingObjectsCommandBase extends SheetCommandBase {
    private String[] floatingObjects;

    public String[] getFloatingObjects() {
        return floatingObjects;
    }

    public void setFloatingObjects(String[] floatingObjects) {
        this.floatingObjects = floatingObjects;
    }
}
