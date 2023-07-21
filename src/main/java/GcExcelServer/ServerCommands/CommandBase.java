package GcExcelServer.ServerCommands;

import com.grapecity.documents.excel.Workbook;

public abstract class CommandBase {
    private String cmd;
    public String getCmd(){
        return cmd;
    }
    public void setCmd(String cmd){
        this.cmd = cmd;
    }

    private String docID;
    public String getDocID(){
        return docID;
    }
    public void setDocID(String value){
        this.docID = value;
    }

    public abstract void execute(Workbook workbook);
}
