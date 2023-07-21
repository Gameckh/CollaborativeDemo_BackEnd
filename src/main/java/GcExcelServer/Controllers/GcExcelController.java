package GcExcelServer.Controllers;

import GcExcelChartExtractor.ChartExtractor;
import GcExcelChartExtractor.ChartInfo;
import GcExcelServer.ServerCommands.*;
import GcExcelServer.WebSocketConfig;
import com.grapecity.documents.excel.*;
import com.grapecity.documents.excel.drawing.*;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Array;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.*;
import java.io.InputStream;
import java.util.concurrent.ConcurrentHashMap;


@RestController
@CrossOrigin
@RequestMapping("/api/gcexcel")
public class GcExcelController {

//    private SpreadJSSocketHandler handler;
    private static ConcurrentHashMap<String, Workbook> _spreadCaches = new ConcurrentHashMap<String, Workbook>();
    private static File DOCS_FOLDER = null;

    static {
        String docsPath = GcExcelController.class.getClassLoader().getResource("documents").getPath();
        DOCS_FOLDER = new File(docsPath);
    }

    public GcExcelController(WebSocketConfig config){
//        this.handler = config.getSpreadJSSocketHandler();
    }

    @RequestMapping(value = "/docs", method = RequestMethod.GET)
    public ArrayList<String> getDocs(HttpServletResponse response) {
        response.setStatus(HttpServletResponse.SC_OK);
        return getDocsList(true);
    }

    @RequestMapping(value = "/xlsx/{name}", method = RequestMethod.PUT)
    public void create(@PathVariable String name,HttpServletResponse response) {
        ArrayList<String> docsList = this.getDocsList(false);
        if(docsList.contains(name)){
            response.setStatus(HttpServletResponse.SC_CONFLICT);
        }else{
            Workbook workbook = new Workbook();
            workbook.save(DOCS_FOLDER.getAbsolutePath() + name + ".xlsx");
        }
    }
    @RequestMapping(value = "/xlsx/{name}", method = RequestMethod.POST)
    public void openExcel(@PathVariable String name, HttpServletRequest request, HttpServletResponse response) {
        try {
            Workbook workbook = this.getDocument(name);
            ArrayList<String> docsList = this.getDocsList(false);
            if (workbook == null) {
                response.setStatus(HttpServletResponse.SC_NOT_FOUND);
            } else {
                response.setStatus(HttpServletResponse.SC_OK);
            }

        } catch (Exception ex) {
            ex.printStackTrace();
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
        }
    }

    @RequestMapping(value = "/xlsx", method = RequestMethod.POST)
    public String openExcel(HttpServletRequest request, HttpServletResponse response) {
        try {
            Workbook workbook = new Workbook();
            String id = UUID.randomUUID().toString();
            _spreadCaches.put(id, workbook);

            workbook.open(request.getInputStream());
            response.setStatus(HttpServletResponse.SC_OK);
            return id;
        } catch (Exception ex) {
            ex.printStackTrace();
        }

        response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
        return null;
    }

    @RequestMapping(value = "/{id}/xlsx", method = RequestMethod.GET)
    public void toXlsx(HttpServletRequest request, @PathVariable String id,HttpServletResponse response) {
        try {
            Workbook workbook = this.getDocument(id);;
            if(workbook == null){
                response.setStatus(HttpServletResponse.SC_NOT_FOUND);
                return;
            }

            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            workbook.save(byteArrayOutputStream);

            // Set the content type and attachment header.
            response.addHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.addHeader("Content-Disposition", "inline;filename=" + id + ".xlsx;");

            // Copy the stream to the response's output stream.
            response.getOutputStream().write(byteArrayOutputStream.toByteArray());
            response.flushBuffer();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }


    @RequestMapping(value = "/json", method = RequestMethod.POST)
    public String fromJson(HttpServletRequest request, HttpServletResponse response) {
        try {
            Workbook workbook = new Workbook();
            String id = UUID.randomUUID().toString();
            _spreadCaches.put(id, workbook);

            workbook.fromJson(request.getInputStream());
            response.setStatus(HttpServletResponse.SC_OK);
            return id;
        } catch (Exception ex) {
            ex.printStackTrace();
        }

        response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
        return null;
    }

	@RequestMapping(value = "/{id}/json", method = RequestMethod.GET)
	public void toJson(HttpServletRequest request, @PathVariable String id,HttpServletResponse response) {
        try {
            Workbook workbook = this.getDocument(id);;
            if(workbook == null){
                response.setStatus(HttpServletResponse.SC_NOT_FOUND);
                return;
            }

            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            workbook.toJson(byteArrayOutputStream);

            // Set the content type and attachment header.
            response.addHeader("Content-Type", "application/json");
            response.addHeader("Content-Disposition", "inline;filename=gcexcel-exported.json;");

            // Copy the stream to the response's output stream.
            response.getOutputStream().write(byteArrayOutputStream.toByteArray());
            response.flushBuffer();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    @RequestMapping(value = "/{id}/pdf", method = RequestMethod.GET)
    public void toPdf(HttpServletRequest request, @PathVariable String id,HttpServletResponse response) {
        try {
            Workbook workbook = this.getDocument(id);;
            if(workbook == null){
                response.setStatus(HttpServletResponse.SC_NOT_FOUND);
                return;
            }

            ArrayList<String> chartImages = null;

            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            workbook.save(byteArrayOutputStream, SaveFileFormat.Pdf);

            // Set the content type and attachment header.
            response.addHeader("Content-Type", "application/pdf");
            response.addHeader("Content-Disposition", "inline;filename=gcexcel-exported.pdf;");

            // Copy the stream to the response's output stream.
            response.getOutputStream().write(byteArrayOutputStream.toByteArray());
            response.flushBuffer();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }


    @RequestMapping(value = "/{id}/csv", method = RequestMethod.GET)
    public void toCsv(HttpServletRequest request, @PathVariable String id,HttpServletResponse response) {
        try {
            Workbook workbook = this.getDocument(id);;
            if(workbook == null){
                response.setStatus(HttpServletResponse.SC_NOT_FOUND);
                return;
            }

            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            workbook.save(byteArrayOutputStream, SaveFileFormat.Csv);

            // Set the content type and attachment header.
            response.addHeader("Content-Type", "application/csv");
            response.addHeader("Content-Disposition", "inline;filename=gcexcel-exported.csv;");

            // Copy the stream to the response's output stream.
            response.getOutputStream().write(byteArrayOutputStream.toByteArray());
            response.flushBuffer();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    @RequestMapping(value = "/{id}/image", method = RequestMethod.GET)
    public void toImage(HttpServletRequest request, @PathVariable String id,HttpServletResponse response) {
        try {
            Workbook workbook = this.getDocument(id);;
            if(workbook == null){
                response.setStatus(HttpServletResponse.SC_NOT_FOUND);
                return;
            }


            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            workbook.getActiveSheet().toImage(byteArrayOutputStream, ImageType.PNG);


            // Set the content type and attachment header.
            response.addHeader("Content-Type", "image/png");
            response.addHeader("Content-Disposition", "inline;filename=gcexcel-exported.png;");

            // Copy the stream to the response's output stream.
            response.getOutputStream().write(byteArrayOutputStream.toByteArray());
            response.flushBuffer();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    @RequestMapping(value = "/{id}/sheets", method = RequestMethod.GET)
    public ArrayList<String> getSheets(HttpServletRequest request,
                               @PathVariable String id,
                               HttpServletResponse response) {
        try {
            Workbook workbook = this.getDocument(id);;
            if(workbook == null){
                response.setStatus(HttpServletResponse.SC_NOT_FOUND);
                return null;
            }

            ArrayList<String> sheets = new ArrayList<String>();
            for(IWorksheet sheet : workbook.getWorksheets()){
                sheets.add(sheet.getName());
            }
            response.setStatus(HttpServletResponse.SC_OK);
            return sheets;
        } catch (Exception ex) {
            ex.printStackTrace();

            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
        }
        return null;
    }

    @RequestMapping(value = "/{id}/value", method = RequestMethod.GET)
    public String getCellValue(HttpServletRequest request,
                             @PathVariable String id,
                               @RequestBody SetValueCommand value,
                             HttpServletResponse response) {
        try {
            Workbook workbook = this.getDocument(id);;
            if(workbook == null){
                response.setStatus(HttpServletResponse.SC_NOT_FOUND);
                return null;
            }

            String text = workbook.getWorksheets().get(value.getSheetName()).getRange(value.getRow(), value.getColumn()).getText();

            response.setStatus(HttpServletResponse.SC_OK);
            return text;
        } catch (Exception ex) {
            ex.printStackTrace();

            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
        }
        return "";
    }

    @RequestMapping(value = "/editCell", method = RequestMethod.POST)
    public void setCellValue(HttpServletRequest request,
                             @RequestBody SetValueCommand setValueCommand,
                             HttpServletResponse response) {

        this.executeCommand(setValueCommand, response);
    }

    @RequestMapping(value = "/resizeRow", method = RequestMethod.POST)
    public void resizeRow(HttpServletRequest request,
                             @RequestBody ResizeRowCommand resizeRowCommand,
                             HttpServletResponse response) {

        this.executeCommand(resizeRowCommand, response);
    }

    @RequestMapping(value = "/resizeColumn", method = RequestMethod.POST)
    public void resizeRow(HttpServletRequest request,
                          @RequestBody ResizeColumnCommand resizeColumnCommand,
                          HttpServletResponse response) {

        this.executeCommand(resizeColumnCommand, response);
    }

    @RequestMapping(value = "/setFontFamily", method = RequestMethod.POST)
    public void setFontFamily(HttpServletRequest request,
                          @RequestBody SetFontFamilyCommand setFontFamilyCommand,
                          HttpServletResponse response) {

        this.executeCommand(setFontFamilyCommand, response);
    }

    @RequestMapping(value = "/setFontSize", method = RequestMethod.POST)
    public void setFontSize(HttpServletRequest request,
                              @RequestBody SetFontSizeCommand setFontSizeCommand,
                              HttpServletResponse response) {

        this.executeCommand(setFontSizeCommand, response);
    }

    @RequestMapping(value = "/setBackColor", method = RequestMethod.POST)
    public void setBackColor(HttpServletRequest request,
                            @RequestBody SetBackColorCommand setBackColorCommand,
                            HttpServletResponse response) {

        this.executeCommand(setBackColorCommand, response);
    }

    @RequestMapping(value = "/setForeColor", method = RequestMethod.POST)
    public void setForeColor(HttpServletRequest request,
                            @RequestBody SetForeColorCommand setForeColorCommand,
                            HttpServletResponse response) {

        this.executeCommand(setForeColorCommand, response);
    }

    @RequestMapping(value = "/moveFloatingObjects", method = RequestMethod.POST)
    public void moveFloatingObjects(HttpServletRequest request,
                             @RequestBody MoveFloatingObjectsCommand movingFloatingObjectsCommand,
                             HttpServletResponse response) {

        this.executeCommand(movingFloatingObjectsCommand, response);
    }

    @RequestMapping(value = "/resizeFloatingObjects", method = RequestMethod.POST)
    public void resizeFloatingObjects(HttpServletRequest request,
                                      @RequestBody ResizeFloatingObjectsCommand resizingFloatingObjectsCommand,
                                      HttpServletResponse response) {

        this.executeCommand(resizingFloatingObjectsCommand, response);
    }

    @RequestMapping(value = "/insertColumns", method = RequestMethod.POST)
    public void insertColumns(HttpServletRequest request,
                                      @RequestBody InsertColumnsCommand insertColumnsCommand,
                                      HttpServletResponse response) {

        this.executeCommand(insertColumnsCommand, response);
    }

    @RequestMapping(value = "/insertRows", method = RequestMethod.POST)
    public void insertRows(HttpServletRequest request,
                                      @RequestBody InsertRowsCommand insertRowsCommand,
                                      HttpServletResponse response) {

        this.executeCommand(insertRowsCommand, response);
    }
    
    @RequestMapping(value = "/setFontWeight", method = RequestMethod.POST)
    public void setFontWeight(HttpServletRequest request,
                                      @RequestBody SetFontWeightCommand setFontWeightCommand,
                                      HttpServletResponse response) {

        this.executeCommand(setFontWeightCommand, response);
    }
    
    @RequestMapping(value = "/setUnderline", method = RequestMethod.POST)
    public void setUnderline(HttpServletRequest request,
                                      @RequestBody SetUnderlineCommand setUnderlineCommand,
                                      HttpServletResponse response) {

        this.executeCommand(setUnderlineCommand, response);
    }
    
    @RequestMapping(value = "/setDoubleUnderline", method = RequestMethod.POST)
    public void setDoubleUnderline(HttpServletRequest request,
                                      @RequestBody SetDoubleUnderlineCommand setDoubleUnderlineCommand,
                                      HttpServletResponse response) {

        this.executeCommand(setDoubleUnderlineCommand, response);
    }
    
    

    private Workbook getDocument(String id) {
        Workbook workbook = null;
        if (_spreadCaches.containsKey(id)) {
            workbook = _spreadCaches.get(id);
        } else {
            ArrayList<String> docs = this.getDocsList(false);
            if (docs.contains(id)) {
                workbook = new Workbook();
                workbook.open(this.getResourceStream(DOCS_FOLDER.getName() + "/" + id + ".xlsx"));
                _spreadCaches.put(id, workbook);
            }
        }
        return workbook;
    }

    private void executeCommand(CommandBase commandBase, HttpServletResponse response) {
        try {

            Workbook workbook = null;
            if (_spreadCaches.containsKey(commandBase.getDocID())) {
                workbook = _spreadCaches.get(commandBase.getDocID());
            } else {
                response.setStatus(HttpServletResponse.SC_NOT_FOUND);
            }

            commandBase.execute(workbook);
            response.setStatus(HttpServletResponse.SC_OK);
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }



    public InputStream getResourceStream(String resource) {
        return GcExcelController.class.getClassLoader().getResourceAsStream(resource);
    }

    private ArrayList<String> getDocsList(Boolean withExtension){
        File[] listOfFiles = DOCS_FOLDER.listFiles();

        ArrayList<String> docs = new ArrayList<>();
        for (int i = 0; i < listOfFiles.length; i++) {
            if (listOfFiles[i].isFile()) {
                String name = listOfFiles[i].getName();
                if(withExtension){
                    docs.add(name);
                }else {
                    docs.add(name.substring(0, name.lastIndexOf(".")));
                }
            } else if (listOfFiles[i].isDirectory()) {

            }
        }
        return docs;
    }



}
