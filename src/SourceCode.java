import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

/*
 
 PORGRAM ODPOWIEDZIALNY ZA TWORZENIE PLIKOW HTML Z ALARMAMI
 KAZDY ALARM JEST GENEROWANY WEDLUG JEDNEGO SCHEMATU
 KAZDY ALARM JEST ZAPISANY W OSOBNYM PLIKU HTML,
 KTORE TRZEBA UMIESCIC W ODPOWIEDNIM KATALOGU NA KARCIE SF- SINUMERIK
 HTML'E ZAPISYWANE SA NA LOKALNYM KOMPUTERZE 
 
* NIE MA MOZLIWOSCI WYBORU JEZYKOW
* NIE MA MOZLIWOSCI WYBORU DLA JAKIEJ MASZYNY MA POBRAC
* DANE ODCZYTYWANE SA Z PLIKU XLS ZNAJDUJACEGO SIE NA LOKALNEJ MASZYNIE
* NIE MA WYGENEROWANEGO JAR/EXE


* ELEMENTY DO ZROBIENIA NA PRZYSZLOSC
 
 */

class Parameters {
	
	public static final String PATH_TO_SAVE = "C:\\Users\\el08\\Desktop\\programiki\\";
	
}


public class SourceCode {
    //public static final String SAMPLE_XLSX_FILE_PATH = "C:\\Users\\el08\\Desktop\\proba2.xls";
  //  public static final String Snazwa = "C:\\Users\\el08\\Desktop\\proba2.xls";


    public static void RUN(String SAMPLE_XLSX_FILE_PATH,String DirFile) throws EncryptedDocumentException, InvalidFormatException, IOException {
   	ArrayList<obiekt> obj = new ArrayList<obiekt>();
    	

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");



        // 1. You can obtain a sheetIterator and iterate over it
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        System.out.println("Retrieving Sheets using Iterator");
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            System.out.println("=> " + sheet.getSheetName());
        }

        /*
           ==================================================================
           Iterating over all the rows and columns in a Sheet (Multiple ways)
           ==================================================================
        */

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

        // 1. You can obtain a rowIterator and columnIterator and iterate over them
        System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
        Iterator<Row> rowIterator = sheet.rowIterator();
        int x = 0;
        while (rowIterator.hasNext()) {
        	obiekt ob = new obiekt();
            Row row = rowIterator.next();

            // Now let's iterate over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\n");
                
                /*
                 x ==0, x== 2 itp. chodzi tutaj o kolumny z danymi w zasysanym dokumencie
                 w tym przypadku dokument rozszerzony jest o jakies niepotrzebne informacje,
                 ktore trzeba przeskoczyc ( dlatego nie ma x ==1)
                 */
                
//                if(x == 0)
//                	ob.numer = cellValue;
//                else if(x == 2)
//                	ob.komentarz = cellValue;
//                else if(x == 4)
//                	ob.reason = cellValue;
//                else if(x == 5)
//                	ob.solution = cellValue;
                
                if(x == 0)
                	ob.numer = cellValue;
                if(x == 1)
                	ob.reason = cellValue;
                else if(x == 3)
                	ob.solution = cellValue;
                else if(x == 2)
                	ob.komentarz = cellValue;
                
                x++;

            }
            obj.add(ob);
            x = 0;
 
            
            
        }
        createDir(DirFile);
        for(int i = 0;i < obj.size();i++) {
        		HTML(obj.get(i).numer,obj.get(i).reason, obj.get(i).solution, obj.get(i).komentarz, DirFile);
        }
        
        // Closing the workbook
        workbook.close();
    }
    
    
    public static void HTML(String numer, String powod, String rozwiazanie, String komenatrz,String DirFile) {
    	
        StringBuilder htmlBuilder = new StringBuilder();
        htmlBuilder.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?><!DOCTYPE html PUBLIC \"-//W3C//DTD");
        htmlBuilder.append("HTML 4.0 Transitional//EN\" >");
        htmlBuilder.append("<html>");
        htmlBuilder.append("<head><title></title></head>");
        htmlBuilder.append("<body>");
        htmlBuilder.append("<table>");
        htmlBuilder.append("<tr>");
        htmlBuilder.append("- <td width=\"15%\">");
        htmlBuilder.append("<b><a name=\"510000\">"+numer+"</a></b>");
        htmlBuilder.append("</td>");
        htmlBuilder.append("<td width=\"85%\">");
        htmlBuilder.append("<b>This is help for user alarm nr:"+numer+" </b>");
        htmlBuilder.append("</td>");
        htmlBuilder.append("</tr>");
        htmlBuilder.append("<tr>");
        htmlBuilder.append("<td valign=\"top\" width=\"15%\">");
        htmlBuilder.append("<b>Explanation</b>");
        htmlBuilder.append("</td>");
        htmlBuilder.append("<td width=\"85%\"> "+powod+"</td>");
        htmlBuilder.append("</tr>");
        htmlBuilder.append("<tr>");
        htmlBuilder.append("<td valign=\"top\" width=\"15%\"><b>Remedy:</b></td>");
        htmlBuilder.append("<td width=\"85%\">"+rozwiazanie+"</td>");
        htmlBuilder.append("</tr>");
        htmlBuilder.append("<tr>");
        htmlBuilder.append("<td valign=\"top\" width=\"15%\"><b>Comment:</b></td>");
        htmlBuilder.append("<td width=\"85%\">"+komenatrz+"</td>");
        htmlBuilder.append("</tr>");
        htmlBuilder.append("</table>");
        htmlBuilder.append("</body>");
        htmlBuilder.append("</html>");

        String html = htmlBuilder.toString();
        

        File s_file = new File(Parameters.PATH_TO_SAVE +DirFile +"\\"+numer+".html");
    
        try {
        	BufferedWriter bw = new BufferedWriter (new FileWriter(s_file));
        	bw.write(html);
        	bw.close();
        }
        catch (IOException e){
        	e.printStackTrace();
        }
    }
    
    public static void createDir(String nazwa) {
        File DirFile = new File(Parameters.PATH_TO_SAVE + nazwa);
     // if the directory does not exist, create it
        if (!DirFile.exists()) {
            boolean result = false;

            try{
            	DirFile.mkdir();
                result = true;
            } 
            catch(SecurityException se){
                //handle it
            }        
            if(result) {    
                System.out.println("DIR created");  
            }
        }
    	
    }
}