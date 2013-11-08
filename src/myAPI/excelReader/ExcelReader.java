package myAPI.excelReader;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.zip.*;
import org.jdom2.*;
import org.jdom2.input.SAXBuilder;


/**
 * The ExcelReader can read out specific informations of an Excel File.
 * </br> E.g. (Value of a cell, range, width, height, background color, etc....)
 * @author Burkard Döpfner
 */
public class ExcelReader {
    final String DIR_WORKSHEETS = File.separator + "xl" + File.separator + "worksheets" + File.separator;
    /**
     * 
     * @param ExcelPath The directory path to the excelfile. Has to be an .xslx file. E.g. "test.xlsx"
     * @param InnerExcel_XmlPath The directory path to a XML File that is in the Excel file included. E.g. "/xs/worksheets/sheet1.xml"
     * @return Returns The an spcified xml file which is in the excelfile included. e.g. Sheet1
     * @throws IOException Throws IOException, when the first parameter is wrong. 
     * @throws JDOMException Throws JDOMException, when the second parameter is wrong.
     */
    public static Document unzipXML(String ExcelPath, String InnerExcel_XmlPath) throws IOException, JDOMException{
        // Unzip the excel-file and get the spcified InnerExcel-Xml-File
        ZipFile excelfile = new ZipFile(ExcelPath);
        ZipEntry entry = excelfile.getEntry(InnerExcel_XmlPath);
        InputStream is = excelfile.getInputStream(entry);
        
        Document doc = new SAXBuilder().build(is);
        return doc;

    } 
    
    

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
    }
}
