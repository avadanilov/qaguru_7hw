package tests;

import com.codeborne.pdftest.PDF;
import net.lingala.zip4j.ZipFile;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.junit.jupiter.api.Test;
import org.xlsx4j.org.apache.poi.ss.usermodel.DataFormatter;
import org.xlsx4j.sml.Cell;
import org.xlsx4j.sml.Row;
import org.xlsx4j.sml.SheetData;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertTrue;

public class FilesTests {

    @Test
    void testDocx() throws Docx4JException {
        ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
        InputStream stream = classLoader.getResourceAsStream("testDocx.docx");
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(stream);
        assertTrue(wordMLPackage.getMainDocumentPart().getContent().toString().contains("Hello, world!"));
    }

    @Test
    void testXlsx() throws Exception {
        ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
        InputStream stream = classLoader.getResourceAsStream("testXlsx.xlsx");
        SpreadsheetMLPackage excelMLPackage = SpreadsheetMLPackage.load(stream);
        SheetData data = excelMLPackage.getWorkbookPart().getWorksheet(0).getContents().getSheetData();
        DataFormatter formatter = new DataFormatter();
        ArrayList formattedStringList = new ArrayList();

        for (Row r : data.getRow()) {
            for (Cell c : r.getC()) {
                String text = formatter.formatCellValue(c);
                formattedStringList.add(text);
            }
        }
        assertTrue(formattedStringList.toString().contains("Hello, world xlsx!"));
    }

    @Test
    void testPdf() throws IOException {
        ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
        URL urlPdf = classLoader.getResource("testPdf.pdf");
        PDF parsedPdf = new PDF(urlPdf);
        assertTrue((parsedPdf.text).contains("Hello, world pdf!"));
    }


    @Test
    void testZip() throws IOException {
        String filesExtractPath = "/files_from_archive";
        String arcPassword = "12345";

        ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
        URL urlFile = classLoader.getResource("testZip.zip");
        String filePath = urlFile.getPath();
        ZipFile zipFile = new ZipFile(filePath, arcPassword.toCharArray());
        zipFile.extractAll(filesExtractPath);

        File dir = new File(filesExtractPath);
        File[] arrFiles = dir.listFiles();
        List<File> filesList = Arrays.asList(arrFiles);
        assertTrue(filesList.toString().contains("Hello zip.txt"));
    }
}
