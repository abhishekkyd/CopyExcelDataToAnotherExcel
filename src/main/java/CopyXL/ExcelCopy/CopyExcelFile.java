package CopyXL.ExcelCopy;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CopyExcelFile {

    public static void main(String[] args) {
	copyExcelFile();
    }

    public static void copyExcelFile() {
	XSSFWorkbook book1 = null;

	try {
	    book1 = new XSSFWorkbook(new FileInputStream("Enrollment.xlsx"));

	    for (int i = 6; i > 0; i--) {
		book1.removeSheetAt(i);
	    }

	    FileOutputStream fileOut = new FileOutputStream(
		    "Enrollment_new.xlsx", false);
	    book1.write(fileOut);
	    fileOut.close();

	} catch (FileNotFoundException e) {
	    e.printStackTrace();
	} catch (IOException e) {
	    e.printStackTrace();
	}
    }
}
