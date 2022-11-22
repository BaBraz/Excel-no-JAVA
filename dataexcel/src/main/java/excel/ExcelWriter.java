package excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {
	private static String[] columns = { "First Name", "Last Name", "Email",
    "Date Of Birth" };
  private static List<Contact> contacts = new ArrayList<Contact>();

  public static void main(String[] args) throws IOException, 
	InvalidFormatException {
    contacts.add(new Contact("Bruno", "Saurel","sylvain.saurel@gmail.com", "17/01/2000"));
    contacts.add(new Contact("Tatiana", "Dupond","sylvain.saurel@gmail.com", "17/08/1989"));
    contacts.add(new Contact("WeldelPierre", "Dupont","sylvain.saurel@gmail.com", "17/07/1956"));
    contacts.add(new Contact("Emanuela", "Diaz", "sylvain.saurel@gmail.com","17/05/1988"));

    Workbook workbook = new XSSFWorkbook();
    
    Sheet sheet = workbook.createSheet("Contacts");

    Row headerRow = sheet.createRow(0);

    for (int i = 0; i < columns.length; i++) {
      Cell cell = headerRow.createCell(i);
      cell.setCellValue(columns[i]);
     
    }
    int rowNum = 1;

    for (Contact contact : contacts) {
      Row row = sheet.createRow(rowNum++);
      row.createCell(0).setCellValue(contact.firstName);
      row.createCell(1).setCellValue(contact.lastName);
      row.createCell(2).setCellValue(contact.email);
      row.createCell(3).setCellValue(contact.dateOfBirth);
    }

    for (int i = 0; i < columns.length; i++) {
      sheet.autoSizeColumn(i);
    }

    FileOutputStream fileOut = new FileOutputStream("contacts.xlsx");
    workbook.write(fileOut);
    fileOut.close();
    workbook.close();
  }

}
