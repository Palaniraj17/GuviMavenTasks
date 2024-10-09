package JavaTask5;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileHandling {

	public static void main(String[] args) throws IOException {

		// Writing the data into the excel
		XSSFWorkbook DetailsFile = new XSSFWorkbook();
		XSSFSheet DetailsSheet = DetailsFile.createSheet("DetailsSheet");
		Object[][] Details = { { "Name", "Age", "Email" }, { "John Doe", "30", "john@test.com" },
				{ "John Doe", "28", "john@test.com" }, { "Bob Smith", "35", "jacky@example.com" },
				{ "Swapnil", "37", "swapnil@example.com" } };

		int recordnum = 0;
		for (Object[] rec1 : Details) {
			XSSFRow record = DetailsSheet.createRow(recordnum++);
			int columnnum = 0;
			for (Object col : rec1) {
				XSSFCell Column = record.createCell(columnnum++);
				if (col instanceof Integer) {

					Column.setCellValue((Integer) col);

				} else if (col instanceof String) {
					Column.setCellValue((String) col);
				}
			}
		}
		FileOutputStream resultfile = new FileOutputStream(
				"D:\\GuviTasks\\JAT_GuviTask_Maven\\src\\main\\java\\JavaTask5\\MemberDetails.xlsx");
		DetailsFile.write(resultfile);
		System.out.println("Details are updated and New excel file created");
		DetailsFile.close();

		// Reading the stored data from the excel
		XSSFWorkbook DetailsRead = new XSSFWorkbook(
				"D:\\GuviTasks\\JAT_GuviTask_Maven\\src\\main\\java\\JavaTask5\\MemberDetails.xlsx");
		XSSFSheet DetailsSheetRead = DetailsFile.getSheet("DetailsSheet");
		int totalrecords = DetailsSheetRead.getLastRowNum();
		int columforecords = DetailsSheetRead.getRow(0).getLastCellNum();
		System.out.println("The saved member details are ");
		for (int i = 1; i <= totalrecords; i++) {
			XSSFRow RecordRead = DetailsSheetRead.getRow(i);
			for (int j = 0; j < columforecords; j++) {
				XSSFCell colrec = RecordRead.getCell(j);
				System.out.println(colrec.getStringCellValue());

			}
			System.out.println("");

		}
		DetailsRead.close();
	}
}
