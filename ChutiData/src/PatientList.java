/**
 * @author Danindu
 *Date: 2020 06 16
 *Desc.: A class for a patient list object that includes the following methods:
 *		processWorkbook(String filename) : processes the patients in an inputted xls file
 *		writeToExcel() : A method that outputs the xls file with the analyzed data (each patient in one column and number of completes in the other column)
 *		A main method that process the required workbooks and outputs the final data
 */

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class PatientList {
	private Patient[] list;
	private int maxSize;
	private int size;
	public PatientList() {
		this.maxSize = 300;
		list = new Patient [maxSize];
		size = 0;
	}
	//method that reads an excel file and appends the data to the patient list
	public void processWorkbook(String fileName) throws EncryptedDocumentException, IOException {
		String nextCellValue;
		Workbook workbook = WorkbookFactory.create(new File(fileName));

		//		Iterator<Sheet> sheetIterator = workbook.sheetIterator();
		//        System.out.println("Retrieving Sheets using Iterator");
		//        while (sheetIterator.hasNext()) {
		//            Sheet sheet = sheetIterator.next();
		//            System.out.println("=> " + sheet.getSheetName());
		//        }
		Sheet sheet = workbook.getSheetAt(0);
		System.out.println("Processing " + sheet.getSheetName() + " ...");
		DataFormatter dataFormatter = new DataFormatter();
		//loop through rows
		int rowIndex = 1;
		while (sheet.getRow(rowIndex)!=null) {
			//creating a row for the looped index
			Row row = sheet.getRow(rowIndex);
			if (row != null) {  //if row has stuff
				Cell idCell = row.getCell(2);  //get the value in the id column
				String idCellValue = dataFormatter.formatCellValue(idCell);			
				if(idCellValue.trim()!="") {
					if(this.idLinearSearch(idCellValue)<0) {
						Patient p = new Patient(idCellValue);
						int span = 1;
						int completes = 0;
						int patStartRow = rowIndex;
						if(sheet.getRow(rowIndex+1) != null) {
							rowIndex++;  //move onto the patient's 2nd row
						}
						idCellValue = dataFormatter.formatCellValue(sheet.getRow(rowIndex).getCell(2));  //get value of current cell
						while(idCellValue.trim().equals("")) {
							span++;  //increase span by 1
							if((sheet.getRow(rowIndex+1) != null) && (dataFormatter.formatCellValue(sheet.getRow(rowIndex+1).getCell(2)).trim().equals(""))) {
								rowIndex++;  //next row
							}
							else {
								break;
							}
							//rowIndex++;  //next row
							idCellValue = dataFormatter.formatCellValue(sheet.getRow(rowIndex).getCell(2));
						}

						int patEndRow = rowIndex;
						p.setSpan(span);
						for (int i = patStartRow; i <= patEndRow; i++) {
							nextCellValue = dataFormatter.formatCellValue(sheet.getRow(i).getCell(3));
							if((nextCellValue.trim().equals("")) ==false) {
								completes++;
							}
							else {
								continue;
							}
						}
						p.setCompletes(completes);
						this.Insert(p);
					}
					else {
						int index = this.idLinearSearch(idCellValue);
						Patient p = this.getList()[index];

						int span = 1;
						int completes = 0;
						int patStartRow = rowIndex;
						if(sheet.getRow(rowIndex+1) != null) {
							rowIndex++;  //move onto the patient's 2nd row
						}
						idCellValue = dataFormatter.formatCellValue(sheet.getRow(rowIndex).getCell(2));  //get value of current cell
						while(idCellValue.trim().equals("")) {
							span++;  //increase span by 1
							if((sheet.getRow(rowIndex+1) != null) && (dataFormatter.formatCellValue(sheet.getRow(rowIndex+1).getCell(2)).trim().equals(""))) {
								rowIndex++;  //next row
							}
							else {
								break;
							}
							//rowIndex++;  //next row
							idCellValue = dataFormatter.formatCellValue(sheet.getRow(rowIndex).getCell(2));
						}

						int patEndRow = rowIndex;
						p.setSpan(p.getSpan() + span);
						for (int i = patStartRow; i <= patEndRow; i++) {
							nextCellValue = dataFormatter.formatCellValue(sheet.getRow(i).getCell(3));
							if((nextCellValue.trim().equals("")) ==false) {
								completes++;
							}
							else {
								continue;
							}
						}
						p.setCompletes(p.getCompletes() + completes);
					}
				}


				//				if(idCellValue.trim()!="") {
				//					System.out.println(idCellValue);
				//				}
			}
			rowIndex++;
		}
	}
	//method that writes final data to an excel sheet
	public void writeToExcel () throws IOException {
		Workbook workbook = new HSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file		
		/* CreationHelper helps us create instances of various things like DataFormat, 
        Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
		//CreationHelper createHelper = workbook.getCreationHelper();

		// Create a Sheet
		Sheet sheet = workbook.createSheet("Total Completes");

		// Create a Font for styling header cells and for data
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.RED.getIndex());

		Font normalFont = workbook.createFont();
		normalFont.setBold(false);
		normalFont.setFontHeightInPoints((short) 14);
		normalFont.setColor(IndexedColors.BLACK.getIndex());

		// Create a CellStyle with the font
		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);

		// Create a CellStyle with the font
		CellStyle normalCellStyle = workbook.createCellStyle();
		normalCellStyle.setFont(normalFont);

		// Create a Row
		Row headerRow = sheet.createRow(0);

		Cell cell = headerRow.createCell(0);
		cell.setCellValue("ID");
		cell.setCellStyle(headerCellStyle);
		cell = headerRow.createCell(1);
		cell.setCellStyle(headerCellStyle);
		cell.setCellValue("# of Completes");

		for (int i = 0; i < this.size; i++) {
			Row row = sheet.createRow(i+1);
			//print id
			cell = row.createCell(0);
			cell.setCellValue(this.getList()[i].getId());
			cell.setCellStyle(normalCellStyle);
			//print completes
			cell = row.createCell(1);
			cell.setCellValue(this.getList()[i].getCompletes());
			cell.setCellStyle(normalCellStyle);		
		}

		// Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("FINAL.xls");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
		
	}

	public boolean Insert (Patient p) {
		if (size<maxSize) {
			list[size] = p;  //adding the record to the list at next available index
			size++;			//increasing the size by 1
			return true;
		}
		return false;
	}

	public int idLinearSearch (String idToFind) {
		for (int i = 0; i < size; i++) {
			if(list[i].getId().equals(idToFind)) {
				return i;
			}
		}
		return -1;
	}

	//***SETTTERS AND GETTERS***
	public Patient[] getList() {
		return list;
	}

	public void setList(Patient[] list) {
		this.list = list;
	}

	public int getMaxSize() {
		return maxSize;
	}

	public void setMaxSize(int maxSize) {
		this.maxSize = maxSize;
	}

	public int getSize() {
		return size;
	}

	public void setSize(int size) {
		this.size = size;
	}
	//temporary output method
	public void output () {
		for (int i = 0; i < this.getSize(); i++) {
			System.out.println(this.getList()[i].toString());
		}
	}

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		// TODO Auto-generated method stub
		PatientList pList = new PatientList();
		pList.processWorkbook("FLUID.xls");
		pList.processWorkbook("HEMO.xls");		
		pList.processWorkbook("PELOD.xls");
		pList.processWorkbook("VASO.xls");
		pList.output();
		try {
			pList.writeToExcel();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		System.out.println("Written to Excel.");
	}

}

