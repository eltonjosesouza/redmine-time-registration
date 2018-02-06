package main;
import java.awt.Container;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;

import javax.swing.BorderFactory;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JProgressBar;
import javax.swing.border.Border;

import org.apache.poi.ss.usermodel.*;

import com.taskadapter.redmineapi.RedmineException;
import com.taskadapter.redmineapi.RedmineManager;
import com.taskadapter.redmineapi.RedmineManagerFactory;
import com.taskadapter.redmineapi.TimeEntryManager;
import com.taskadapter.redmineapi.bean.TimeEntry;
import com.taskadapter.redmineapi.bean.TimeEntryFactory;

public class Main {

	private static String apiAccessKey;
	private static String url;
	private final static String filePath = 
			System.getProperty("user.home") + File.separator + 
			"Documents" + File.separator + 
			"Redmine Time Registration" + File.separator + 
			"Time Sheet.xlsx";
	private static RedmineManager redmineManager;
	private static TimeEntryManager timeEntryManager;
	private static Sheet sheet;
	private static CellStyle cellStyleOK, cellStyleNOK, cellStyleNULL;
	private static Workbook workbook;
	private static JFrame frame = null;
	private static HashMap<String, Integer> activitiesTypeEntries = new HashMap<>();
	private static JProgressBar progress;
	private static Border border;

	/**
	 * Run application
	 * @param args 
	 * @throws IOException 
	 */
	public static void main(String[] args) throws IOException {
		FileOutputStream fileOut = null;

		//Creating the Frame
		frame = new JFrame("Redmine Time Registration");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		Container content = frame.getContentPane();

		//Creating the Progress Bar
		int min = 0;
		int max = 100;
		progress = new JProgressBar(min, max);
		progress.setStringPainted(true);
		progress.setVisible(true);
		
		//Adding the progress bar in the frame
		content.add(progress);

		//Updating the progress
		progress.setValue(0);
		border = BorderFactory.createTitledBorder("Initializing...");
		progress.setBorder(border);

		//Finish setting up the frame and showing it
		frame.setSize(300, 100);
		frame.setLocationRelativeTo(null);
		frame.setVisible(true);

		try {

			//Creating the file of the sheet
			File file = new File(filePath);

			//Trying to open the file
			//I do this because I need to certify that the sheet is closed
			//If not an Exception is thrown and a error message appears to the user
			fileOut = new FileOutputStream(file, true);
			fileOut.close();

			//Opening the file and creating the workbook
			workbook = WorkbookFactory.create(new FileInputStream(filePath));

			//Updating the progress
			border = BorderFactory.createTitledBorder("Getting Parameters...");
			progress.setBorder(border);
			progress.setValue(10);

			//Gettin user parameters, like url of Redmine and API Key
			if(!getParameters()) {//If there is no parameter I show the error message and end the application
				String message = "There were some error trying to read the parameters in the Excel File!\n\n"
						+ "Open the sheet, go to the first tab and set the configuration data!";

				showErrorMessage(message);

				return;
			}

			//Creating objects from Redmine
			redmineManager = RedmineManagerFactory.createWithApiKey(url, apiAccessKey);
			timeEntryManager = redmineManager.getTimeEntryManager();

			border = BorderFactory.createTitledBorder("Reading Entries and Saving on Redmine...");
			progress.setBorder(border);
			progress.setValue(30);

			//Call the method thats handles the real action
			saveTimeEntries();

			border = BorderFactory.createTitledBorder("Saving Excel File...");
			progress.setBorder(border);
			progress.setValue(90);

			// Write the output to the file
			fileOut = new FileOutputStream(file);
			workbook.write(fileOut);
			fileOut.close();

			//Showing the Done message
			JOptionPane.showMessageDialog(frame, "Done!", "Done", JOptionPane.PLAIN_MESSAGE);

			//Hidding the Frame
			frame.dispose();

			//Open the Excel file
			Desktop.getDesktop().open(new File(filePath));

		} catch (Exception e) {
			e.printStackTrace();

			//Something went wrong, let's make the user know it
			String message = "There were some error trying to open the file or Connecting to Redmine!\n\n"
					+ "See if any of the following apply:\n"
					+ "1 - Is the file already open? If yes, close it!\n"
					+ "2 - Does the file actually exist?\n"
					+ "3 - Has the file been corrupted? To check try opening it with double click and see if any error message appear\n"
					+ "4 - The file is in the correct path? It should be inside your Document folder, then inside Redmine Time Registration folder\n"
					+ "5 - Did you filled the first tab on the sheet with the right information?";

			showErrorMessage(message);

			if(fileOut != null)
				fileOut.close();
			if(frame != null)
				frame.dispose();
		}
	}

	//Method used to show the error messages
	private static void showErrorMessage(String message) {
		//Hidding the frame with the progress bar
		frame.dispose();

		JOptionPane.showMessageDialog(frame, message, "Error!", JOptionPane.ERROR_MESSAGE);
	}

	//Methos used to get the user parameters
	private static boolean getParameters(){
		String activityName;
		int id;

		//Getting the first tab, where the configuration data is
		Sheet sheet = workbook.getSheetAt(0);

		//Let's iterate over the rows
		Iterator<Row> rowIterator = sheet.rowIterator();

		//The first row is the title, let's jump it
		if(rowIterator.hasNext())
			//Getting the second row
			rowIterator.next();

		//Iterate over the rest of the lines with the type of activities 
		while (rowIterator.hasNext()) {
			if(rowIterator.hasNext()) {//what if there is no data registered, right?
				Row row = rowIterator.next();

				//Get first column, the activity name
				activityName = row.getCell(0).getStringCellValue();
				//Get first column, the activity id
				id = (int) row.getCell(1).getNumericCellValue();

				activitiesTypeEntries.put(activityName, id);
			}	
		}

		//If there is no activity type configured, let's exit the program
		if(activitiesTypeEntries.size() == 0)
			return false;

		apiAccessKey = sheet.getRow(1).getCell(3).getStringCellValue();
		//If there is no api acess key configured, let's exit the program
		if(apiAccessKey.isEmpty())
			return false;

		url = sheet.getRow(4).getCell(3).getStringCellValue();
		//If there is no url of redmine configured, let's exit the program
		if(url.isEmpty())
			return false;

		return true;
	}

	private static void saveTimeEntries(){
		Date date; 
		float hours = 0;
		int issue = 0, lines;
		String activity = null, comments, status;
		StringBuilder msgError = new StringBuilder();
		Cell cell, cellStatus;
		Iterator<Cell> cellIterator;

		//Getting the second tab, the one with the time entries
		sheet = workbook.getSheetAt(1);

		//Getting the row iterator
		Iterator<Row> rowIterator = sheet.rowIterator();

		//The first line have just the titles, let's jump'em
		if(rowIterator.hasNext())
			rowIterator.next();

		//Lets iterate over all the rows
		while (rowIterator.hasNext()) {
			msgError.setLength(0);
			lines = 0;

			if(rowIterator.hasNext()) {//If there is even one time entry
				Row row = rowIterator.next();

				// Now let's iterate over the columns of the current row
				cellIterator = row.cellIterator();

				//Getting the status cell in diferent object because I'll need it later
				cellStatus = cellIterator.next();
				status = cellStatus.getStringCellValue();

				if(status.equals("OK")) {//If status is OK I jump to next line
					setCellOK(cellStatus);
					continue;
				}

				//Getting Date
				cell = cellIterator.next();
				date = cell.getDateCellValue();
				if(date == null) {
					msgError.append("Date is empty!").append("\n");
					lines++;
				}

				//Jumping the columns until the hours calculated
				cell = cellIterator.next();
				cell = cellIterator.next();
				cell = cellIterator.next();

				if(cell.getCellTypeEnum() == CellType.STRING) {//if it's text is a total line, let's draw black and jump to next line
					setTotalCellsBlack(row);
					continue;
				}

				//Getting hours calculated
				cell = cellIterator.next();
				if(cell.getNumericCellValue() != 0)
					hours = (float) cell.getNumericCellValue();
				else {
					msgError.append("Hours Calculated is empty!").append("\n");
					lines++;
				}

				//Getting the issue number
				cell = cellIterator.next();
				if(cell.getNumericCellValue() != 0)
					issue = (int) cell.getNumericCellValue();
				else {
					msgError.append("Issue is empty!").append("\n");
					lines++;
				}

				//Getting the Activity Number
				cell = cellIterator.next();
				if(!cell.getStringCellValue().isEmpty())
					activity = cell.getStringCellValue();
				else {
					msgError.append("Activity type is empty!").append("\n");
					lines++;
				}

				//Getting the comments
				cell = cellIterator.next();
				comments = cell.getStringCellValue();

				if(msgError.length() != 0){//If I have some error message, there was some error Duh!
					setCellNOK(cellStatus, msgError.toString().trim(), lines, row);						
				}else {
					saveTimeEntry(cellStatus, issue, date, hours, activitiesTypeEntries.getOrDefault(activity, 9999), comments);
				}
			}
		}

	}

	//Method to set the cell as not ok and expand the size of the line
	private static void setCellNOK(Cell cell, String mensagem, int linhas, Row row) {
		//Creatind a new Cell Style... I don't know I have to create a new Instance every time
		cellStyleNOK = workbook.createCellStyle();
		cellStyleNOK.setFillForegroundColor(IndexedColors.RED.getIndex());//color red
		cellStyleNOK.setFillPattern(FillPatternType.SOLID_FOREGROUND);//fill the cell with the color
		cellStyleNOK.setAlignment(HorizontalAlignment.CENTER);//aling in center the text
		cellStyleNOK.setWrapText(true);//set the cell as multi line

		cell.setCellStyle(cellStyleNOK);

		row.setHeight((short) (linhas*sheet.getDefaultRowHeight()));//improving the size of the line if needed
		sheet.autoSizeColumn((short)linhas);
		cell.setCellValue(mensagem);
	}

	//Method to set the cell as not ok
	private static void setCellNOK(Cell cell, String mensagem) {
		//Creatind a new Cell Style... I don't know I have to create a new Instance every time
		cellStyleNOK = workbook.createCellStyle();
		cellStyleNOK.setFillForegroundColor(IndexedColors.RED.getIndex());//color red
		cellStyleNOK.setFillPattern(FillPatternType.SOLID_FOREGROUND);//fill the cell with the color
		cellStyleNOK.setAlignment(HorizontalAlignment.CENTER);//aling in center the text
		cellStyleNOK.setWrapText(true);//set the cell as multi line

		cell.setCellStyle(cellStyleNOK);

		cell.setCellValue(mensagem);
	}

	//Method used to save a new time entry
	private static void saveTimeEntry(Cell cell, int issue, Date date, float hours, int activities, String comments) {
		TimeEntry te = TimeEntryFactory.create();
		te.setIssueId(issue);
		te.setSpentOn(date);
		te.setHours(hours);
		te.setActivityId(activities);
		te.setComment(comments);

		try {
			timeEntryManager.createTimeEntry(te);

			setCellOK(cell);
		} catch (RedmineException e) {
			//Some error ocurred, let's make the user know it

			setCellNOK(cell, e.getMessage().trim());
		}
	}

	//Method used to set the cell as OK
	private static void setCellOK(Cell cell) {
		//Creatind a new Cell Style... I don't know I have to create a new Instance every time
		cellStyleOK = workbook.createCellStyle();
		cellStyleOK.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		cellStyleOK.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellStyleOK.setAlignment(HorizontalAlignment.CENTER);
		cellStyleOK.setWrapText(true);

		cell.setCellStyle(cellStyleOK);

		cell.setCellValue("OK");
	}

	//Method to paint and make more visible the total line
	private static void setTotalCellsBlack(Row row) {
		//Creatind a new Cell Style... I don't know I have to create a new Instance every time
		cellStyleNULL = workbook.createCellStyle();
		cellStyleNULL.setFillForegroundColor(IndexedColors.BLACK.getIndex());
		cellStyleNULL.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		//Painting balck the unused cells
		row.getCell(0).setCellStyle(cellStyleNULL);
		row.getCell(1).setCellStyle(cellStyleNULL);
		row.getCell(2).setCellStyle(cellStyleNULL);
		row.getCell(3).setCellStyle(cellStyleNULL);
		row.getCell(6).setCellStyle(cellStyleNULL);
		row.getCell(7).setCellStyle(cellStyleNULL);
		row.getCell(8).setCellStyle(cellStyleNULL);
	}

}
