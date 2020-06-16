import java.io.File;  
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFSheet;  
import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;
import java.util.Map.Entry;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

public class Main {
	static ArrayList<HashMap<String, Integer>> allMaps = new ArrayList<>();
	static final int NMONTHS = 6;
	static final int NQUAF = 5;
	
	//public static final String SAMPLE_XLSX_FILE_PATH = System.getProperty("user.dir")+ System.getProperty("file.separator") + "Aug 2019.xls";
	
	public static void main(String[] args) throws IOException, InvalidFormatException {
		// TODO Auto-generated method stub
		String[] months = {"Jan", "Feb", "Mar", "Apr", "May" , "Jun", "Jul", 
				"Aug", "Sep", "Oct", "Nov", "Dec"};
		
		HashMap<String,Integer> hm = new HashMap<String, Integer>();
		for (int i = 0; i < NMONTHS; i++) {
			hm = parseWorkbook(months[i], "2019", hm);
		}
		//printMap(hm);
		allMaps.add(hm);
		printFirstReport(months[NMONTHS % 12], 
				Integer.toString(2019 + (NMONTHS / 12)), allMaps.get(0));
		int counter = 1;
		while (true) {
			HashMap<String,Integer> temphm = new HashMap<String, Integer>();
			try {
				for (int x = counter; x < counter + NMONTHS; x++) {
					String year = Integer.toString(2019 + (x / 12));
					temphm = parseWorkbook(months[x % 12], year, temphm);
				}
				allMaps.add(temphm);
				printNextReport(months[(counter + NMONTHS) % 12], 
						Integer.toString(2019 + ((counter + NMONTHS) / 12)), counter);
				counter++;
				//System.out.println(counter);
			} catch (FileNotFoundException e) {
				break;
			}
		}
		
	}
	
	public static HashMap<String,Integer> parseWorkbook(String month, 
			String year, HashMap<String,Integer> hm) 
					throws InvalidFormatException, IOException {
		
		String filename = System.getProperty("user.dir") + 
				System.getProperty("file.separator") + month + " " + year 
				+ ".xls";
		
		Workbook workbook = WorkbookFactory.create(new File(filename));
        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();
        
        // 2. Or you can use a for-each loop to iterate over the rows and columns
        //System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
        int row_counter = 0;
        int cell_counter = 0;
        String cellValueString = "";
    	double cellValueDouble = 0.0;
        for (Row row: sheet) {
            for(Cell cell: row) {
            	if (row_counter <= 4 || cell_counter >=2) {
            		cell_counter = 0;
            		break;
            	}
            	
            	switch (cell.getCellTypeEnum()) {
            	case STRING: 
            		cellValueString = dataFormatter.formatCellValue(cell);
                	if (cellValueString.equals("Totals"))
                		return hm;
                	//System.out.print(cellValueString + "\t");
                	cell_counter++;
                	break;
            	case NUMERIC:
                	cellValueDouble = cell.getNumericCellValue();
                	//System.out.print(cellValueDouble + "\t");
                	cell_counter++;
                	break;
				default:
					break;
            	
            	}
            }
            hm = addToMap(cellValueString, cellValueDouble, hm);
            row_counter++;
            //System.out.println();
        }
        workbook.close();
        return hm;
	}
	
	public static HashMap<String,Integer> addToMap(String name, double hours, 
			HashMap<String, Integer> hm) {
		if (!hm.containsKey(name)) {
			Integer over120 = hours >= 120 ? 1 : 0;
			hm.put(name, over120);
			return hm;
		}
		Integer current = hm.get(name);
		current = hours >= 120 ? current + 1 : current;
		hm.put(name, current);
		return hm;
	}
	
	public static void printMap(HashMap<String, Integer> hm) {
		hm.forEach((k,v) -> System.out.println("key: "+k+" value:"+v));
	}
	
	public static void printFirstReport(String month, String year, 
			HashMap<String,Integer> hm) 
			throws IOException, InvalidFormatException{
		// Create a Workbook
        Workbook workbook = new HSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

        /* CreationHelper helps us create instances of various things like DataFormat, 
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet(month + " " + year + " Report");

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);
        
        String[] columns = {"Qualified", "New Additions", "Dropped"};

        // Create cells
        for(int i = 0; i < 3; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Create Other rows and cells with employees data
        int rowNum = 1;
        Set<Entry<String, Integer>> entries = hm.entrySet();
        for(Map.Entry<String, Integer> entry : entries) {
        	if (entry.getValue() < NQUAF) //MIGHT HAVE TO CHANGE TO NMONTHS
        		continue;
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(entry.getKey());
            row.createCell(1).setCellValue(entry.getKey());

        }

		// Resize all columns to fit the content size
        for(int i = 0; i < 3; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        String filename = "Reports" + System.getProperty("file.separator");
        
        FileOutputStream fileOut = new FileOutputStream(filename + month + " " + year + " Report.xls");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
	}
	
	public static void printNextReport(String month, String year, int counter) 
			throws IOException, InvalidFormatException{
        Workbook workbook = new HSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

        /* CreationHelper helps us create instances of various things like DataFormat, 
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet(month + " " + year + " Report");

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);
        
        String[] columns = {"Qualified", "New Additions", "Dropped"};

        // Create cells
        for(int i = 0; i < 3; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Create Other rows and cells with employees data
        int rowNum = 1;
        Set<Entry<String, Integer>> entries = allMaps.get(counter).entrySet();
        HashMap<String, Integer> prevMap = allMaps.get(counter - 1);
        Set<String> currentSet = new HashSet<>();
        Set<String> prevSet = new HashSet<>();
        for (Map.Entry<String, Integer> entry : prevMap.entrySet()) {
        	if (entry.getValue() < NQUAF) { // MIGHT HAVE TO CHANGE TO NMONTHS
        		continue;
        	}
        	prevSet.add(entry.getKey());
        }
        
        for(Map.Entry<String, Integer> entry : entries) {
        	if (entry.getValue() < NQUAF) // MIGHT HAVE TO CHANGE TO NMONTHS
        		continue;
            Row row = sheet.createRow(rowNum++);
            String newEntry = entry.getKey();

            row.createCell(0).setCellValue(newEntry);
            currentSet.add(newEntry);
            if (!prevSet.contains(newEntry)) {
            	row.createCell(1).setCellValue(newEntry);
            }
        }
        
        Iterator<String> it = prevSet.iterator();
        rowNum =1;
        while(it.hasNext()){
            String temp = it.next();
            if (!currentSet.contains(temp)) {
            	Row row = sheet.getRow(rowNum++);
            	if (row == null)
            		row = sheet.createRow(rowNum - 1);
            	row.createCell(2).setCellValue(temp);
            }
            	
        }
        

		// Resize all columns to fit the content size
        for(int i = 0; i < 3; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file 
        String outfile = "Reports" + System.getProperty("file.separator") + 
        	month + " " + year + " Report.xls";
        
        FileOutputStream fileOut = new FileOutputStream(outfile);
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
	}
	
}

