import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.streaming.*;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class SpreadsheetReader {
  
    private File input;
    private XSSFWorkbook workbook;
    private SXSSFWorkbook workbookStream;
    private Sheet sheet; //Holds the current sheet
    private Row currRow; //Holds the current row
    private Cell currCell; //Holds the current cell
    private int featureCount; //number of features
    private int totalTrials; //Total number of trials per user
    private int subjectCount; //Total number of subjects
    
    public SpreadsheetReader(String filePath, int features, int trials, int users) throws IOException{
        input = new File(filePath);
        workbook = (XSSFWorkbook) WorkbookFactory.create(input);
        workbookStream = new SXSSFWorkbook(workbook, 100);
        sheet = workbook.getSheetAt(0);
        featureCount = features;
        totalTrials = trials;
        subjectCount = users;
        
        workbookStream.setCompressTempFiles(true);
    }
    
    /**Creates template data for all users and saves them to their own 
     * sheet in the xls file
     */
    public void createAllTemplates(int trials) throws IOException{
        
        createTemplateSheet();
        
        for(int userIndex=0; userIndex<=(subjectCount-1); userIndex++){
            sheet = workbook.getSheetAt(0);
            int baseRow = (userIndex*totalTrials)+1;
            createTemplate(trials, baseRow);
        }
       
    }
    
    //Checks for Template Vectors sheet and creates it if not found
    private void createTemplateSheet(){
        boolean sheetFound = false; //Boolean flag for whether the Template sheet is found
        for(int i=0; i<workbook.getNumberOfSheets(); i++){
            if(workbook.getSheetName(i).equals("Template Vectors")){
                sheetFound = true;
            }
        }
        //If Template sheet was not found, create it and label the columns
        if(sheetFound == false){
            workbook.createSheet("Template Vectors");
            sheet = workbook.getSheet("Template Vectors");
            Row headerRow = sheet.createRow(0);

            String[] columns = {"Subject", "H.period", "DD.period.t", "UD.period.t", "H.t", "DD.t.i", 
                                "UD.t.i", "H.i", "DD.i.e", "UD.i.e", "H.e", "DD.e.five", "UD.e.five", 
                                "H.five", "DD.five.Shift.r", "UD.five.Shift.r", "H.Shift.r", "DD.Shift.r.o", 
                                "UD.Shift.r.o", "H.o", "DD.o.a", "UD.o.a", "H.a", "DD.a.n", "UD.a.n", "H.n", 
                                "DD.n.l", "UD.n.l", "H.l", "DD.l.Return", "UD.l.Return", "H.Return"};
            for(int i=0; i<columns.length; i++){
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns[i]);
            }
        }
    }
    
    //Method creates template array for a single user across all features
    private void createTemplate(int trials, int baseRow)throws IOException{
        
        double total = 0; //Keeps track of total value for a single feature (for calculating average)
        double[] template = new double[featureCount]; //Array to store the template values for all features
        int rowIndex = baseRow; //Keeps track of current row
        currRow = sheet.getRow(rowIndex);
        int n = 1; //Used for iteration through the designated amount of trials
        int colIndex=0; //Keeps track of the current feature (Column)
        currCell = currRow.getCell(colIndex+3);
        
        
        //Iterates through the different features (Columns) 
        while(colIndex<featureCount){
            //Iterates down a single feature for designated amount of trials
            while (n<=trials){
               // System.out.println(currCell.getNumericCellValue()); //For testing purposes
                total += currCell.getNumericCellValue();
                rowIndex++;
                currRow = sheet.getRow(rowIndex);
                currCell = currRow.getCell(colIndex+3);
                n++;
            }
            //System.out.println(colIndex+1 + " The total is: " + total + " \nThe average is: " + total/trials); //For testing purposes
            template[colIndex] = total/trials; //Put average value into template array
            //Iterate column (next feature) and reset tracker variables
            colIndex++;
            total = 0;
            rowIndex=baseRow;
            n=1;
            currRow = sheet.getRow(rowIndex);
            currCell = currRow.getCell(colIndex+3);
        }
        currCell = currRow.getCell(0);
        String currSubject = currCell.getStringCellValue();
        writeTemplateSheet(template, currSubject);
        //Prints for testing purposes
        /**
       System.out.println("The entire template for user " + baseRow + " is: ");
        for(double i: template){
            System.out.print(i +", ");
        }
          System.out.print("\n\n\n"); 
          **/
    }
    
    //Writes template values to spreadsheet for a single user
    private void writeTemplateSheet(double[] template, String subject){
        String currSubject = subject;
        sheet = workbook.getSheet("Template Vectors");
        Row writeRow = null;
        //Checks if this is the first time filling out sheet so it can create rows & cells as needed
        if(sheet.getPhysicalNumberOfRows() < (subjectCount+1)){
            writeRow = sheet.createRow(sheet.getLastRowNum()+1);
            Cell cell = writeRow.createCell(0);
            cell.setCellValue(currSubject);
            for(int i=0; i<template.length; i++){
                cell = writeRow.createCell(i+1);
                cell.setCellValue(template[i]);
            }
        }
        //All rows and cells have already been created. Just updates values.
        else{
            for(int i=1; i<sheet.getPhysicalNumberOfRows(); i++){
                Row testRow = sheet.getRow(i);
                Cell cell = testRow.getCell(0);
                if(cell.getStringCellValue().equals(subject)){
                    writeRow = testRow;
                    break;
                }  
            }
            if(writeRow != null){
                for(int j=0; j<template.length; j++){
                    Cell cell = writeRow.getCell(j+1);
                    cell.setCellValue(template[j]);
                }
            }
        }       
    }
    
    //Fetches a single template vector for the designated subject
    private double[] fetchTemplate(String subject){
        double[] template = new double[featureCount];
        sheet = workbook.getSheet("Template Vectors");
        int rowIndex = 1;
        currRow = sheet.getRow(rowIndex);
        for(int i=1; i<sheet.getPhysicalNumberOfRows();i++){
            Cell cell = currRow.getCell(0);
            if(subject.equals(cell.getStringCellValue()))
                break;
            else{
                rowIndex++;
                currRow = sheet.getRow(rowIndex);
            }
        }
        for(int cellIndex=1; cellIndex<=featureCount;cellIndex++){
            Cell cell = currRow.getCell(cellIndex);
            template[cellIndex-1] = cell.getNumericCellValue();
        }
        return template;
    }
    
    //Generates a single score for one probe vector against the designated subject's template vector
    private double generateSingleScore(int rowIndex, double[] template){   
        sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(rowIndex);
        int cellIndex = 0;
        double[] probe = new double[featureCount];
        //Iterates through features
        while(cellIndex < featureCount){
            currCell = row.getCell(cellIndex+3);
            probe[cellIndex] = currCell.getNumericCellValue();
            cellIndex++;
        } 
        double score = 0;
        for(int i=0; i<featureCount; i++){
           //System.out.println("Subtraction " + i + ": " + Math.abs((template[i]-probe[i])));
            score += Math.abs((template[i]-probe[i]));
            //System.out.println("Current total: " + score +"\n");
        }
        score /= featureCount;
        return score;
    }
    
    /**
    This method should be called after templates are generated. It will iterate
    * through all template vectors and generate scores for all test vectors against
    * each template. All scores are written to the Score sheet.
     */
    public void generateScores(int templateTrials)throws IOException{
        
        createScoreSheet();
        
        String templateSubject;
        double[] currTemplate;
        
        //Iterates through the template vectors
        for(int templateIndex = 1; templateIndex<=subjectCount; templateIndex++){
            sheet = workbook.getSheet("Template Vectors");
            currRow = sheet.getRow(templateIndex);
            currCell = currRow.getCell(0);
            templateSubject = currCell.getStringCellValue();
            //System.out.println("Current Template Subject: " + templateSubject);
            currTemplate = fetchTemplate(templateSubject);
            sheet = workbook.getSheetAt(0);
            
            //Iterates through subjects for test vectors
            for(int testIndex=0; testIndex<subjectCount; testIndex++){
                sheet = workbook.getSheetAt(0);
                int rowIndex = ((testIndex*totalTrials)+templateTrials)+1;
                //System.out.println("Row Index: " + rowIndex);
                currRow = sheet.getRow(rowIndex);
                currCell = currRow.getCell(0);
                String testSubject = currCell.getStringCellValue();
                //System.out.println("Current Test Subject: " + testSubject);
                boolean genuine = templateSubject.equals(testSubject);
                //System.out.println("Genuine? -- " + genuine);
                int stopIndex = (rowIndex+(totalTrials-templateTrials));
                
                //Iterates through rows(test vectors) for a single user
                while(rowIndex< stopIndex){
                    double score = generateSingleScore(rowIndex, currTemplate);
                    writeToScoreSheet(templateSubject, score, genuine);
                    rowIndex++;
                }
            }
        }
    }
    
    /**Creates a fresh Score sheet. Will delete an old Score sheet if it exists, 
     * and then replace it with a blank sheet. */
    private void createScoreSheet(){
        boolean sheetFound = false; //Boolean flag for whether the Score sheet is found
        for(int i=0; i<workbook.getNumberOfSheets(); i++){
            if(workbook.getSheetName(i).equals("Scores")){
                sheetFound = true;
            }
        }
        //If Score sheet was found, delete it and recreate the sheet
        if(sheetFound == true){
            workbook.removeSheetAt(workbook.getSheetIndex("Scores"));
        }
            workbook.createSheet("Scores");
            sheet = workbook.getSheet("Scores");
            Row headerRow = sheet.createRow(0);

            String[] columns = {"Subject", "Type", "Score"};
            for(int i=0; i<columns.length; i++){
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns[i]);
            }
    }

    //Writes a single score to the Score sheet
    private void writeToScoreSheet(String subject, double score, boolean genuine){
        String currSubject = subject;
        sheet = workbook.getSheet("Scores");
        Row writeRow = sheet.createRow(sheet.getLastRowNum()+1);
        Cell cell = writeRow.createCell(0);
        cell.setCellValue(currSubject);
        cell = writeRow.createCell(1);
        if(genuine)
            cell.setCellValue("Genuine");
        else
            cell.setCellValue("Imposter");
        cell = writeRow.createCell(2);
        cell.setCellValue(score);
        
    }
    
    /**Creates a fresh Rates sheet. Will delete an old Rates sheet if it exists, 
     * and then replace it with a blank sheet. (Currently not used in favor of
     * CSV format) */
    private void createRatesSheet(){
        boolean sheetFound = false; //Boolean flag for whether the Rates sheet is found
        for(int i=0; i<workbook.getNumberOfSheets(); i++){
            if(workbook.getSheetName(i).equals("Rates")){
                sheetFound = true;
            }
        }
        //If Rates sheet was found, delete it and recreate the sheet
        if(sheetFound == true){
            workbook.removeSheetAt(workbook.getSheetIndex("Rates"));
        }
            workbook.createSheet("Rates");
            sheet = workbook.getSheet("Rates");
            Row headerRow = sheet.createRow(0);

            String[] columns = {"Threshold", "FRR", "FPR"};
            for(int i=0; i<columns.length; i++){
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns[i]);
            }
    }
    
    /*Generates all FPR and FRR rates at every score as a threshold then 
     * writes to a csv file. */ 
    public void generateRates(int templateTrials) throws FileNotFoundException{
    	sheet = workbook.getSheet("Scores");
    	Iterator<Row> rowIterator = sheet.rowIterator();
    	double falsePass = 0;
    	double falseReject =  (totalTrials - templateTrials) * subjectCount;
    	double totalGenuine = (totalTrials - templateTrials) * subjectCount;
    	double totalImposters = (totalTrials - templateTrials) * subjectCount * (subjectCount - 1);
    	StringBuilder sb = new StringBuilder();
    	PrintWriter writer = new PrintWriter(new File("Rates.csv"));
    	sb.append("Threshold,FRR,FPR\n0,1.0,0.0");
    	Row row = rowIterator.next();
    	while(rowIterator.hasNext()){
    		row = rowIterator.next();
    		currCell = row.getCell(1);
    		String scoreType = currCell.getStringCellValue();
    		currCell = row.getCell(2);
    		double score = currCell.getNumericCellValue();
    		if(scoreType.equals("Genuine"))
    			falseReject--;
    		else
    			falsePass++;
    		double frr = falseReject / totalGenuine;
    		double fpr = falsePass / totalImposters;
    		sb.append("\n"+score+","+frr+","+fpr);
    	}
    	writer.write(sb.toString());
    	writer.close();
    	
    }
    
    //Writes a single rate to the Rates sheet (Currently not used in favor of CSV format)
    private void writeToRatesSheet(double threshold, double frr, double fpr){
        sheet = workbook.getSheet("Rates");
        Row writeRow = sheet.createRow(sheet.getLastRowNum()+1);
        Cell cell = writeRow.createCell(0);
        cell.setCellValue(threshold);
        cell = writeRow.createCell(1);
        cell.setCellValue(frr);
        cell = writeRow.createCell(2);
        cell.setCellValue(fpr);
        
    }
    
    
    //Used to sort Scores in ascending order (Not currently in use due to method causing issues)
    public void sortColumns(){
    	sheet = workbook.getSheet("Scores");
    	Map<Double, Row> sortedRows = new TreeMap<>();
    	Iterator<Row> rowIterator = sheet.rowIterator();
    	currRow = rowIterator.next();
    	while(rowIterator.hasNext()){
    		currRow = rowIterator.next();
    		sortedRows.put(currRow.getCell(2).getNumericCellValue(), currRow);
    	}
    	createScoreSheet();
    	sheet = workbook.getSheet("Scores");
    	int rowIndex = 1;
    	for(Row row : sortedRows.values()){
    		Row newRow = sheet.createRow(rowIndex);
    		copyRow(row, newRow);
    		rowIndex++;
    	}
    }
    
    private void copyRow(Row row, Row newRow){
    	Iterator<Cell> cellIterator = row.cellIterator();
    	int cellIndex = 0;
    	while(cellIterator.hasNext()){
    		Cell cell = cellIterator.next();
    		Cell newCell = newRow.createCell(cellIndex);
    		if(cellIndex < 2)
    			newCell.setCellValue(cell.getStringCellValue());
    		else
    			newCell.setCellValue(cell.getNumericCellValue());
    		cellIndex++;
    	}
    }
    
    public String ratesAt(double threshold){
    	sheet = workbook.getSheet("Rates");
    	Iterator<Row> rowIterator = sheet.rowIterator();
    	currRow = rowIterator.next();
    	while(rowIterator.hasNext()){
    		currRow = rowIterator.next();
    		currCell = currRow.getCell(0);
    		double currThreshold = currCell.getNumericCellValue();
    		if(currThreshold > threshold){
    			currCell = currRow.getCell(1);
    			double frr = currCell.getNumericCellValue();
    			currCell = currRow.getCell(2);
    			double fpr = currCell.getNumericCellValue();
    			return ("At threhold t= " + threshold + "\nFRR= " + frr + "\nFPR= " + fpr + "\n");
    		}
    	}
    	return ("At threhold t= " + threshold + "\nFRR= 1" + "\nFPR= 0" + "\n");
    }
    
    //Saves all scores to a csv file
    public void saveScores() throws FileNotFoundException{
    	sheet = workbook.getSheet("Scores");
    	Iterator<Row> rowIterator = sheet.rowIterator();
    	StringBuilder sb = new StringBuilder();
    	currRow = rowIterator.next();
    	sb.append("Subject,Type,Score\n");
    	while(rowIterator.hasNext()){
    		currRow = rowIterator.next();
    		currCell = currRow.getCell(0);
    		sb.append(currCell.getStringCellValue()+",");
    		currCell = currRow.getCell(1);
    		sb.append(currCell.getStringCellValue()+",");
    		currCell = currRow.getCell(2);
    		sb.append(currCell.getNumericCellValue()+"\n");
    	}
    	PrintWriter writer = new PrintWriter(new File("Scores.csv"));
    	writer.write(sb.toString());
    	writer.close();
    }
    
    //Saves all changes that were made on the Workbook to the xlsx file (Memory intensive, will likely cause out of memory error)
    public void saveChanges()throws IOException{
    	   	    	
        FileOutputStream fileOut = new FileOutputStream("Test.xlsx");
        workbookStream.write(fileOut);
        fileOut.close();
        workbook.close();
        workbookStream.dispose();
        workbookStream.close();
    }
}


