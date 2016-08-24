package JavaTesting;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelComparison {


	static int KEY_COLUMNS[] = {1,2}; 
	static int CONSIDER_SKIP_ROW_COUNT = 5;
	

static boolean compareKeyColumns(Row a, Row b, int columnCount, int keyColumns[]) {
	
	//if (a != null) System.out.println("Row A is not null!");
	//if (b != null) System.out.println("Row B is not null!");
	//System.out.println("Keys to match is: "+keyColumns.length);
	//boolean matchValue = true;
	System.out.println("====> In compare Method.");
	
	for (int i=0; i<keyColumns.length; i++){
		
		String aCelVal = a.getCell(keyColumns[i]-1).toString().trim();
		String bCelVal = b.getCell(keyColumns[i]-1).toString().trim();
		
		if(aCelVal.equals(bCelVal)) System.out.println("|__> A cell value: "+aCelVal+", B cell value: "+bCelVal+" --> values matched! "); 
		else	return false;
		
	}
	
	System.out.println("Row Matched!");
	
	return true;
}

	public static void main(String[] args) throws Exception{
		// TODO Auto-generated method stub
		
		String fileName = "C:\\Users\\bk\\Desktop\\test.xlsx";
		if (!doBasicValidationOfSheets(fileName) ) {
			System.out.println("Basic Validations failed, exiting from program!");
			System.exit(1);
		}
// Do Basic Validation

		
		Workbook wb = null;
		Sheet firstSheet = null;
		Sheet secondSheet = null;
		
		
			wb = new XSSFWorkbook(new FileInputStream(fileName));
			firstSheet = wb.getSheetAt(0);
			secondSheet = wb.getSheetAt(1);
		
           int firstSheetRowCount = firstSheet.getPhysicalNumberOfRows();
           int secondSheetRowCount = secondSheet.getPhysicalNumberOfRows();
           int iterationCount;
           int firstSheetColumnCount = firstSheet.getRow(0).getPhysicalNumberOfCells();
           int secondSheetColumnCount = secondSheet.getRow(0).getPhysicalNumberOfCells();
           
           if(firstSheetRowCount >= secondSheetRowCount)	iterationCount = firstSheetRowCount;
           else iterationCount = secondSheetRowCount;
           
           System.out.println("now of rows in sheet1  = " + firstSheetRowCount);
           System.out.println("now of rows in sheet2  = " + secondSheetRowCount);
           System.out.println("now of rows for iteration  = " + iterationCount);
           System.out.println("now of cloumns in sheet1 = " + firstSheetColumnCount);
           System.out.println("now of cloumns in sheet2 = " + secondSheetColumnCount);
          
           
           // start Data Validation
           // start row iteration
           
        // Compare Cells
   		// compare key cells
   			// if key cells are not matching, check for other row till it matches / till defined count and insert one row
   			// else continue comparison with next row
   		// compare other optional cells
   			// if not matching highlight
   			// else continue
           ArrayList missingRows = new ArrayList();
           ArrayList totalMissingRows = new ArrayList();
            for (int rowCurrentPosition=0; rowCurrentPosition<iterationCount; rowCurrentPosition++){
            	int misMatchCount = 0;
            		while (misMatchCount <= CONSIDER_SKIP_ROW_COUNT){
            			//System.out.println("Sheet1 "+firstSheet.getRow(i).toString());
            			//System.out.println("Sheet2 "+secondSheet.getRow(i).toString());
            			System.out.println("Mismatch count: "+misMatchCount);
            			if (!compareKeyColumns(firstSheet.getRow(rowCurrentPosition),secondSheet.getRow(rowCurrentPosition+misMatchCount), firstSheetColumnCount, KEY_COLUMNS)) {
            				missingRows.add(secondSheet.getRow(rowCurrentPosition+misMatchCount));
            				
            				System.out.println("In if condition >>>>>>>>>>");
            				misMatchCount += 1;
            			}else{
            				
            				if (misMatchCount > 0) {
            					totalMissingRows.addAll(missingRows);
                				insertMissingRowsToSheet(missingRows, rowCurrentPosition, fileName);
                				missingRows.removeAll(missingRows);
            				}
            				misMatchCount = 0;
            				System.out.println("In else condition");
            				break;
            				// Add missing Row to Sheet 3 and insert missing row to current sheet
            				
            			}
            		}
            }
            
            
            // Write the output to a file
            FileOutputStream fileOut = new FileOutputStream("C:\\Users\\bk\\Desktop\\Tidal Blast.xlsx");
            wb.write(fileOut);
            fileOut.close();

	}
	
	/**
	 * 
	 * @param missingRow
	 * @param startingPosition, from which position we need to insert the records (Rows start from index 0);
	 * @return true if record insertion is successful
	 */
	public static boolean insertMissingRowsToSheet(ArrayList missingRow, int startingPosition, String fileName){
		
		boolean validationStatus = false;
		Workbook wb = null;
		Sheet firstSheet = null;
		Sheet secondSheet = null;
		
		try {
			wb = new XSSFWorkbook(new FileInputStream(fileName));
			firstSheet = wb.getSheetAt(0);
			secondSheet = wb.getSheetAt(1);
			
			int firstSheetRowCount = firstSheet.getPhysicalNumberOfRows();
		    int secondSheetRowCount = secondSheet.getPhysicalNumberOfRows();
		    
		    if (missingRow != null) {
		    	firstSheet.shiftRows(startingPosition, firstSheet.getLastRowNum(), missingRow.size());
		    	for(Object row: missingRow){
		    		firstSheet.createRow(startingPosition).createCell(0).setCellValue("Cell Value inserted!");
		    		
		    		System.out.println("Will Insert a new row here!");
		    		//firstSheet.createRow(arg0)
		    	
		    		
		    	}
		    }
		    FileOutputStream outFile = new FileOutputStream(new File("C:\\Users\\bk\\Desktop\\test.xlsx"));
   		 wb.write(outFile);
   	 	                    outFile.close();
   	 	                    wb.close();
		    
		}catch(Exception e){
			System.out.println("Exception while inserting missing records: ");
			e.printStackTrace();
		}
		return true;
	}
	
	static boolean doBasicValidationOfSheets(String fileName){
		
		boolean validationStatus = false;
		Workbook wb = null;
		Sheet firstSheet = null;
		Sheet secondSheet = null;
		
		try {
			wb = new XSSFWorkbook(new FileInputStream(fileName));
			firstSheet = wb.getSheetAt(0);
			secondSheet = wb.getSheetAt(1);
			
			int firstSheetRowCount = firstSheet.getPhysicalNumberOfRows();
		    int secondSheetRowCount = secondSheet.getPhysicalNumberOfRows();
		    if (firstSheetRowCount <= 1 || secondSheetRowCount <= 1){
		    	System.out.println("Row data is missing from sheet! ");
		    	return false;
		    }
		    
		    
		    int firstSheetColumnCount = firstSheet.getRow(0).getPhysicalNumberOfCells();
		    int secondSheetColumnCount = secondSheet.getRow(0).getPhysicalNumberOfCells();
		    
		    if (firstSheetColumnCount != secondSheetColumnCount ){
		    	System.out.println("Columns are not matching! Sheet1 Column Count = "+firstSheetColumnCount + ", Sheet2 Column Count = "+secondSheetColumnCount);
		    	return false;
		    }
		    
		} catch (FileNotFoundException e) {
			System.out.println("File ("+fileName+") is not present!");
			e.printStackTrace();
			return validationStatus;
		}catch (IOException e){
			System.out.println("Got error while reading file");
			return validationStatus;
		}catch(IllegalArgumentException iE){
			System.out.println("Sheet does not exist!");
			return validationStatus;
		}finally{
			try{
				wb.close();
			}catch(Exception e){}
		}
			System.out.println("Basic validation completed!");
			return true;
		}

	

}