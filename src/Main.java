import java.io.File;
import java.io.IOException;
import java.util.Date; 
import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Number;
import jxl.write.biff.RowsExceededException; 
 
public class Main {

	static Cell[] c;

	public static void main(String args[]) throws BiffException, IOException, RowsExceededException, WriteException{
		
		readSheet("C:\\Users\\user\\Documents\\workspace\\SpredsheetApp\\dataorg.xls");
		
		makeWorkbook("C:\\Users\\user\\Documents\\workspace\\SpredsheetApp\\test.xls");
		
	}
	
	
	//method to open an xls file
	public static void readSheet(String path){
		
		Workbook workbook;
		try {
			
			workbook = Workbook.getWorkbook(new File(path));
			Sheet sheet = workbook.getSheet(0);
			
			final int NUM_OF_ROWS = sheet.getRows();
			final int NUM_OF_CELLS = sheet.getColumns() * NUM_OF_ROWS;
			
			
			/* reads all the cells and store them */
			
			c = new Cell[NUM_OF_CELLS];
			int i=0;
			
			for(int y=0; y<NUM_OF_ROWS; y++){
				
				for(int x=0; x<6; x++){ //the condition is not dynamic 
					c[i] = sheet.getCell(x, y);
					i++;
				}
				
				
			}
			
			/** end reading cells */
			
			
			
			
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}
	
	public static void makeWorkbook(String path) throws IOException, RowsExceededException, WriteException{
		
		WritableWorkbook workbook = Workbook.createWorkbook(new File(path));
		WritableSheet sheet = workbook.createSheet(getEmployeeName(), 0);
		
		sheet.mergeCells(0, 0, 2, 0); //if you want to add to the merged cell, use the starting cell (first two args)
		
		Label label2 = new Label(0, 0, "adding to merged"); //used the starting cell
		sheet.addCell(label2); 
		
		Label label3 = new Label(3, 0, "ßÔÝ ÇáÍÖæÑ æÇáÅäÕÑÇÝ ááÈÇÈ ÇáÃæá / 2016");
		sheet.addCell(label3);
		
		Label label = new Label(0, 2, "A label record"); 
		sheet.addCell(label); 

		Number number = new Number(3, 4, 3.1459); 
		sheet.addCell(number); 
		 
		workbook.write(); 
		workbook.close();
		
	}
	
	public String getEmployeeName(int id){
		
		  switch (id) {
          case 1001: return "Mohsen";
          case 1002: return "emp2";
          default: return "Unnamed";
                   
      }
		
	}



}
