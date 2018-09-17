
package nlc2;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ibm.watson.developer_cloud.natural_language_classifier.v1.NaturalLanguageClassifier;
import com.ibm.watson.developer_cloud.natural_language_classifier.v1.model.Classification;
import com.ibm.watson.developer_cloud.natural_language_classifier.v1.model.Classifier;
import com.ibm.watson.developer_cloud.natural_language_classifier.v1.model.Classifiers;

public class nlcbro 
{

	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException 
	{
		
		File myFile = new File("nlcPredict1.xlsx");
		FileInputStream fis = null;
		
		int countHITS = 0;
		
		fis = new FileInputStream(myFile);
	
		@SuppressWarnings("resource")
		XSSFWorkbook myWorkBook = null;
		
		myWorkBook = new XSSFWorkbook (fis);
		
		int numSheet = myWorkBook.getNumberOfSheets();
		
		for (int x = 0; x < numSheet; x++)
		{
		
			XSSFSheet mySheet = myWorkBook.getSheetAt(x);
			Iterator<Row> rowIterator = mySheet.iterator();
		
		
			//Making new heading
			Row r = mySheet.getRow(0);
			int targetColumnIndex = r.getLastCellNum();
		
			Cell cellHead = CellUtil.getCell(CellUtil.getRow(0, mySheet), targetColumnIndex);
			cellHead.setCellValue("Class");
			////
		
		
			//Getting the column number for comments col
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			int sourceColumnIndex = -1;
          
			while (cellIterator.hasNext()) {
	
				Cell cell1 = cellIterator.next();
	      	
				String comment = cell1.toString();
	      	
				if (comment.equals("Text"))
				{
					sourceColumnIndex = cell1.getColumnIndex();
				}	      	
			}
	    
	    
			if (sourceColumnIndex != -1)
			{
				int totalRows = mySheet.getLastRowNum();
				
				NaturalLanguageClassifier service = new NaturalLanguageClassifier();
				service.setUsernameAndPassword("61dcbf0c-b649-47d3-a4eb-2723e6e087bd", "FU7LAaFTQTjX");
								
		    
		    	for (int i = 1; i <= totalRows ; ++i)
			    {
		    		Cell myCell = CellUtil.getCell(CellUtil.getRow(i, mySheet), sourceColumnIndex);
			    	
			    	String text = myCell.toString();
			    			    	
			    	
			    	if(text != null && !text.isEmpty())
			    	{
			    		Classification classification = service.classify("f7ec36x309-nlc-938", text).execute();
			    		
			    		
			    		String myClass = classification.getTopClass();
			    		
		    		
			    		countHITS++;
						
						System.out.println(countHITS);
						
						Cell myCell2 = CellUtil.getCell(CellUtil.getRow(i, mySheet), targetColumnIndex);
										    
					    myCell2.setCellValue(myClass);
			    	
			    		
			    	}
			    }
				
			}
		}
		
		myWorkBook.write(new FileOutputStream("NLCresults.xlsx"));
		myWorkBook.close();
		
		System.out.println("End");
		
	}

}
