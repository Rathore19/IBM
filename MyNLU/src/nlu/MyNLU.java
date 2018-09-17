
package nlu;


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

import com.ibm.watson.developer_cloud.natural_language_understanding.v1.NaturalLanguageUnderstanding;
import com.ibm.watson.developer_cloud.natural_language_understanding.v1.model.AnalysisResults;
import com.ibm.watson.developer_cloud.natural_language_understanding.v1.model.AnalyzeOptions;
import com.ibm.watson.developer_cloud.natural_language_understanding.v1.model.EntitiesOptions;
import com.ibm.watson.developer_cloud.natural_language_understanding.v1.model.Features;
import com.ibm.watson.developer_cloud.natural_language_understanding.v1.model.KeywordsOptions;
import com.ibm.watson.developer_cloud.natural_language_understanding.v1.model.SentimentOptions;

public class MyNLU {

	public static void main(String[] args) throws FileNotFoundException, IOException 
	{
		File myFile = new File("Filename.xlsx");
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
			cellHead.setCellValue("Sentiments");
			////
		
		
			//Getting the column number for comments col
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			int sourceColumnIndex = -1;
          
			while (cellIterator.hasNext()) {
	
				Cell cell1 = cellIterator.next();
	      	
				String comment = cell1.toString();
	      	
				if (comment.equals("Comments"))
				{
					sourceColumnIndex = cell1.getColumnIndex();
				}	      	
			}
	    
	    
			if (sourceColumnIndex != -1)
			{

	    	
		    	int totalRows = mySheet.getLastRowNum();
		    	
//		    	NaturalLanguageUnderstanding service = new NaturalLanguageUnderstanding(
//						  NaturalLanguageUnderstanding.VERSION_DATE_2017_02_27,
//						  "3764a2f7-c641-44dd-8d73-7061f945a233",
//						  "EW2btCaF2v7N"
//						);
		    	
		    	NaturalLanguageUnderstanding service = new NaturalLanguageUnderstanding(
						  NaturalLanguageUnderstanding.VERSION_DATE_2017_02_27,
						  "f98b90f1-7caf-4ca0-929d-e186b4f62b58",
						  "n8SrvmGTc7v7"
						);
		    	
		    	List<String> targets = new ArrayList<>();
				targets.add("COMMENTS");
				
				
		    
		    	for (int i = 1; i <= totalRows ; ++i)
			    {
		    		
		    			    
		    		Cell myCell = CellUtil.getCell(CellUtil.getRow(i, mySheet), sourceColumnIndex);
			    	
			    	String text = myCell.toString();
			    			    	
			    	
			    	if(text != null && !text.isEmpty())
			    	{    
			
			    		text = "COMMENTS: " + text;
			    				    		
						SentimentOptions sentiment = new SentimentOptions.Builder()
						  .targets(targets)
						  .document(true)
						  .build();
						
						
						KeywordsOptions keywords= new KeywordsOptions.Builder()
						 .sentiment(true)
						 .build();
						
			
						Features features = new Features.Builder()
						  .sentiment(sentiment)
						  .keywords(keywords)
						  .build();
							
			
						AnalyzeOptions parameters = new AnalyzeOptions.Builder()
						  .text(text)
						  .features(features)
						  .language("en")
						  .build();
							
						try {
						AnalysisResults response = service
						  .analyze(parameters)
						  .execute();
						
						String mySentiment = response.getSentiment().getDocument().getLabel();
						
						countHITS++;
						
						System.out.println(countHITS);
						//System.out.println(text);
						
						Cell myCell2 = CellUtil.getCell(CellUtil.getRow(i, mySheet), targetColumnIndex);
										    
					    myCell2.setCellValue(mySentiment);
					    
					    //System.out.println(response);
					    
						}
						catch (Exception ex){
							String mySentiment = "Insufficent Text";
									
							Cell myCell2 = CellUtil.getCell(CellUtil.getRow(i, mySheet), targetColumnIndex);
						    
						    myCell2.setCellValue(mySentiment);		
							
						}
						
						
			    	}
						
			    }
			}
			
	    }
		
		myWorkBook.write(new FileOutputStream("results.xlsx"));
		myWorkBook.close();
	    
	    
	    System.out.println("End");

	}

}
