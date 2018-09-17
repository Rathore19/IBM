package MyEval;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;



public class MyEval {
	
	public static void main(String[] args) throws IOException
	{
		File myFile = new File("Lenovo NLP Output.xlsx");
		FileInputStream fis = null;
		
		fis = new FileInputStream(myFile);
		
		@SuppressWarnings("resource")
		XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
		
		int numSheet = myWorkBook.getNumberOfSheets();
		
		System.out.println(numSheet);
		
		XSSFSheet mySheet = myWorkBook.getSheetAt(0);
		Iterator<Row> rowIterator = mySheet.iterator();
		
			
		
		
		Row row = rowIterator.next();
		Iterator<Cell> cellIterator = row.cellIterator();
		int SentiColumnIndex = -1;
		int MySentiColumnIndex = -1;
		

		
	    while (cellIterator.hasNext()) {
	
	    	Cell cell1 = cellIterator.next();
	      	
	      	String comment = cell1.toString();
	      	
	      	if (comment.equals("Sentiments"))
	      	{
	      		SentiColumnIndex = cell1.getColumnIndex();
	      	}
	      	if (comment.equals("MySentiments"))
	      	{
	      		MySentiColumnIndex = cell1.getColumnIndex();
	      	}
	    }
	    
	    
	    if (SentiColumnIndex != -1 && MySentiColumnIndex != -1)
	    {
	    		    
		    int totalRows = mySheet.getLastRowNum();
		    int totalScore = 0;
		    int countTotal = 0;
		    int countAcc = 0;
		    
		    int countOfNegPosErrors = 0;
		    
		    for (int i = 1; i <= totalRows ; ++i)
		    {
		    	int tempScore1 = 0;
		    	int tempScore2 = 0;
		    	
		    	int Score = 0;
		    	
		    	
		    	Cell myCell1 = CellUtil.getCell(CellUtil.getRow(i, mySheet), SentiColumnIndex);
		    	Cell myCell2 = CellUtil.getCell(CellUtil.getRow(i, mySheet), MySentiColumnIndex);
		    	
		    	String GeneratedSenti = myCell1.toString();
		    	String MySenti = myCell2.toString();
		    	
		    	if (GeneratedSenti.equals("positive"))
		    	{GeneratedSenti = "Positive";}
		    	else if (GeneratedSenti.equals("negative"))
		    	{GeneratedSenti = "Negative";}
		    	else if (GeneratedSenti.equals("neutral"))
		    	{GeneratedSenti = "Neutral";}
		    	else if (GeneratedSenti.equals("very positive"))
		    	{GeneratedSenti = "Very Positive";}
		    	else if (GeneratedSenti.equals("very negative"))
		    	{GeneratedSenti = "Very Negative";}
		    	
		    	if (MySenti.equals("positive"))
		    	{MySenti = "Positive";}
		    	else if (MySenti.equals("negative"))
		    	{MySenti = "Negative";}
		    	else if (MySenti.equals("neutral"))
		    	{MySenti = "Neutral";}
		    	else if (MySenti.equals("very positive"))
		    	{MySenti = "Very Positive";}
		    	else if (MySenti.equals("very negative"))
		    	{MySenti = "Very Negative";}
		    	

		    	
		    	if (GeneratedSenti.equals("Positive"))
		    	{tempScore1 = 3;}
		    	else if (GeneratedSenti.equals("Negative"))
		    	{tempScore1 = 1;}
		    	else if (GeneratedSenti.equals("Neutral"))
		    	{tempScore1 = 2;}
		    	else if (GeneratedSenti.equals("Very Positive"))
		    	{tempScore1 = 4;}
		    	else if (GeneratedSenti.equals("Very Negative"))
		    	{tempScore1 = 0;}
		    	
		    	
		    	if (MySenti.equals("Positive"))
		    	{tempScore2 = 3;}
		    	else if (MySenti.equals("Negative"))
		    	{tempScore2 = 1;}
		    	else if (MySenti.equals("Neutral"))
		    	{tempScore2 = 2;}
		    	else if (MySenti.equals("Very Positive"))
		    	{tempScore2 = 4;}
		    	else if (MySenti.equals("Very Negative"))
		    	{tempScore2 = 0;}
		    	
		    	Score = tempScore1 - tempScore2;
		    	
		    	if (Math.abs(Score) >= 2)
		    	{
		    		countOfNegPosErrors++;
		    	}
		    		
		    	
		    	Score = (int) Math.pow(Score, 2);
		    	
		    	totalScore = totalScore + Score;
		    	
		    	
		    	/////////SECOND METHOD//////////
		    	if (GeneratedSenti.isEmpty())
		    	{ 	}
		    	
		    	else
		    	{
		    		countTotal++;
		    		
		    		if (GeneratedSenti.equals(MySenti))
		    		{
		    			countAcc++;
		    		}
		    		
		    		else
		    		{
		    			Cell myCell3 = CellUtil.getCell(CellUtil.getRow(i, mySheet), MySentiColumnIndex + 1);
		    			myCell3.setCellValue("Bruh");
		    		}
		    		
		    	}
		    	
		    }
		    
		    System.out.println("Residual Square Score: " + totalScore);
		    
		    System.out.println();
		    
		    System.out.println("Number of errors between Pos and Neg: " + countOfNegPosErrors);
		    
		    System.out.println();
		    
		    System.out.println("Total Labels: " + countTotal);
		    System.out.println("Accurate Predictions: " + countAcc);
		    
		    
		    float MyAccuracy = (float) countAcc/countTotal;
		    System.out.println("Accuracy: " + MyAccuracy);
		    
		    myWorkBook.write(new FileOutputStream("Lenovo Sentiments FINAL With Err.xlsx"));
			myWorkBook.close();
		    
		    
	    }

		
	}

}
