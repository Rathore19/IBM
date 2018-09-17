package myNLC;

import java.io.File;

import com.ibm.watson.developer_cloud.natural_language_classifier.v1.NaturalLanguageClassifier;
import com.ibm.watson.developer_cloud.natural_language_classifier.v1.model.Classification;
import com.ibm.watson.developer_cloud.natural_language_classifier.v1.model.Classifier;
import com.ibm.watson.developer_cloud.natural_language_classifier.v1.model.Classifiers;

public class mynlc 
{

	public static void main(String[] args) 
	{
		//CREDENTIALS
		NaturalLanguageClassifier service = new NaturalLanguageClassifier();
		service.setUsernameAndPassword("61dcbf0c-b649-47d3-a4eb-2723e6e087bd", "FU7LAaFTQTjX");
		
		//FOR TRAINING
		//Classifier classifier = service.createClassifier("WESA-Classifier2", "en", new File("Filename.csv")).execute();
		
		////Only CSV
		////No less than 5 records
		////UTF-8 format
		////No tabs - No endline etc
		////Record length limit
		
		
		//FOR ID
		//Classifiers classifiers = service.getClassifiers().execute();
		//System.out.println(classifiers);
		
	
		//DELETE
		//service.deleteClassifier("b91ef7x447-nlc-645").execute();
		
		
		//FOR CLASS
//		String st = "Good Job";
//		
//		Classification classification = service.classify("f7ec36x309-nlc-938", st).execute();
//		
//		
//		String myClass = classification.getTopClass();
//		
//		System.out.println(classification);
		

				
		System.out.println("End");
		
	}

}
