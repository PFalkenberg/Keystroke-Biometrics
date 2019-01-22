import javafx.fxml.FXML;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;

public class MainController {
	
	@FXML
	private TextField fileName;
	
	@FXML
	private TextField trainingInstances;
	
	@FXML
	private TextField thresholdTxt;
	
	@FXML
	private TextArea display;
	
	public void generateScores(){
		String file = fileName.getText();
		int instances = Integer.parseInt(trainingInstances.getText());
		try{
	    	   
		       SpreadsheetReader reader = new SpreadsheetReader(file, 31, 400, 51);
		       
		       display.appendText("Generating templates...\n");
		       System.out.println("Generating templates...");
		       
		       reader.createAllTemplates(instances);		   
		       display.appendText("Templates created!\nCalculating Scores...\n");
		       System.out.println("Templates created!\nCalculating Scores...");
		       
		       reader.generateScores(instances);
		       
		       display.appendText("Scores Calculated!\nSaving Changes...\n");
		       System.out.println("Scores Calculated!\nSaving Changes...");
		       
		       
		       reader.saveScores();
		       display.appendText("Completed. Please use excel to sort the scores in ascending order,\n"
		       		+ " then save as xlsx and run the Calculate rates function.\n");
		         
	    	   
	       }
	       catch(Exception e){
	           System.out.println(e);
	       }
	}
	
	public void calculateRates(){
		String file = fileName.getText();
		int instances = Integer.parseInt(trainingInstances.getText());
		
		try{
			SpreadsheetReader reader = new SpreadsheetReader(file, 31, 400, 51);
			
			display.appendText("Generating Rates...\n");
			reader.generateRates(instances);
			display.appendText("Completed!");
			
		}
		catch(Exception e){
			System.out.println(e);
		}
		
	}
	
	public void ratesAt(){
		
		String file = fileName.getText();
		double threshold = Double.parseDouble(thresholdTxt.getText());
				
		try{
			SpreadsheetReader reader = new SpreadsheetReader(file, 31, 400, 51);
			
			display.appendText(reader.ratesAt(threshold));
			
		}
		catch(Exception e){
			System.out.println(e);
		}
	}

}
