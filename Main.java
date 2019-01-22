import java.io.FileOutputStream;
import java.io.IOException;
import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.layout.StackPane;
import javafx.stage.Stage;


public class Main extends Application {
	
	@Override
	public void start(Stage primaryStage){
		try{
	        Parent root = FXMLLoader.load(getClass().getResource("Main.fxml"));
	        Scene scene = new Scene(root, 500, 300);
	        
	        primaryStage.setTitle("Keystroke Biometrics");
	        primaryStage.setScene(scene);
	        primaryStage.show();
	        }
	        catch(Exception e){
	            e.printStackTrace();
	        }
	}
	
	public static void main(String[] args) {
		
		launch(args);
		
	    }
	
}
