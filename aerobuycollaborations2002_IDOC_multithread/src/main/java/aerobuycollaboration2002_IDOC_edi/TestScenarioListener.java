package aerobuycollaboration2002_IDOC_edi;
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLConnection;


public class TestScenarioListener {
	public static void  testSuiteName(String name, String filepath) throws Exception{
		URL url;
			  try {
				  if (name.equals("excelClose")) {
					 url = new URL("http://localhost:9080/automationserver/aerobuy/edib2b/provider/excelClose?filePath="+filepath);
					//  url = new URL("http://localhost:8090/automationserver/aerobuy/edib2b/provider/excelClose?filePath="+filepath+"&excelName=");
				}else {
					//url = new URL("http://localhost:9080/automationserver/aeroxchange/excelheaderwrite?filePath=output&product=aerocomponent&excelName="+name+"&module=validationRequired");
					url = new URL("http://localhost:8090/automationserver/aeroxchange/excelheaderwrite?filePath=output&product=aerobuy&excelName="+name+"&module=2002_idoc");			
									 
					}
				
		             URLConnection urlConnection = url.openConnection();
			         HttpURLConnection connection = null;
			         if(urlConnection instanceof HttpURLConnection) {
			            connection = (HttpURLConnection) urlConnection;
			         }else {
			            System.out.println("Please enter an HTTP URL.");
			            return;
			         }			         
			         BufferedReader in = new BufferedReader(new InputStreamReader(connection.getInputStream()));			      		         
			         in.close();			       
			      } catch (IOException e) {
			         e.printStackTrace();
			      }	
			  System.out.println("excel created successfully");

		  }
}