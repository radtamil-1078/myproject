package aerobuycollaboration2002_IDOC_edi;

public class filePathHandling {

	// Home Dir and Default Path
	//public static String homeDirectory = "/home/automation/jenkins/aeroxchange/output/aerorepair/output/";
	//public static String homeDirectory = "/home/automation/jenkins/aeroxchange/output/";	
	public static String homeDirectory = "D:\\home\\automation\\jenkins\\aeroxchange\\output\\";	
	public static String defaultInputPath = "reference/input/data/";
	public static String defaultXmlInputPath = "reference/input/xml/";
	public static String defaultOutputPath = "reference/output/";
	public static String defaultXmlOutputPath = "reference/output/xml/";
	public static String defaultPropertiesPath = "properties/";
	public static String aerobuyEDIMessageDataExcelFileName,aerobuyEDIPropertiesFileName,aerobuyEDIMappingExcelName,A2002_idocMessagePropFileName,
	aerobuyEDIMessageOutputExcelFileName,aerobuyEDIScenarioExcelFileName,outputExcelPath,mappingExcelPath,validationOrderDetailExcelPath;
	
	public static void aerobuyEDIMessagePath() throws Exception {
		String module="2002_idoc";
		//String common = "aerobuy/output";
		String common = "aerobuy\\output";
		String currentDate = radiantGenericCode.getDateTime("currentDate", "MMddyyyy", 0);	
		A2002_idocMessagePropFileName = "credential.properties";
		aerobuyEDIPropertiesFileName = "info.properties";	
		aerobuyEDIMappingExcelName="MappingDetailsExcel.xlsx";	
		aerobuyEDIScenarioExcelFileName="scenario.xlsx";
		aerobuyEDIMessageDataExcelFileName = "DATA.xlsx";
		aerobuyEDIMessageOutputExcelFileName="AeroBuy.xlsx";		
		//outputExcelPath = homeDirectory +common+"/"+currentDate+"/"+module+"/"+aerobuyEDIMessageOutputExcelFileName;
		outputExcelPath = homeDirectory +common+"\\"+currentDate+"\\"+module+"\\"+aerobuyEDIMessageOutputExcelFileName;
		//mappingExcelPath = homeDirectory +common+"/"+currentDate+"/"+module+"/"+ aerobuyEDIMappingExcelName;
		mappingExcelPath = homeDirectory+common+"\\"+currentDate+"\\"+module+"\\"+ aerobuyEDIMappingExcelName;
		//validationOrderDetailExcelPath=homeDirectory +common+"/"+currentDate+"/"+module+"/" + "OrderNumberDetailsExcel.xlsx";
		validationOrderDetailExcelPath=homeDirectory+common+"\\"+currentDate+"\\"+module+"\\" + "OrderNumberDetailsExcel.xlsx";
	}

}
