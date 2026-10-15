package aerobuycollaboration2002_IDOC_edi;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import aerobuycollaboration2002_IDOC_edi.App;
import aerobuycollaboration2002_IDOC_edi.TestScenarioListener;
import aerobuycollaboration2002_IDOC_edi.excelEditor;
import aerobuycollaboration2002_IDOC_edi.filePathHandling;

public class scenarioTest {

	@Before
	public void beforeTest() {
		System.out.println("Aerobuy Execution Started");
	}

	@Test
	public void testCase() throws Exception {
		try {
			App.initialize();
			filePathHandling.aerobuyEDIMessagePath();			
			TestScenarioListener.testSuiteName("AeroBuy", "");
			TestScenarioListener.testSuiteName("OrderNumberDetailsExcel", "");
			TestScenarioListener.testSuiteName("MappingDetailsExcel", "");
			int threadPoolSize = 1; // <-- Define your thread numbers here
			mainWorkflow.workflowProcess(threadPoolSize);
			System.out.println("Excel close starting");
		//	 Thread.sleep(120000);
			// closing the order detail excel
			TestScenarioListener.testSuiteName("excelClose", filePathHandling.validationOrderDetailExcelPath);
			System.out.println("Excel close ending");
		//	Thread.sleep(60000);
			// verifying the output excel
			System.out.println("verification starting");
			excelEditor.verifyingOutputExcel();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@After
	public void afterTest() throws Exception {
		System.out.println("Aerobuyediusecase Execution Ended");
		email.emailTriggerProton("Aerobuy_collaboration_2002_IDOC", "298", "8");
	}
}
