package progressreport;


import java.util.ArrayList;
import java.util.HashMap;

import javax.activation.DataSource;

public class StudentData {	
	Report report;
	ReportEmail reportEmail;
	public StudentData(Report report,ReportEmail reportEmail) {		
		this.report = report;
		this.reportEmail = reportEmail;
	}

	public static void main(String[] args) {
		
		//Object creation for used class		
		Report report = new Report();
		ReportEmail reportEmail = new ReportEmail();
		StudentData data = new StudentData(report,reportEmail);
		
		//Create HashMap to each student for storing a student data in Single Object.
		HashMap<String, String> reg_0001 = new HashMap<String, String>();
		HashMap<String, String> reg_0002 = new HashMap<String, String>();
		
		//Student Data
		reg_0001.put("name", "Ganu");
		reg_0001.put("cls", "10");
		reg_0001.put("roleNumber", "0001");
		reg_0001.put("email", "ganapathy8974.v2@gmail.com");
		reg_0001.put("english", "45");
		reg_0001.put("tamil", "26");
		reg_0001.put("maths", "46");
		reg_0001.put("social", "55");
		reg_0001.put("science", "71");
		
		//Student Data
		reg_0002.put("name", "Barath");
		reg_0002.put("cls", "10");
		reg_0002.put("roleNumber", "0002");
		reg_0002.put("email", "barathdhoni851@gmail.com");
		reg_0002.put("english", "10");
		reg_0002.put("tamil", "20");
		reg_0002.put("maths", "05");
		reg_0002.put("social", "56");
		reg_0002.put("science", "35");
		
		//Create a ArrayList for Store the Student HashMap Object.
		ArrayList<HashMap> st = new ArrayList<HashMap>();
		st.add(reg_0002);
		st.add(reg_0001);	
		
		//Call setData method in this class
		try {
			for(HashMap<String, String> mapData : st) {
				data.setData(mapData);
			}
		}catch(Exception ex) {
			ex.printStackTrace();
		}		
		
	}
	
	public void setData(HashMap<String, String> mapData) throws Exception {
		//Retrieve the Student data from the HashMap
		String name = mapData.get("name");
		String roleNumber= mapData.get("roleNumber");
		String cls = mapData.get("cls");
		String email = mapData.get("email");
		String []parentName = email.split("@");
		int english = Integer.parseInt(mapData.get("english"));
		int tamil =Integer.parseInt(mapData.get("tamil"));
		int maths =Integer.parseInt(mapData.get("maths"));
		int social = Integer.parseInt(mapData.get("social"));
		int science =Integer.parseInt(mapData.get("science"));		
		
		//Call the get report method in the report class pass the all needed data as argument
		DataSource excelReport = report.getReport(name, roleNumber, cls, english, tamil, maths, social, science);
		System.out.println(name+"'s mark report Generated.");
		
		//Call the get email message in the report eamil class pass the all needed data as argument
		reportEmail.MailAuthentication();
		reportEmail.EmailMessage(email, parentName[0], excelReport);
		
		System.out.println(" = Email Sent Successfully\n");
	}

}
