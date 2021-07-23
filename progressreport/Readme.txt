Project Name : Excel Auto Progress Report Via Email

Application:
------------
This project helps you to automatically send student mark report in Excel format via email.

Used Dependencies:
------------------
<dependency>
	<groupId>org.apache.poi</groupId>
	<artifactId>poi</artifactId>
	<version>5.0.0</version>
</dependency>

<dependency>
	<groupId>org.apache.poi</groupId>
	<artifactId>poi-ooxml</artifactId>
	<version>5.0.0</version>
</dependency> 
 	
<dependency>
    <groupId>com.sun.mail</groupId>
    <artifactId>javax.mail</artifactId>
    <version>1.5.1</version>
</dependency>

<dependency>
    <groupId>javax.activation</groupId>
    <artifactId>activation</artifactId>
    <version>1.1.1</version>
</dependency>

Used Classes:
-------------
	1) StudentData:
	---------------
		This class has main method and setData method and carry the all student data in the main method.
		
		i)setData - Method
		------------------
			*This is retrieve the student inform from hash map and store it into strings.
			
			*Now, this method pass the strings to Report class getReport method that is return the excel data source.
			
			*Finally pass the data source to the ReportEmail class.
	2) Report:
	----------
		This is class only for create a excel sheet.

		i)getReport Method
		------------------
			*This method receiving the passed arguments and create a excel sheet and return the Excel Sheet Data Source. 
	1) SendEmail
	------------
		This class have two methods one is emailMessage other one is mailAuthentication
		
		i) mailAuthentication - Method
		------------------------------
			* This method only for Authentication user name and password is correct it will create a session. 
			
			* Using this session to send the email.
			
		ii) emailMessage - Method
		-------------------------
			* This method has 3 parameters one is for to email address other one is for name final one is for Attachment(This name splits form the email id).
			
			* Email message and Email subject given inside this method. if your want to change the message or subject of the email
			  simply replace the text what i given.

How to Send Email?
------------------
	1) Create a hash map for each student and put them all needed information into this hash map along with below given key values.
	
	Note: Don't change the key value.(If you are change then you will get exception) 
	
	Key values
	----------
		name (Student name)
		cls (Class)
		roleNumber(Reg.Number)
		email(Parents Email)
		english (English mark)
		tamil (Tamil Mark)
		maths (Maths Mark)
		social (Social Mark)
		science (Science Mark)
		
	2) Add the hash map object into the studentDetails Arrays list.
	
	3) Run the code