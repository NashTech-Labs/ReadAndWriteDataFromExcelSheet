# Read and Write Data from Excel Sheet Using Apache POI
This template will help you to have a Read the data from Excel sheet and also Write to Data to Excel sheet Using Apache POI.
## What is the Apache POI
* Apache POI is basically an open source Java library developed by Apache which provides an API for read and write data in Excel sheet using Java programs. 
* It has classes and methods to decode the user input data or a file into Excel file. So for data driven testing using Selenium Web Driver, we use it for reading data stored in excel sheets. 
* It can also be used if you want to write any data to an excel file in your scripts.

## I was using
* Java as the programming language
* TestNG as the assertion framework
* Maven as the build tool
* IntelliJ as the IDE

#### Type of framework - Data driven framework.
## New package--
### ReadData class -
* This class is responsible for loading the test cases from excel sheet.
* Download Apache POI libraries dependency in pom files.
* XSSF(XML Spreadsheet Format) this is the interface using for read and write the data from excel sheet using XSSFWorkbook workbook.
* Here we Read the data from Excel sheet

### WriteData class -
* This class is handle all value to store in Excel sheet through Selenium Script.
* Firstly we create a sheet.
XSSFSheet spreadsheet = workbook.createSheet(" Employee Data ");
* After we create a row object
XSSFRow row;
* After that we need to Write the data in script using map function.
Map<String, Object[]> studentData = new TreeMap<String,Object[]>();

#### Create a Testng.xml file so there we add both of test
#### Run the Testng file.

### Apache POI  dependency we use in this framework:-
* To read/Write data from an excel file we need Apache POI maven dependency

			<dependency>
				<groupId>org.apache.poi</groupId>
					<artifactId>poi</artifactId>
					<version>3.9</version>
			</dependency>
			<dependency>
					<groupId>org.apache.poi</groupId>
					<artifactId>poi-ooxml</artifactId>
					<version>3.9</version>
			</dependency>

## Result- 		

![Screenshot from 2021-03-19 18-36-13](https://user-images.githubusercontent.com/43197101/111787211-ae1d2a00-88e4-11eb-8624-fedbd526cc69.png)

![Screenshot from 2021-03-19 18-15-42](https://user-images.githubusercontent.com/43197101/111787189-a9587600-88e4-11eb-9519-4f029b73e4d4.png)



