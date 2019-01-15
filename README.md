# DBSaveLoadWebService  
DB Excel Load/Save WebService  

Install SQLExpress 2017  
Install Microsoft SQL Server Managment Studio  
Connect to SQL Server, make note of server name if different then SQLExpress write it down  

Run Query

CREATE DATABASE Employee

USE Employee

CREATE TABLE Employee(
ID int PRIMARY KEY identity(1,1),
FirstName varchar(255),
LastName varchar (255),
Age int
)

INSERT INTO Employee(FirstName,LastName,Age) VALUES ('FirstTestName','LastTestName',16)
INSERT INTO Employee(FirstName,LastName,Age) VALUES ('SecondTestName','LastTestName',120)
INSERT INTO Employee(FirstName,LastName,Age) VALUES ('ThirdTestName','LastTestName',162)
INSERT INTO Employee(FirstName,LastName,Age) VALUES ('FourthTestName','LastTestName',11)
INSERT INTO Employee(FirstName,LastName,Age) VALUES ('FifthTestName','LastTestName',200)
INSERT INTO Employee(FirstName,LastName,Age) VALUES ('SixthTestName','LastTestName',30)
INSERT INTO Employee(FirstName,LastName,Age) VALUES ('SeventhTestName','LastTestName',82)


SELECT * FROM Employee

Install Visual Studio 2017 with Office/SharePoint Development Tools  
Download Service Solution  
If SQL server name change DBConString @"Data Source=.\sqlexpress; Initial Catalog=Employee; Integrated Security=True"  
To @"Data Source=.\-ServerName-; Initial Catalog=Employee; Integrated Security=True";  
Run WebService1.asmx.cs  
Use LoadDB webfunction  
If everything done correct it shows a XML with the Table Entries  
Leave Running  
Download DBSaveLoad Excel Addin solution  
Remove DBService from Connected Services  
Then Add Service reference  
Find http://localhost:55650/WebService1.asmx from the drop down menu  
Select WebService1Soap check if LoadDB and SaveDB operations are there  
If so  
Change namespace to DBService  
Check app.config for endpoint variable name if different then WebService1Soap copy into Ribbon1.cs client call variable  
Run  
Use Ribbon Load button  
Database should be loaded  
Make changes or add new data  
Use Ribbon Save button  

Done!
