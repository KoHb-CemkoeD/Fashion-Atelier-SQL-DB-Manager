# Fashion-Atelier-SQL-DB-Manager
Project for developing a software system designed for managing the MS ACCESS database of the "Fashion Atelier" enterprise using SQL queries. The system includes data about available fabrics, a list of cutters, existing clothing models, ready-to-issue products, and current orders.

## Objective
The objective is to create software for data management at the "Fashion Atelier" enterprise, including data management for available fabrics, cutters, fabric purchases, clothing models, finished products, and orders. This software is intended for use by employees of the "Fashion Atelier" to quickly and conveniently input, search, and process data related to fabrics, cutters, clothing models, finished products, and orders, including user identification capabilities.

## Mathematical Models
Utilized set theories and graphs to formalize and optimize data processing processes within the enterprise system.
Based on the input data on the number of tissues, their types, and the list of possible tissue operations, it is necessary to form a formalized data structure of the organization, including a set of possible tissue processing operations.

<img width="371" height="121" alt="image" src="https://github.com/user-attachments/assets/65eca8ad-c19e-4b14-8623-e4a04c68faba" />


Formula below shows the mathematical model of tissue processing - MOI in the electronic data management system in tuple form:

<img width="129" height="20" alt="image" src="https://github.com/user-attachments/assets/546e13f6-16ea-47b3-8d7d-4a8ca9caa5b1" />

where S is the structure of the organization's electronic data circulation, which will be understood as a formalized representation in the form of a tuple model of the set of fabric circulation objects U that change their state as a result of operations O by a set of cutters P. 

## Data Structure
Formulated a structured data organization reflecting fabric processing operations based on input data about fabric quantities, types, and processing operations.

After the system initialization, the program is in the standby mode until the user performs any action, namely: adding, deleting, editing, forming a sample of records. The figure shows a functional diagram that clearly describes the behavior of the system while the program is running.

<img width="379" height="261" alt="image" src="https://github.com/user-attachments/assets/c0e83fa2-dc25-4ef7-a566-38fb4c95cae0" />

The block diagram of the "Fashion Atelier" software complex, which demonstrates the interaction of the functions of the main program, is shown in the figure.

<img width="353" height="213" alt="image" src="https://github.com/user-attachments/assets/9b73d003-f75f-475d-99d8-68260c9eb52f" />


## Class Description
Designed main program classes including "Person," "Cutter," "Order," and "Model" to store and manage relevant data efficiently.

![image](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/06d7fbe1-d0b4-4184-ba14-f92388fc6e4d)

### Person 
Class stores information about a person, including their full name, address, and phone number. It has methods for setting and editing this information.

### Cutter
Class extends the "Person" class and adds fields for specialization and work experience. It also inherits the methods for setting and editing basic personal information from the "Person" class.

### Order
Class stores information about a specific order of a fabric product. It includes fields for the order number, date, customer, cutter, product name, size, and order status. It has methods for setting and editing this information.

### FinishedProduct 
Class stores information about finished products. It includes fields for the product code, model, size, customer, date of manufacture, and material. It has methods for setting and editing this information, as well as checking if the record meets specified conditions.

### MainMenu
Class is the main widget that provides the program's interface and allows the user to choose which data to work with. It includes fields for temporary tables, lists of objects for different data categories (such as finished products, orders, and cutters), and methods for displaying the main form of the application and handling the database.

Additionally, there is an exception handling class to handle errors when setting incorrect arguments or accessing non-existent items in the database.


## DB Description. UML Scheme
The information in DataBase include fabrics, cutters, purchases, models, finished products, and orders. Each table is described with its corresponding fields and data types.

<img width="584" height="229" alt="image" src="https://github.com/user-attachments/assets/7016b71a-8557-4865-a7fa-03e55b858dd0" />

- The fabric information table is used to keep track of the available fabrics in the atelier. It includes fields such as article, type, price per meter, width, remains, and picture.

- The cutters table contains data about the personnel who sew fabric products in the atelier. Fields include code, name, phone, address, specialization, and work experience.

- The purchases table records fabric purchases made by the atelier, with fields such as code, date, fabric article, length, and supplier.

- The models table contains data on the fabric models currently being sewn in the atelier. Fields include model article, purpose, specification, and size.

- The finished products table records information about fabric products before delivery, including code, model, size, customer, date ready, and material.

- The orders table contains data on current and in-progress orders for fabric products. Fields include number, date, customer, cutter, name, size, and issued.

## User Interface
Utilized visual forms and Qt Designer for UI layout and design. The main widget of the user interface is the "MainWindow" class, which creates the main form of the "Fashion Atelier" application. The form contains the main menu, tabs with tables for displaying information, fields for entering data, pop-up lists, radio, and regular buttons for interacting with the user.

<img width="554" height="329" alt="image" src="https://github.com/user-attachments/assets/57e7630b-abdc-4a3e-852a-90e5e7263113" />

Main form. Generating reports by parameters

<img width="610" height="359" alt="image" src="https://github.com/user-attachments/assets/9efd8237-69d7-4cd1-8ae9-c782e608f596" />


The class contains a constructor that loads the visual form from a file, connects the controls to the appropriate processing functions. The class also contains methods for connecting, loading the database; displaying and updating the data of the current table; methods for tracking and responding to user actions when selecting menu items, tables, query types, and reports. 

## Reports Creating
Implemented report generation in Microsoft Word, Excel, and PDF formats.

### Report for the period

![image](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/901348ee-9e72-429a-ada5-e23c8129aaf2)

Contents of the generated order report for the period in PDF format

<img width="584" height="454" alt="image" src="https://github.com/user-attachments/assets/3ae48a80-568e-48e2-9132-6a49504e1f70" />

### Reports by customer

![image](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/b33a9f22-85f0-4839-89c0-de80b348ea31)

Contents of the finished goods report by customer in MS Word format 

<img width="491" height="332" alt="image" src="https://github.com/user-attachments/assets/309efbc0-d26c-456a-a606-f538d67cb589" />

### Reports on product models 

![image](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/efd18a0a-a324-48c1-8779-9d4b52901136)
 
Contents of the generated report on product models in MS Excel format 

<img width="397" height="261" alt="image" src="https://github.com/user-attachments/assets/36b37120-9df8-48ba-a736-78e43e479aee" />

### View query execution statistics 

![0](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/2ca754c7-e4b8-4426-9a42-0691ad2c3350)


## Implementation Details
- Programming Language: Python
- Additional libraries: PyQt5
- DB: Microsoft Access via library win32com

## User Guide
Included a user guide document ("user_guide.doc") within the program folder, providing instructions on running the software, required input data format, and operational guidelines.

<img width="343" height="552" alt="image" src="https://github.com/user-attachments/assets/6a1cdefe-1a6f-4da2-a2de-b84facbf8df9" />


## Conclusion
The project successfully developed a software system for managing data related to fabrics, cutters, fabric purchases, clothing models, finished products, and orders at the "Fashion Atelier." By leveraging database capabilities, the system ensures reliable data storage and convenient data access, demonstrating the versatility and effectiveness of database functionalities in system description and management.
