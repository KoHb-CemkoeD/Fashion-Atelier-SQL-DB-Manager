# Fashion-Atelier-Manager
Project for developing a software system designed for managing the database of the "Fashion Atelier" enterprise. The system will include data about available fabrics, a list of cutters, existing clothing models, ready-to-issue products, and current orders.

## Objective
The objective is to create software for data management at the "Fashion Atelier" enterprise, including data management for available fabrics, cutters, fabric purchases, clothing models, finished products, and orders. This software is intended for use by employees of the "Fashion Atelier" to quickly and conveniently input, search, and process data related to fabrics, cutters, clothing models, finished products, and orders, including user identification capabilities.

## Mathematical Models
Utilized set theories and graphs to formalize and optimize data processing processes within the enterprise system.
Based on the input data on the number of tissues, their types, and the list of possible tissue operations, it is necessary to form a formalized data structure of the organization, including a set of possible tissue processing operations.

![image](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/12ad07d3-ab54-48d5-8026-8c1b347979df)

Formula below shows the mathematical model of tissue processing - MOI in the electronic data management system in tuple form:

![image](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/3f12d4b3-138d-446b-a3dc-189234b5af06)

where S is the structure of the organization's electronic data circulation, which will be understood as a formalized representation in the form of a tuple model of the set of fabric circulation objects U that change their state as a result of operations O by a set of cutters P. 

## Data Structure
Formulated a structured data organization reflecting fabric processing operations based on input data about fabric quantities, types, and processing operations.

After the system initialization, the program is in the standby mode until the user performs any action, namely: adding, deleting, editing, forming a sample of records. The figure shows a functional diagram that clearly describes the behavior of the system while the program is running.

![image](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/75a4d6e2-5c19-41b0-9996-8d1b26b84122)

The block diagram of the "Fashion Atelier" software complex, which demonstrates the interaction of the functions of the main program, is shown in the figure.

![image](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/ae48f1eb-ce5e-402c-b5c2-38f5bb1642d4)


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

![0](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/3f88f0d8-29f3-44a0-806d-d19bc2ef3ea8)

- The fabric information table is used to keep track of the available fabrics in the atelier. It includes fields such as article, type, price per meter, width, remains, and picture.

- The cutters table contains data about the personnel who sew fabric products in the atelier. Fields include code, name, phone, address, specialization, and work experience.

- The purchases table records fabric purchases made by the atelier, with fields such as code, date, fabric article, length, and supplier.

- The models table contains data on the fabric models currently being sewn in the atelier. Fields include model article, purpose, specification, and size.

- The finished products table records information about fabric products before delivery, including code, model, size, customer, date ready, and material.

- The orders table contains data on current and in-progress orders for fabric products. Fields include number, date, customer, cutter, name, size, and issued.

## User Interface
Utilized visual forms and Qt Designer for UI layout and design. The main widget of the user interface is the "MainWindow" class, which creates the main form of the "Fashion Atelier" application. The form contains the main menu, tabs with tables for displaying information, fields for entering data, pop-up lists, radio, and regular buttons for interacting with the user.

![0](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/3f137894-1592-40fc-96f9-7eca9c058c70)

Main form. Generating reports by parameters

![0](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/940fc447-5ab9-496f-8b57-2110e6c65e34)


The class contains a constructor that loads the visual form from a file, connects the controls to the appropriate processing functions. The class also contains methods for connecting, loading the database; displaying and updating the data of the current table; methods for tracking and responding to user actions when selecting menu items, tables, query types, and reports. 

## Reports Creationg
Implemented report generation in Microsoft Word, Excel, and PDF formats.

### Report for the period

![image](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/901348ee-9e72-429a-ada5-e23c8129aaf2)

Contents of the generated order report for the period in PDF format

![0](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/c43f7dc0-bd8e-43f0-811f-00d45ed08951)

### Reports by customer

![image](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/b33a9f22-85f0-4839-89c0-de80b348ea31)

Contents of the finished goods report by customer in MS Word format 

![0](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/07c0746c-7e4d-483b-a8f4-780b540969fa)

### Reports on product models 

![image](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/efd18a0a-a324-48c1-8779-9d4b52901136)
 
Contents of the generated report on product models in MS Excel format 

![0](https://github.com/KoHb-CemkoeD/Fashion-Atelier-Manager/assets/32577543/38d2066a-9055-48ef-bac1-dcfac897a091)


## Implementation Details
- Programming Language: Python
- Additional libraries: PyQt5
- DB: Microsoft Access via library win32com

## User Guide
Included a user guide document ("user_guide.doc") within the program folder, providing instructions on running the software, required input data format, and operational guidelines.

## Conclusion
The project successfully developed a software system for managing data related to fabrics, cutters, fabric purchases, clothing models, finished products, and orders at the "Fashion Atelier." By leveraging database capabilities, the system ensures reliable data storage and convenient data access, demonstrating the versatility and effectiveness of database functionalities in system description and management.
