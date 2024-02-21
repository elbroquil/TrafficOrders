# Traffic Orders

## Introduction to the Project
The spreadsheet and templates were used to track, process and file Traffic Orders for the Royal Borough of Greenwich (RBG). 

While the Projects spreadsheet was a pre-existing projects tracker, I have set up a VBA user form/macro, email templates, and file structure to keep track of Traffic Orders progress. 

The Traffic Order process consisted of:
* Receiving a commission for a Traffic Order (project manually logged in the spreadsheet)
* Drafting the Order (legal documents, drafted manually and in ParkMap, a MapInfo extension)
     + Sending the Order to the RBG engineers for approval (automated Step 1, subtab 1 of the Form tab)
     + Sending the Order publication requests to newspapers (automated Step 2, subtab 2 of the Form tab)
         + Generating newspaper booking forms (automated)
* Sending notices to statutory stakeholders (automated Step 4, subtab 5 of the Form tab)
* Sending a reminder to the engineer that the objection period has expired, asking whether the Order should be Made (become enforceable)

If so, Making the Order included (as above):
* Sending the Order to the RBG engineers for approval (automated Step 1, subtab 1 of the Form tab)
* Sending the Order publication requests to newspapers (automated Step 2, subtab 2 of the Form tab)
    + Generating newspaper booking forms (automated)

Automating Traffic Order admin saved time and limited human errors (as legal documents, any errors in the consultation or Making process meant that the Order was not valid - if you ever get a parking ticket in the UK question adherence to the legal procedure in making the traffic law, the fine will likely be removed!).

As the person responsible for RBG admin the macro was initially a tool to make my life easier, however it was passed on to the person taking over when I was leaving the company - hence the instructions tabs built into the user form. 

## Instructions
Windows only, might run into ActiveX issues for email/word doc generation on other operating systems.

Open the Projects spreadsheet. The spreadsheet includes a redacted list of Traffic Order projects. Click on the 'Open the User Form' button at the bottom.
![alt text](https://github.com/elbroquil/TrafficOrders/blob/main/InstructionPictures/Open%20UserForm.png)

Go to the 'Form' tab. 
Type in row number of any a sample Order (rows 2 to 8 for Orders in the proposal stage, rows 9 to 15 for Orders in the Making stage) and click 'Autofill'. All details will be automatically pulled into the user tab. The macro deals with whether the Order is in the consultation or Making stage (NoP/NoM)).

Depending on the stage, go to the appropriate tab at the bottom of the Form.  
![alt text](https://github.com/elbroquil/TrafficOrders/blob/main/InstructionPictures/Filled%20in%20UserForm.png)

Clicking 'Email' will fill in the email template, attach the required Order documents and address it to the right people. 
The email will open in new window for review. __In test scenarios please do not hit send! The Metropolitan Police and TfL might be a bit surprised if you do...__

Ticking/unticking checkboxes will add/clear a 'Y' in the progress tracking columns.

For the 'Newspapers' tab, clicking 'Generate Forms' will copy, prefill and save booking forms. A message will be displayed and Word will be left open to review the documents.

![alt text](https://github.com/elbroquil/TrafficOrders/blob/main/InstructionPictures/Generate%20Newspaper%20Forms.png)

![alt text](https://github.com/elbroquil/TrafficOrders/blob/main/InstructionPictures/Prefilled%20Form.png)


       
