# CIDI-GaDOE-Code
Consists of VBA macros and python scripts that were used in the automation of a manual data entry process

**FindNewUsers.vb**
Constraints - You can only hae two weeks worth of Users sheets present in the workbook. You have to clean the Users Sheet before applying this macro. 
Functionalities
1. Displays new users in message box
2. Modifies Users Weekly Sheet and Users Monthly Sheet with updated values

**FormattingBoldItalics.vb**
Constraints - you have to be on the sheet to use the specific macro. 
This macro adjusts the row and column heights/widths. It also makes the header bold. 

**UsersDataCleaning.vb**
Constraints - you have to be on the sheet to use the specific macro
This macro fixes the andrew gelinas, matthew blake, and oconee glrs records that are redundant. 

**findRequestsDiff.vb**
Constraints - You can only have two weeks of Requests sheets present in the workbook. 
This macro finds the difference between the two requests sheets. Then it updates the combined sheet with the new records. It also updates the Loans and Consults sheet. 
