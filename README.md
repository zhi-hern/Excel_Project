# Excel_Project
This project showcases my proficiency in querying data using Excel.

## Manpower Tracker
[[Manpower tracker.xlsx]](https://github.com/zhi-hern/Excel_Project/blob/main/Manpower%20tracker.xlsx)

## Project Description
This project creates a summary for management to keep track of manpower. 

### Problem Statement:
Management need to make sure 3 shifts and 2 roles are filled everyday. Hence, they created a rotation table to keep track of the workers' shifts, rests days and leaves. 
![image](https://github.com/user-attachments/assets/76cb99e9-6cf0-4c2c-8711-14fef9b3d313)
But, the information is not well-accessible. 

### Project Details:
| Row name | Description |
| -------- | ----------- |
| Shift    | Shift that workers work on <br> (D1: 6am-2pm, D2: 2pm-10pm, D3: 10pm-6am)|
| Status    | Workers on Sick Leave (SL)/ Annual Leave (AL)/ Outstation (OS)|
| Role    | Role that workers need to perform on the shift <br> (Role 1/ Role 2) |
| Cover By | Names that cover that particular shift | 

### Solution:
A manpower tracker was created by querying the rotation table created by management.
By inputting the week number, management can easily accessed the rotation's information. 
![image](https://github.com/user-attachments/assets/9e32e976-7697-4ef8-88df-560a4281e977)

From this, management can easily know that no one is working on **_2/1/2025 & 3/1/2025_** of **_D2's ROLE 2_**. <br>

As for 31/12/2024 and 1/1/2025, Yusof will be covering the 1st half of D2 shift (continuing from his D1 shift) and Murphy will be covering the 2nd half of D2 shift (and continue working for his D3 shift).

### Task Performed:
A "Transform sheet" is created to query the "Rotation" table by using `'HLOOKUP'` function and determine who's on shift using `'IF'` and `'Nested IFs'` statements.

To determine who's on duty (scheduled worker/coverage/no one), a new row `"On Duty"` "is created using `"Nested IFs"`based on the hierachy as shown below:

![image](https://github.com/user-attachments/assets/11bc6ff2-441d-4ccd-9eda-a86e13dd7332)

A date sheet is created to display the dates based on week number. 

`"Index"` and `"Match"` are used to create the summary of the rotation by matching the DATE, SHIFT and ROLE.

Example of code:
```
=INDEX("On Duty" row,MATCH(1,(DATE ='DATE' from 'Transform Sheet')*(SHIFT = 'SHIFT' from 'Transform Sheet')*(ROLE = 'ROLE' from 'Transform Sheet'),0))
```

