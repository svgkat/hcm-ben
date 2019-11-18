# hcm-ben
Useful artifacts to ease workflow/debugging while implementing oracle cloud benefits

#Artifact Name: create_pbdr_v6.py
#Description: Creates a spreadhseet of the Person benefits diagnostic report. Input is the html version of the person benefits diagnostic report. The spreadheet will be created in the same directory with name pbdr_<xxxxx>_<yyyymmdd>_<hhmmss>.xlsx 
  xxxx: person number 
  yyyymmdd: date on which the spreadsheet was created
  hhmmss: timestamp
  
  Usage: python create_pbdr_v6.py Person_Benefits_Diagnostic_Test_004143.html
  
  Person_Benefits_Diagnostic_Test_004143.html is the person benefits diag report saved from cloud benefits.
