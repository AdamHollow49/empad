# empad
Employee Self-Service Active Directory data collection tool


## configuration
In future version I would like to include an installer however there is no demand for such a thing at the moment as it was created to use in-house by myself. Instead, on the off-chance anyone would like to use it I will include configuration notes.

You must have pyad, cryptography and pywin32 installed.

You must modify the following code sections:

### encryptIT.py
  'Key' variable must be created using cryptography - https://github.com/pyca/cryptography - the docs are really straight forward.
  'encrypted' variable must be set to the location where you will store your CSV file ( this stores encrypted employee data ).
### decryptIT.py
  'Key' variable must match the one created for encryptIT.py
  Path must be changed to where you wish the decrypted file to be placed. ( Line 19 )
### EmployeeAD.py
  'filePath' variable must be changed to the location you wish to store your user config files, these are created automatically by the application. ( Line 18 )
  'SERVICE_ACCNT' must be changed to the username of your service account inside your active directory. ( Line 19 )
  'LDAPSERVER' must be changed to your orgs LDAP server address. ( Line 19 )
  'Path_to_storage_folder' must be set to the location you wish to store your encrypted information. ( Line 96 )
  Please enter your internal service desk number. ( Line 263 )
  If your company has multiple locations, please specify in a list. ( Line 274 )
  You may include a picture of your companies logo here. ( Line 248 )
  
 ## Explanation
 
 This application connects to the active directory of your organisation, designed to be run on sign in, automatically logs in using the service account and targets the currently logged in user. At this point the user has the options to change a certain number of specified attributes. Employee Number, Login ID, Personal E-mail (optional), Location etc..
 
 This then sends the relevant information to active directory and any information that would be incorrect to store in that manner is encrypted to a file as a comma separated list. This can be decrypted using the decrpytIT script however this is designed to be only used by someone with clearance to do so. The idea of collecting this information is to store it in HR Databases, designed for holding personal data.
 
 If an employee refuses to fill out the form ( there is a small pop-up box that requests the employee's attention to fill out a form, they may say no ) then it is logged.
 Once an employee refuses to fill out this form 10 times, it sticks to their screen, always on top, and removes itself from the task bar. 
 When an employee fills out this form, they are then flagged as completed and are not asked to fill out this form for another 180 days.
 
 ## Custom Attributes
 If you wish to add further attributes to change - simply add them in line 187 onwards.
 The format that follows is : change_Attr("attribute_as_named_in_ad", var)
 With var being a variable storing the value you wish to assign.
 
