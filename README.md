# SqlUserAdmin

Manage user database accounts on SQL Server from an Access FE. Using cached ODBC connection in Access emulating Active Directory single sign-on.

I made this repository public as my first open-source project, hoping it will be of use and as a small give back, with the ulterior motive of improving the code for my production applications. I welcome  all feedback and suggestions.    

See Utter Access post.
https://www.utteraccess.com/topics/2065854/posts/2826059# And https://www.utteraccess.com/topics/2066313/posts/2826203 for early history on this project. You will see per suggestion; I replace my very weak cryptography with Gustav Brock's much improved system and moved the 1st ODBC connection inside the ClassAdm.

**ClassAdm** is the project's primary object; handling user administration while creating the cached ODBC connection. It uses the hidden table **uSysVar** to store encrypted keys. 
See comments in the class module for setup instructions for the SQL server and Access front end.
The class requires a cryptography class, using the following functions **Decrypt**, **Encrypt**, **Hash**, **Random**, **RandomInt**. 

The **SQLAdmSample** folder has SQL server scripts to build the sample database and a backup of my sample database.

**ClassCrypt** is a class version of Gustav Brock's BCript cryptography code included for the sample code but is not part of the project. 
                  (https://github.com/GustavBrock/VBA.Cryptography). You are free to supply your own class with compatible definitions. 
                  
Sample code objects.

**sysVar** is a sample linked table to the server containing one record indicating Test or Production database.

**feUser** is a sample application user list.

**modLog** is a basic logging function using debug.print, plus a function to generate the application key.

**Form_A** is a sample startup form used for user setup between the FE and server.

**Form_MainMenu** and **Form_UserList** are sample form/sub-form for working with the database.

I am using the **Joyfullservice** version control systems to generate these objects. It can be used to build the accdb from source. Check out the repository at https://github.com/joyfullservice/msaccess-vcs-addin
