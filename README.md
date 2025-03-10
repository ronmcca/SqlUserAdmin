# SqlUserAdmin

Manange user database accounts on SQL Server from an Access FE. Using cached ODBC connection in Access emulating Active Directory single sign-on.

I have moved it to public as my first open source project and welcome feedback and suggestions, with the ulterior modive of improving the code for my production applications.  

See Utter Access post
https://www.utteraccess.com/topics/2065854/posts/2826059# And https://www.utteraccess.com/topics/2066313/posts/2826203 for early history on this project.

**ClassAdm** is the primary user admin class. 
See comments in the module for setup of the server and front end.
Hidden table **uSysVar** stores encrypted keys.

**ClassCrypt** is a class version of Gustav Brock's modBcript cryptography code and is included for the sample code but is not part of the project. 
                  https://github.com/GustavBrock/VBA.Cryptography/blob/main/LICENSE

The application uses the following functions 
**Decrypt**, **Encrypt**, **Hash**, **Random**, **RandomInt**. 
Supply these functions if you replace this cryptography with something else.

Sample objects

**sysVar** is a sample linked table to the server containing 1 record indicating Test or Production database.

**feUser** is a sample application user list.

**modLog** is a basic logging function using debug.print, plus a function to generate the application key.

**Form_A** is a sample startup form used for user setup between the FE and server.

**Form_MainMenu** and **Form_UserList** are sample form/sub-form for working with the database.

I am using the **Joyfullservice** version control systems to generate these objects. It can be used to build the accdb from source. Check out the repository at https://github.com/joyfullservice/msaccess-vcs-addin
