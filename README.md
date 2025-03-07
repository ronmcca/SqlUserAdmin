# SqlUserAdmin

**ClassAdm** is the primary user admin class. 
See comments in the module for setup of the server and front end.
Uses hidden table **uSysVar** to store encrypted keys.

**ClassCrypt** is a class version of Gustav Brock's modBcript cryptography code. 
                  https://github.com/GustavBrock/VBA.Cryptography/blob/main/LICENSE

he application uses the following functions **Decrypt**, **Encrypt**, **Hash**, **Random**, **RandomInt**. Supply these functions if you replace this cryptography with something else.

Sample objects

**sysVar** is a sample linked table to the server containing 1 record indicating Test or Production database.

**feUser** is a sample application user list.

**modLog** is a basic logging function using debug.print, plus a function to generate the application key.

**Form_A** is a sample startup form used for user setup between the FE and server.

**Form_MainMenu** and **Form_UserList** are sample form/sub-form for working with the database.

