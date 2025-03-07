# SqlUserAdmin

**ClassAdm** is the primary user admin class. 
See comments in the module for setup of the server and front end.
<<<<<<< HEAD
The hidden table **uSysVar** stores encrypted keys.
=======
Uses hidden table **uSysVar** to store encrypted keys.
>>>>>>> f73c65638377589e6fbdbcb72b7c6c9f149556cb

**ClassCrypt** is a class version of Gustav Brock's modBcript cryptography code. 
                  https://github.com/GustavBrock/VBA.Cryptography/blob/main/LICENSE

<<<<<<< HEAD
The application uses the following functions **Decrypt**, **Encrypt**, **Hash**, **Random**, **RandomInt**. Supply these functions if you replace this cryptography with something else.
=======
The application uses the following functions **Decrypt**, **Encrypt**, **Hash**, **Random**, **RandomInt**. Supply these functions if you replace this cryptography with something else.
>>>>>>> f73c65638377589e6fbdbcb72b7c6c9f149556cb

Sample objects

**sysVar** is a sample linked table to the server containing 1 record indicating Test or Production database.

**feUser** is a sample application user list.

**modLog** is a basic logging function using debug.print, plus a function to generate the application key.

**Form_A** is a sample startup form used for user setup between the FE and server.

**Form_MainMenu** and **Form_UserList** are sample form/sub-form for working with the database.

<<<<<<< HEAD
I am using the **Joyfullservice** version control systems to generate these objects. 
It can be used to build the accdb from source. Check out the repository at https://github.com/joyfullservice/msaccess-vcs-addin
=======
>>>>>>> f73c65638377589e6fbdbcb72b7c6c9f149556cb
