# Biologic-COM-Interface
This repository goes through how to use OLE/COM with BT Lab on the Biologic to automate testing protocols with C++ or Python. This is essentially an unpolished tutorial on getting BTLab to communicate with coding languages through COM. This enables functionality not currently present on the Biologic such as batch assigning settings files to channels and running them at scale. This can vastly reduce the amount of time spent clicking through options for battery testing at scale. This tutorial is not an official Biologic tutorial, please contact a service representative if you have further questions and documentation.

1. First thing that you will need is the Windows SDK for applications such as OLE/COM Viewer, MIDL, etc.
2. Next depending on whether you will be using Python or C++ you will need to ensure that you have python installed, or C++ with a compiler such as Visual Studio. The rest of this will go through with Python. 
3. Launch the BTLab regserver with the command
4. You should now be able to open the registry editor and see it listed there under BTLabCom.BTLabExe
5. Open OLE/Viewer (oleviewer.exe) which should be located in :xxx
6. You should be able to find BTLabCOM and double click it
7. This should open the ITypeLib Viewer. From here you can File->Save As and select the destination folder where you are working in and save as a *.idl file.
8. The open your python IDE, you will need to have comtypes installed (I was unable to get pywin32 to work for this it is unclear if it IDispatchable which could be why it wasn't working, instead we need to go directly through the type library instead of trying to dispatch)
9. You will need to use MIDL.exe to generate the nesseccary type library files from the *.idl file. To do this you will want launch Developer Command Prompt for VS. This will launch a command prompt with the necessary environment for development and compilation, including access to tools like cl.exe and midl.exe instead of having to add these to your path.
10. In this console cd to where your idl file is, then call midl "insert_name.idl" and it will generate the *.tlb type library file.
11. With this file you will now be able to access the COM Module in both python and C++. 
