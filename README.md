# InjectorClassVB
DLL Injector Class used to hook into quake3 engine

Look into the InjectorClass.vb for the code
Usage:

Get the Process of your quake3 engine game and call this method

It will return a boolean true if done correctly.

C#
InjectorClass.DoInject(process, fileToInject, returnString);
VB.net
InjectorClass.DoInject(process, fileToInject, returnString)
