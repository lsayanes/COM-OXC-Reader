# COM-OXC-Reader
Accessing a COM object (OCX) to read properties and functions

I had the need to access a COM object (OCX) from the code and list the available functions (methods and properties).

It is necessary to know the ProgID (Programmatic Identifier - readable name) of the COM object. 
This code assumes that the COM object implements the IDispatch interface and provides type information *(ITypeInfo)* for its members.

In addition, it uses the ATL classes *(CComPtr, CComQIPtr)* to simplify the management of COM interfaces and 
the automatic release of resources when exiting the scope.

This simple console program takes the ProgID as a parameter and allows to list and display
the names of the functions (methods and properties) available in a COM object (OCX) using the Windows API and standard COM interfaces. 
