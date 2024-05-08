

#include <iostream>
#include <windows.h>

#include <atlbase.h> 


LPCOLESTR stringToPCOLESTR(const char* str)
{
    int length = static_cast<int>(strlen(str));
    int unicodeLength = static_cast<int>(MultiByteToWideChar(CP_ACP, 0, str, length, NULL, 0));
    wchar_t* unicodeStr = new wchar_t[unicodeLength + 1];
    MultiByteToWideChar(CP_ACP, 0, str, length, unicodeStr, unicodeLength);
    
    unicodeStr[unicodeLength] = L'\0';

    return reinterpret_cast<LPCOLESTR>(unicodeStr);
}


const char* BSTRtoString(BSTR bstr) 
{
    int size = WideCharToMultiByte(CP_ACP, 0, bstr, -1, NULL, 0, NULL, NULL);
    char* str = new char[size + 1];
    WideCharToMultiByte(CP_ACP, 0, bstr, -1, str, size, NULL, NULL);
    str[size] = '\0'; 
    return str;
}


/*
    El OCX debe estar registrado en HKEY_CLASSES_ROOT\CLSID
    Use: regsvr32 Archivo.ocx

*/
int main(int argv, char** argc)
{

    if (2 != argv)
    {
        std::cout << "Use: COMReader.exe ProgID" << std::endl;
        return 1;
    }

    auto chkError = [&](HRESULT hr, const std::string &desc = "")
        {
            if (FAILED(hr))
            {

                LPSTR messageBuffer = nullptr;
                size_t size = FormatMessageA(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                    NULL, hr, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPSTR)&messageBuffer, 0, NULL);
                
                std::string errmsg(messageBuffer, size);
                LocalFree(messageBuffer);

                std::cout << (desc.length() ? desc : "Error : ");
                std::cout << errmsg << std::endl;
                
                CoUninitialize();
                std::cin.get();
                ::exit(1);
            }
        };


    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr))
    {
        std::cout << "Error: No se pudo inicializar COM." << std::endl;
        return 1;
    }

    CLSID clsid;
    auto progID = stringToPCOLESTR(const_cast<const char*>(argc[1]));
    HRESULT hrClsid = CLSIDFromProgID(progID, &clsid);
    chkError(hrClsid, "Error: No se pudo obtener CLSID del OCX ");

    
    // Crear una instancia del objeto utilizando CoCreateInstance
    CComPtr<IUnknown> spUnknown;
    hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IUnknown, (void**)&spUnknown);
    chkError(hr, "Error: No se pudo crear una instancia del OCX ");

    CComQIPtr<IDispatch> spDispatch{ spUnknown };
    if (!spDispatch)
    {
        hr = spUnknown->QueryInterface(IID_IDispatch, (void**)&spDispatch);
        chkError(hr, "Error: No se pudo obtener la interfaz IDispatch");
    }


    // Obtener el número de miembros (métodos y propiedades)
    UINT numMembers = 0;
    hr = spDispatch->GetTypeInfoCount(&numMembers);
    chkError(hr, "Error: No se pudo obtener la información del tipo del OCX");

    std::cout << "Funciones Exportadas:" << std::endl;

    CComPtr<ITypeInfo> spTypeInfo;
    for (UINT i = 0; i < numMembers; i++)
    {        
        hr = spDispatch->GetTypeInfo(i, 0, &spTypeInfo);
        if (SUCCEEDED(hr) && spTypeInfo)
        {
            TYPEATTR* pTypeAttr;
            hr = spTypeInfo->GetTypeAttr(&pTypeAttr);
            if (SUCCEEDED(hr))
            {
                BSTR bstrName;
                hr = spTypeInfo->GetDocumentation(-1, &bstrName, NULL, NULL, NULL);
                if (SUCCEEDED(hr))
                {
                    std::cout << "* " << BSTRtoString(bstrName) << std::endl;
                    SysFreeString(bstrName);
                }

                spTypeInfo->ReleaseTypeAttr(pTypeAttr);
            }
        }

    }

    CoUninitialize(); 
    return std::cin.get();
}
