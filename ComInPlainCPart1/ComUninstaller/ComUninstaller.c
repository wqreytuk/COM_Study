// Unregisters the IExample.DLL COM component.

#include <windows.h>
#include <tchar.h>
#include <initguid.h>
#include "../ComInPlainCPart1/IExample.h"



static const TCHAR	OurDllName[] = _T("IExample.dll");
static const TCHAR	ObjectDescription[] = _T("IExample COM component");
static const TCHAR	FileDlgTitle[] = _T("Locate IExample.dll to de-register it");
static const TCHAR	FileDlgExt[] = _T("DLL files\000*.dll\000\000");
static const TCHAR	ClassKeyName[] = _T("Software\\Classes");
static const TCHAR	CLSID_Str[] = _T("CLSID");
static const TCHAR	InprocServer32Name[] = _T("InprocServer32");
static const TCHAR	GUID_Format[] = _T("{%08lX-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}");






static void stringFromCLSID(LPTSTR buffer, IID ri)
{
	wsprintf(buffer, &GUID_Format[0],
		ri.Data1, ri.Data2, ri.Data3, ri.Data4[0],
		ri.Data4[1], ri.Data4[2], ri.Data4[3],
		ri.Data4[4], ri.Data4[5], ri.Data4[6],
		ri.Data4[7]);
}





/************************** WinMain() ************************
 * Program Entry point
 */

int WINAPI WinMain(HINSTANCE hinstExe, HINSTANCE hinstPrev, LPSTR lpszCmdLine, int nCmdShow)
{
		HKEY		rootKey;
		HKEY		hKey;
		HKEY		hKey2;
		TCHAR		buffer[39];

		stringFromCLSID(buffer, CLSID_IExample);

		// Open "HKEY_LOCAL_MACHINE\Software\Classes"
		if (!RegOpenKeyEx(HKEY_LOCAL_MACHINE, &ClassKeyName[0], 0, KEY_WRITE, &rootKey))
		{
			// Delete our CLSID key and everything under it
			if (!RegOpenKeyEx(rootKey, CLSID_Str, 0, KEY_ALL_ACCESS, &hKey))
			{
				if (!RegOpenKeyEx(hKey, buffer, 0, KEY_ALL_ACCESS, &hKey2))
				{
					RegDeleteKey(hKey2, InprocServer32Name);

					RegCloseKey(hKey2);
					RegDeleteKey(hKey, buffer);
				}

				RegCloseKey(hKey);
			}

			RegCloseKey(rootKey);
		}

		MessageBox(0, "De-registered IExample.DLL as a COM component.", &ObjectDescription[0], MB_OK | MB_ICONEXCLAMATION);


		return 0;
}
