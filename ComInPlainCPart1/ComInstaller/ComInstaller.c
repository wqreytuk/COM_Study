// Registers our IExample.DLL as a COM component.

#include <windows.h>
#include <tchar.h>
#include <initguid.h>
#include "../ComInPlainCPart1/IExample.h"


// These first 3 strings will change if you create your own object
static const TCHAR	OurDllName[] = _T("IExample.dll");
static const TCHAR	ObjectDescription[] = _T("IExample COM component");
static const TCHAR	FileDlgTitle[] = _T("Locate IExample.dll to register it");

static const TCHAR	FileDlgExt[] = _T("DLL files\\000*.dll\\000\\000");
static const TCHAR	ClassKeyName[] = _T("Software\\Classes");
static const TCHAR	CLSID_Str[] = _T("CLSID");
static const TCHAR	InprocServer32Name[] = _T("InprocServer32");
static const TCHAR	ThreadingModel[] = _T("ThreadingModel");
static const TCHAR	BothStr[] = _T("both");
static const TCHAR	GUID_Format[] = _T("{%08lX-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}");



static void stringFromCLSID(LPTSTR buffer, IID ri)
{
	wsprintf(buffer, &GUID_Format[0],
		ri.Data1, ri.Data2, ri.Data3, ri.Data4[0],
		ri.Data4[1], ri.Data4[2], ri.Data4[3],
		ri.Data4[4], ri.Data4[5], ri.Data4[6],
		ri.Data4[7]);
}





/************************ cleanup() ***********************
 * Removes all of the registry entries we create in WinMain.
 * This is called to do cleanup if there's an error
 * installing our DLL as a COM component.
 */

static void cleanup(void)
{
	HKEY		rootKey;
	HKEY		hKey;
	HKEY		hKey2;
	TCHAR		buffer[39];

	stringFromCLSID(buffer, CLSID_IExample);

	// Open "HKEY_LOCAL_MACHINE\Software\Classes"
	if (!RegOpenKeyExA(HKEY_LOCAL_MACHINE, ClassKeyName, 0, KEY_WRITE, &rootKey))
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
}





/************************** WinMain() ************************
 * Program Entry point
 */

int WINAPI WinMain(HINSTANCE hinstExe, HINSTANCE hinstPrev, LPSTR lpszCmdLine, int nCmdShow)
{
	int				result;
	TCHAR			filename[MAX_PATH];

	{
		OPENFILENAME	ofn;

		// Pick out where our DLL is located. We need to know its location in
		// order to register it as a COM component
		lstrcpy(filename, OurDllName);
		ZeroMemory(&ofn, sizeof(OPENFILENAME));
		ofn.lStructSize = sizeof(OPENFILENAME);
		ofn.lpstrFilter = FileDlgExt;
		ofn.lpstrFile = filename;
		ofn.nMaxFile = MAX_PATH;
		ofn.lpstrTitle = FileDlgTitle;
		ofn.Flags = OFN_FILEMUSTEXIST | OFN_EXPLORER | OFN_PATHMUSTEXIST;
		// 通过弹窗的方式提示用户输入dll路径
		result = GetOpenFileName(&ofn);
	}

	if (result > 0)
	{
		HKEY		rootKey;
		HKEY		hKey;
		HKEY		hKey2;
		HKEY		hkExtra;
		TCHAR		buffer[39];
		DWORD		disposition;

		// Assume an error
		result = 1;

		// Open "HKEY_LOCAL_MACHINE\Software\Classes"
		if (!RegOpenKeyExA(HKEY_LOCAL_MACHINE, ClassKeyName, 0, KEY_WRITE, &rootKey))
		{
			// Open "HKEY_LOCAL_MACHINE\Software\Classes\CLSID"
			if (!RegOpenKeyExA(rootKey, CLSID_Str, 0, KEY_ALL_ACCESS, &hKey))
			{
				// 将IExample COM的clsid转换成字符串
				stringFromCLSID(buffer, CLSID_IExample);
				// 使用clsid创建subkey
				if (!RegCreateKeyExA(hKey, buffer, 0, 0, REG_OPTION_NON_VOLATILE, KEY_WRITE, 0, &hKey2, &disposition))
				{
					// 设置subkey的值(default)，描述信息
					RegSetValueExA(hKey2, 0, 0, REG_SZ, (const BYTE*)ObjectDescription, sizeof(ObjectDescription));

					// 创建subkey InprocServer32
					if (!RegCreateKeyEx(hKey2, InprocServer32Name, 0, 0, REG_OPTION_NON_VOLATILE, KEY_WRITE, 0, &hkExtra, &disposition))
					{
						// 设置InprocServer32的值(default)，即dll的路劲
						if (!RegSetValueEx(hkExtra, 0, 0, REG_SZ, (const BYTE*)filename, lstrlen(filename) + 1))
						{
							// 再添加一个值 ThreadingModel，值为 both
							if (!RegSetValueEx(hkExtra, ThreadingModel, 0, REG_SZ, (const BYTE*)BothStr, sizeof(BothStr)))
							{
								result = 0;
								MessageBox(0, "Successfully registered IExample.DLL as a COM component.", &ObjectDescription[0], MB_OK);
							}
						}

						// Close all keys we opened/created.

						RegCloseKey(hkExtra);
					}

					RegCloseKey(hKey2);
				}

				RegCloseKey(hKey);
			}

			RegCloseKey(rootKey);
		}

		// If an error, make sure we clean everything up
		if (result)
		{
			cleanup();
			MessageBox(0, "Failed to register IExample.DLL as a COM component.", &ObjectDescription[0], MB_OK | MB_ICONEXCLAMATION);
		}
	}

	return(0);
}
