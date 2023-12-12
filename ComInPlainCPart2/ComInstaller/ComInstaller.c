// Registers our IExample.DLL as a COM component.

#include <windows.h>
#include <tchar.h>
#include <initguid.h>
#include "../ComInPlainCPart2/IExample.h"

BOOL installTypeExe = FALSE;

// These first 3 strings will change if you create your own object
static const TCHAR	OurDllName[] = _T("IExample.dll");
static const TCHAR	OurExeName[] = _T("IExampleLocalServer.exe");
static const TCHAR	TypeLibName[] = _T("IExample.tlb");
static const TCHAR	OurProgID[] = _T("IExample.object");
static const TCHAR	ObjectDescription[] = _T("IExample COM component");
static const TCHAR	FileDlgTitle[] = _T("Locate IExample.dll to register it");

static const TCHAR	FileDlgExt[] = _T("DLL files\\000*.dll\\000\\000");
static const TCHAR	FileDlgExtForExe[] = _T("EXE files\\000*.exe\\000\\000");
static const TCHAR	ClassKeyName[] = _T("Software\\Classes");
static const TCHAR	CLSID_Str[] = _T("CLSID");
static const TCHAR	InprocServer32Name[] = _T("InprocServer32");
static const TCHAR	LocalServer32Name[] = _T("LocalServer32");
static const TCHAR	ProgIDName[] = L"ProgID";
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
	// ProgID和CLSID都要删除
	// Open "HKEY_LOCAL_MACHINE\Software\Classes"
	if (!RegOpenKeyEx(HKEY_LOCAL_MACHINE, ClassKeyName, 0, KEY_WRITE, &rootKey))
	{
		// Delete our CLSID key and everything under it
		if (!RegOpenKeyEx(rootKey, CLSID_Str, 0, KEY_ALL_ACCESS, &hKey))
		{
			if (!RegOpenKeyEx(hKey, buffer, 0, KEY_ALL_ACCESS, &hKey2))
			{
				RegDeleteKey(hKey2, installTypeExe ? LocalServer32Name : InprocServer32Name);

				RegDeleteKey(hKey2, ProgIDName);

				RegCloseKey(hKey2);
				RegDeleteKey(hKey, buffer);
			}

			RegCloseKey(hKey);
		}

		// 删除OurProgID
		if (!RegOpenKeyEx(rootKey, OurProgID, 0, KEY_ALL_ACCESS, &hKey))
		{
			RegDeleteKey(hKey, CLSID_Str);
			RegDeleteKey(rootKey, OurProgID);
		
			RegCloseKey(hKey);
		}

		RegCloseKey(rootKey);

		// 注销TypLib
		UnRegisterTypeLib(&CLSID_TypeLib, 1, 0, LOCALE_NEUTRAL, SYS_WIN32);
	}
}





/************************** WinMain() ************************
 * Program Entry point
 */

int WINAPI WinMain(HINSTANCE hinstExe, HINSTANCE hinstPrev, LPSTR lpszCmdLine, int nCmdShow) {
	if (lstrlenA(lpszCmdLine) > 0) {
		// 说明要安装的是一个EXE
		installTypeExe = TRUE;
	}


	int				result;
	TCHAR			filename[MAX_PATH];

	{
		OPENFILENAME	ofn;

		// Pick out where our DLL is located. We need to know its location in
		// order to register it as a COM component
		lstrcpy(filename, installTypeExe ? OurExeName : OurDllName);
		ZeroMemory(&ofn, sizeof(OPENFILENAME));
		ofn.lStructSize = sizeof(OPENFILENAME);
		ofn.lpstrFilter = installTypeExe ? FileDlgExtForExe : FileDlgExt;
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
		if (!RegOpenKeyEx(HKEY_LOCAL_MACHINE, ClassKeyName, 0, KEY_WRITE, &rootKey))
		{
			// 需要在HKEY_LOCAL_MACHINE\Software\Classes下创建一个以我们的ProgID命名的Key
			if (!RegCreateKeyEx(rootKey, OurProgID, 0, 0, REG_OPTION_NON_VOLATILE, KEY_WRITE, 0, &hKey2, &disposition)) {
				stringFromCLSID(buffer, CLSID_IExample);
				disposition = RegSetValueEx(hKey2, 0, 0, REG_SZ, (const BYTE*)ObjectDescription, 2 * (lstrlen(ObjectDescription) + 1));
				// 还需要再在里面创建一个Key，键名为CLSID
				if (!disposition && !RegCreateKeyEx(hKey2, CLSID_Str, 0, 0, REG_OPTION_NON_VOLATILE, KEY_WRITE, 0, &hkExtra, &disposition)) {
					disposition = RegSetValueEx(hkExtra, 0, 0, REG_SZ, (const BYTE*)buffer, 2 * (lstrlen(buffer) + 1));
					RegCloseKey(hkExtra);
				}
				RegCloseKey(hKey2);
			}

			// Open "HKEY_LOCAL_MACHINE\Software\Classes\CLSID"
			if (!disposition && !RegOpenKeyEx(rootKey, CLSID_Str, 0, KEY_ALL_ACCESS, &hKey))
			{
				// 将IExample COM的clsid转换成字符串
				stringFromCLSID(buffer, CLSID_IExample);
				// 使用clsid创建subkey
				if (!RegCreateKeyEx(hKey, buffer, 0, 0, REG_OPTION_NON_VOLATILE, KEY_WRITE, 0, &hKey2, &disposition))
				{
					// 设置subkey的值(default)，描述信息
					RegSetValueEx(hKey2, 0, 0, REG_SZ, (const BYTE*)ObjectDescription, sizeof(ObjectDescription));

					// 创建subkey InprocServer32
					if (!RegCreateKeyEx(hKey2, installTypeExe ?  LocalServer32Name: InprocServer32Name, 0, 0, REG_OPTION_NON_VOLATILE, KEY_WRITE, 0, &hkExtra, &disposition))
					{
						// 设置InprocServer32的值(default)，即dll的路劲
						// Unicode两倍字节数
						if (!RegSetValueEx(hkExtra, 0, 0, REG_SZ, (const BYTE*)filename, 2*(lstrlen(filename) + 1)))
						{
							// 再添加一个值 ThreadingModel，值为 both
							disposition = RegSetValueEx(hkExtra, ThreadingModel, 0, REG_SZ, (const BYTE*)BothStr, 2 * (lstrlen(BothStr) + 1));
						}

						// Close all keys we opened/created.

						RegCloseKey(hkExtra);

						// 创建ProgID
						if (!disposition && !RegCreateKeyEx(hKey2, ProgIDName, 0, 0, REG_OPTION_NON_VOLATILE, KEY_WRITE, 0, &hkExtra, &disposition)) {
							// 给ProgID赋值
							if (!RegSetValueEx(hkExtra, 0, 0, REG_SZ, (const BYTE*)OurProgID, 2 * (lstrlen(OurProgID) + 1)))
								result = 0;
						}
						RegCloseKey(hkExtra);
					}

					RegCloseKey(hKey2);
				}

				RegCloseKey(hKey);
			}

			RegCloseKey(rootKey);
		}

		MessageBoxA(NULL, "hold", "hold", MB_OK);
		// 注册TypeLib，默认tlb文件和dll文件位于同一目录
		if (!result) {
			ITypeLib* pTypeLib;
			LPTSTR str;
			
			// 通过\\截取目录路径
			str = filename + lstrlen(filename);
			while (str > filename && *(str-1) != L'\\') str--;
			lstrcpy(str, TypeLibName);
			// 现在filename就是TLB文件的完整路径了
			if (SUCCEEDED(result = LoadTypeLibEx(filename, REGKIND_REGISTER, &pTypeLib))) {
				// 这个对象用完要记得释放
				pTypeLib->lpVtbl->Release(pTypeLib);
			}
			if (!result)
				MessageBox(0, L"Successfully registered IExample.DLL as a COM component.", &ObjectDescription[0], MB_OK);
		}



		// If an error, make sure we clean everything up
		if (result)
		{
			cleanup();
			MessageBox(0, L"Failed to register IExample.DLL as a COM component.", &ObjectDescription[0], MB_OK | MB_ICONEXCLAMATION);
		}
	}

	return 0;
}
