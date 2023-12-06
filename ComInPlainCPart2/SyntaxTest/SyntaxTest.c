#include <Windows.h>

int main()
{
	// 这个就是BSTR自动变量的格式，前两个WORD是小端格式的字节长度 9个字符*2bytes=18  不计算终止WORD 0
	WORD MyString[] = { 18, 0, 'S','o','m','e',' ','t','e','x','t', 0 };
	BSTR  strptr;
	// 实际开始位置是真正字符开始的地方，前两个WORD不是
	// 虽然BSTR指向的是字符串开始的地方，但是向前移动4bytes，还是有值的，就是长度，除以2即可得到字符数量
	// 不得不说这个格式确实逆天
	strptr = &MyString[2];

	// MultiByteToWideChar 可以将 ansi c string 转换成 每个字符串占用2bytes的UNICODE string

	// SysAllocStringLen  可以用于给unicode string 分配内存，一个字符俩字节，并且还会根据
	// 传入的参数len自动把前缀DWORD（字符串所占字节数）添加在最前面


	DWORD len;
	BSTR strPtr;
	char ansiString = "Some String";
	len = MultiByteToWideChar(CP_ACP, 0, ansiString, -1, 0, 0);
}