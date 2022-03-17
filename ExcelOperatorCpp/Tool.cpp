#include "Tool.h"
namespace Tool
{
	using namespace std;
	std::string ConvertString(System::String^ str)
	{
		std::string result = msclr::interop::marshal_as<std::string>(str);
		return result;
	}
	System::String^ ConvertString(const std::string& str)
	{
		String^ result = gcnew String(str.c_str());
		return result;
	}

	std::wstring ConvertWString(System::String^ str)
	{
		std::wstring result = msclr::interop::marshal_as<std::wstring>(str);
		return result;
	}

	System::String^ ConvertWString(const std::wstring& str)
	{
		String^ result = gcnew String(str.c_str());
		return result;
	}

	std::wstring StringToWstring(const std::string& str)
	{
		if (str == "") return L"";
		wstring resstr;
		int size = MultiByteToWideChar(CP_ACP, 0, str.c_str(), str.size(), NULL, 0);
		wchar_t* ch = new wchar_t[size + 1];
		if (!MultiByteToWideChar(CP_ACP, 0, str.c_str(), str.size(), ch, size + 1))
		{
			resstr = wstring();
		}
		ch[size] = 0;
		resstr = wstring(ch);
		delete[]ch;
		return resstr;
	}

	std::string WstringToString(const std::wstring& wstr)
	{
		if (wstr == L"") return "";
		string mstr;
		int size = WideCharToMultiByte(CP_ACP, 0, wstr.c_str(), -1, NULL, 0, NULL, NULL);
		char* ch = new char[size + 1];
		if (!WideCharToMultiByte(CP_ACP, 0, wstr.c_str(), -1, ch, size, NULL, NULL))
		{
			mstr = string();
		}
		mstr = string(ch);
		delete[] ch;
		return mstr;
	}

}
