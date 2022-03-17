#pragma once
#include <string>
#include <msclr/marshal_cppstd.h>
using namespace System;
namespace Tool
{
	std::wstring StringToWstring(const std::string& str);
	std::string WstringToString(const std::wstring& wstr);
	std::string ConvertString(System::String^ str);
	System::String^ ConvertString(const std::string& str);
	std::wstring ConvertWString(System::String^ str);
	System::String^ ConvertWString(const std::wstring& str);
}
