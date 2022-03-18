#include <iostream>
#include "Tool.h"
#include "Workbook.h"
#include <boost/optional.hpp>
//#using "F:\ExcelOperator\ExcelOperator\bin\Debug\ExcelOperator.dll"
using namespace System;
/*
* @Type	  函数
* @Author ZhangZM
* @Date   2022/03/17
* @Brief  直接使用C#中的ExcelOperator对象进行接口测试
* @Other
* @Param
* @Return
*/
void TestForCSharp()
{
	std::string filePath = R"(F:\ExcelOperator\TestData\Data.xlsx)";//文件需要修改成本地路径
	String^ filePathCSharp = Tool::ConvertString(filePath);
	ExcelOperator::ExcelOperator^ excelOperator = gcnew ExcelOperator::ExcelOperator();
	excelOperator->HelloWorld();
	bool bRes = excelOperator->ReadExcel(filePathCSharp);
	if (!bRes) {
		std::string message = Tool::ConvertString(excelOperator->Message());
		std::cout << message;
	}
	else {
		//读取单元格内容测试
		String^ strResult;
		bRes = excelOperator->GetCellValue(0, 0, 0, strResult);
		if (bRes) {
			char message[64];
			sprintf_s(message, "读取字符串数据成功：%s", Tool::ConvertString(strResult).c_str());
			std::cout << message << std::endl;
		}
		else {
			std::cout << "读取字符串数据失败" << std::endl;
		}

		int intResult = 0;
		bRes = excelOperator->GetCellValue(0, 1, 0, intResult);
		if (bRes) {
			char message[64];
			sprintf_s(message, "读取整型数据数据成功：%d", intResult);
			std::cout << message << std::endl;
		}
		else {
			std::cout << "读取整型数据失败" << std::endl;
		}

		double doubleResult = 0;
		bRes = excelOperator->GetCellValue(0, 2, 0, doubleResult);
		if (bRes) {
			char message[64];
			sprintf_s(message, "读取浮点型数据成功：%f", doubleResult);
			std::cout << message << std::endl;
		}
		else {
			std::cout << "读取浮点型数据失败" << std::endl;
		}

		bool boolResult = 0;
		bRes = excelOperator->GetCellValue(0, 3, 0, boolResult);
		if (bRes) {
			char message[64];
			sprintf_s(message, "读取布尔型数据成功：%d", boolResult);
			std::cout << message << std::endl;
		}
		else {
			std::cout << "读取布尔型数据失败" << std::endl;
		}

		//写入单元格内容测试
		bRes = excelOperator->SetCellValue(0, 3, 3, "文本内容");
		if (bRes) {
			std::cout << "写入字符串数据成功" << std::endl;
		}
		else {
			std::cout << "写入字符串数据失败" << std::endl;
		}
		bRes = excelOperator->SetCellValue(0, 3, 4, 5);
		if (bRes) {
			std::cout << "写入整型数据成功" << std::endl;
		}
		else {
			std::cout << "写入整数数据失败" << std::endl;
		}
		bRes = excelOperator->SetCellValue(0, 3, 5, 5.5);
		if (bRes) {
			std::cout << "写入浮点型数据成功" << std::endl;
		}
		else {
			std::cout << "写入浮点型数据失败" << std::endl;
		}
		bRes = excelOperator->SetCellValue(0, 3, 6, true);
		if (bRes) {
			std::cout << "写入布尔型数据成功" << std::endl;
		}
		else {
			std::cout << "写入布尔型数据失败" << std::endl;
		}
		bRes = excelOperator->WriteExcel();
		if (bRes) {
			std::cout << "保存文件成功" << std::endl;
		}
		else {
			std::cout << (std::string)"保存文件失败，失败原因为：" + Tool::ConvertString(excelOperator->Message()).c_str() << std::endl;
		}
	}
}

/*
* @Type	  函数
* @Author ZhangZM
* @Date   2022/03/17
* @Brief  使用ExcelOperatorCpp项目封装好的对象进行接口测试
* @Other
* @Param
* @Return
*/
void TestForCpp()
{
	std::string filePath = R"(F:\ExcelOperator\TestData\Data.xlsx)";//文件需要修改成本地路径
	WorkbookImpl workbook;
	bool bRes = workbook.Parse(filePath.c_str());
	if (bRes) {
		int sheetCount = workbook.SheetCount();
		auto sheet = workbook.GetSheet(0);
		if (sheet) {
			int rowCount = sheet->RowCount();

			//读取单元格内容测试
			auto cell = sheet->GetCell(0, 0);
			if (cell) {
				const char* result = cell->GetValue();
				if (result != NULL) {
					char message[64];
					sprintf_s(message, "读取字符串数据成功：%s", result);
					std::cout << message << std::endl;
				}
				else {
					std::cout << "读取字符串数据失败" << std::endl;
				}
			}
			cell = sheet->GetCell(1, 0);
			if (cell) {
				int result;
				bRes = cell->GetValue(result);
				if (bRes) {
					char message[64];
					sprintf_s(message, "读取整型数据成功：%d", result);
					std::cout << message << std::endl;
				}
				else {
					std::cout << "读取整型数据失败" << std::endl;
				}
			}
			cell = sheet->GetCell(2, 0);
			if (cell) {
				double result;
				bRes = cell->GetValue(result);
				if (bRes) {
					char message[64];
					sprintf_s(message, "读取浮点型数据成功：%f", result);
					std::cout << message << std::endl;
				}
				else {
					std::cout << "读取浮点型数据失败" << std::endl;
				}
			}
			cell = sheet->GetCell(3, 0);
			if (cell) {
				bool result;
				bRes = cell->GetValue(result);
				if (bRes) {
					char message[64];
					sprintf_s(message, "读取布尔型数据成功：%d", result);
					std::cout << message << std::endl;
				}
				else
				{
					std::cout << "读取布尔型数据失败" << std::endl;
				}
			}

			//写入单元格内容测试
			cell = sheet->GetCell(3, 3);
			if (cell) {
				std::string value = "文本内容";
				bRes = cell->SetValue(value.c_str());
				if (bRes) {
					std::cout << "写入字符串数据成功" << std::endl;
				}
				else {
					std::cout << "写入字符串数据失败" << std::endl;
				}
			}
			cell = sheet->GetCell(3, 4);
			if (cell) {
				int value = 5;
				bRes = cell->SetValue(value);
				if (bRes) {
					std::cout << "写入整型数据成功" << std::endl;
				}
				else {
					std::cout << "写入整型数据失败" << std::endl;
				}
			}
			cell = sheet->GetCell(3, 5);
			if (cell) {
				double value = 5.5;
				bRes = cell->SetValue(value);
				if (bRes) {
					std::cout << "写入浮点型数据成功" << std::endl;
				}
				else {
					std::cout << "写入浮点型数据失败" << std::endl;
				}
			}
			cell = sheet->GetCell(3, 6);
			if (cell) {
				bool value = true;
				bRes = cell->SetValue(value);
				if (bRes) {
					std::cout << "写入布尔型数据成功" << std::endl;
				}
				else {
					std::cout << "写入布尔型数据失败" << std::endl;
				}
			}
			bRes = workbook.Save();
			if (bRes) {
				std::cout << "保存文件成功" << std::endl;
			}
			else {
				std::cout << (std::string)"保存文件失败，失败原因为：" + workbook.Message() << std::endl;
			}

		}
	}
	else {
		std::cout << (std::string)"解析文件失败，失败原因为：" + workbook.Message() << std::endl;
	}
}

void main()
{
	//调试入口，用于平常调试C#项目（ExcelOperator）生成的dll或者封装的C++项目（ExcelOperatorCpp)

	TestForCSharp();
	//TestForCpp();
}