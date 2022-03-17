#pragma once
#include <string>
#include <memory>
#include <vcclr.h>
#include "Export.h"
class CellStyleImpl :public Export::CellStyle
{
public:
	CellStyleImpl(const short& styleIndex, gcroot<ExcelOperator::ExcelOperator^>* excelOperator);

	//获取接口
	short Index()const override;

	//设置接口
	bool SetBorderStyle(Export::BorderStyle borderStyle)override;
	bool SetBorderColor(Export::ColorIndex colorIndex) override;
	bool SetBackgroundColor(Export::ColorIndex colorIndex) override;
	bool SetFontColor(Export::ColorIndex colorIndex) override;
	bool SetFontSize(double size) override;
	bool SetHorizontalAlignment(Export::HorizontalAlignment align)override;
	bool SetVerticalAlignment(Export::VerticalAlignment align)override;
	bool SetAlignment(Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment)override;

private:
	short m_styleIndex = -1;
	gcroot<ExcelOperator::ExcelOperator^>* m_excelOperator;
};
class CellImpl :public Export::Cell
{
public:
	CellImpl(const int& sheetIndex, const int& row, const int& column, gcroot<ExcelOperator::ExcelOperator^>* excelOperator);

	//获取其他属性接口
	int Row()const override;
	int Column()const override;
	Export::CellType GetCellType()const override;
	Export::CellStyle* GetCellStyle()const override;

	//获取值接口
	//bool GetValue(std::string& result)const;
	const char* GetValue() override;
	const wchar_t* GetValueW() override;
	bool GetValue(int& result)const override;
	bool GetValue(double& result)const override;
	bool GetValue(bool& result)const override;

	//设置值接口
	//bool SetValue(const std::string& value);
	bool SetValue(const char* value)override;
	bool SetValue(const wchar_t* value)override;
	bool SetValue(const int& value) override;
	bool SetValue(const double& value) override;
	bool SetValue(const bool& value)override;
	bool SetValue(const char* value, Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment);
	bool SetValue(const wchar_t* value, Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment);
	bool SetValue(const int& value, Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment);
	bool SetValue(const double& value, Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment);
	bool SetValue(const bool& value, Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment);

	//设置样式接口
	bool SetBorderStyle(Export::BorderStyle borderStyle)override;
	bool SetBorderColor(Export::ColorIndex colorIndex) override;
	bool SetBackgroundColor(Export::ColorIndex colorIndex) override;
	bool SetFontColor(Export::ColorIndex colorIndex) override;
	bool SetFontSize(double size) override;
	bool SetHorizontalAlignment(Export::HorizontalAlignment align)override;
	bool SetVerticalAlignment(Export::VerticalAlignment align)override;
	bool SetAlignment(Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment)override;

	bool SetCellStyle(const short& styleIndex)override;
private:
	int m_sheetIndex = -1;
	int m_row = -1;
	int m_column = -1;
	std::string m_content;//保存取到的值
	std::wstring m_contentw;//保存取到的值
	gcroot<ExcelOperator::ExcelOperator^>* m_excelOperator;
};
class SheetImpl :public Export::Sheet
{
public:
	SheetImpl(const int& index, gcroot<ExcelOperator::ExcelOperator^>* excelOperator);
	Export::Cell* GetCell(const int& row, const int& column)override;
	int Index()const;
	int RowCount()const;
	int ColumnCount()const;
private:
	int m_index = -1;
	gcroot<ExcelOperator::ExcelOperator^>* m_excelOperator;
};
class WorkbookImpl :public Export::Workbook
{
public:
	WorkbookImpl();
	~WorkbookImpl();
	//解析获取接口
	bool Parse(const char* filePath) override;
	bool Parse(const wchar_t* filePath) override;
	Export::Sheet* GetSheet(const int& index)override;
	int SheetCount()const override;

	//保存写入接口
	bool Save() override;
	const char* Message() override;
	const wchar_t* MessageW() override;

	//创建接口
	Export::CellStyle* CreateCellStyle() override;

private:
	std::string m_message;//用于ExcelOperator对象构建失败时错误信息存储，其他情况都是通过ExcelOperator.Message()获取
	std::wstring m_messagew;//wstring版本
	gcroot<ExcelOperator::ExcelOperator^>* m_excelOperator;
};
