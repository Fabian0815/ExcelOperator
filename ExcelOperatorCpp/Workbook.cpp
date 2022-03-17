// ExcelOperatorCpp.cpp : 定义静态库的函数。
//

#include "Workbook.h"
#include "Tool.h"
#include <iostream>
WorkbookImpl::WorkbookImpl()
{
	m_excelOperator = nullptr;
}

WorkbookImpl::~WorkbookImpl()
{
	if (m_excelOperator)
	{
		delete m_excelOperator;
	}
}

bool WorkbookImpl::Parse(const char* filePath)
{
	try
	{
		String^ filePathClr = Tool::ConvertString(filePath);
		ExcelOperator::ExcelOperator^ excelOperator = gcnew ExcelOperator::ExcelOperator();
		bool bRes = excelOperator->ReadExcel(filePathClr);
		if (bRes)
		{
			m_excelOperator = new gcroot<ExcelOperator::ExcelOperator^>();
			*m_excelOperator = excelOperator;
		}
		else
		{
			m_message = Tool::ConvertString(excelOperator->Message());
		}
		return bRes;
	}
	catch (std::exception ex)
	{
		if (m_excelOperator)
		{
			String^ messageClr = Tool::ConvertString(ex.what());
			(*m_excelOperator)->SetMessage(messageClr);
		}
		else
		{
			m_message = ex.what();
		}
		return false;
	}
}

bool WorkbookImpl::Parse(const wchar_t* filePath)
{
	try
	{
		String^ filePathClr = Tool::ConvertWString(filePath);
		ExcelOperator::ExcelOperator^ excelOperator = gcnew ExcelOperator::ExcelOperator();
		bool bRes = excelOperator->ReadExcel(filePathClr);
		if (bRes)
		{
			m_excelOperator = new gcroot<ExcelOperator::ExcelOperator^>();
			*m_excelOperator = excelOperator;
		}
		else
		{
			if (m_excelOperator)
			{
				delete m_excelOperator;
			}
			m_excelOperator = nullptr;
			m_message = Tool::ConvertString(excelOperator->Message());
		}
		return bRes;
	}
	catch (std::exception ex)
	{
		if (m_excelOperator)
		{
			String^ messageClr = Tool::ConvertString(ex.what());
			(*m_excelOperator)->SetMessage(messageClr);
		}
		else
		{
			m_message = ex.what();
		}
		return false;
	}
}

Export::Sheet* WorkbookImpl::GetSheet(const int& index)
{
	if (m_excelOperator && index >= 0 && index <= SheetCount())
	{
		auto sheet = new SheetImpl(index, m_excelOperator);
		return sheet;
	}
	else
	{
		return nullptr;
	}
}

int WorkbookImpl::SheetCount() const
{
	if (m_excelOperator)
	{
		return (*m_excelOperator)->SheetCount();
	}
	return -1;
}

bool WorkbookImpl::Save()
{
	try
	{
		if (m_excelOperator)
		{
			bool bRes = (*m_excelOperator)->WriteExcel();
			return bRes;
		}
		return false;
	}
	catch (std::exception ex)
	{
		if (m_excelOperator)
		{
			String^ messageClr = Tool::ConvertString(ex.what());
			(*m_excelOperator)->SetMessage(messageClr);
		}
		else
		{
			m_message = ex.what();
		}
		return false;
	}
}

const char* WorkbookImpl::Message()
{
	if (m_excelOperator)
	{
		String^ messageClr = (*m_excelOperator)->Message();
		return Tool::ConvertString(messageClr).c_str();
	}
	else
	{
		return m_message.c_str();
	}
}

const wchar_t* WorkbookImpl::MessageW()
{
	if (m_excelOperator)
	{
		String^ messageClr = (*m_excelOperator)->Message();
		return Tool::ConvertWString(messageClr).c_str();
	}
	else
	{
		m_messagew = Tool::StringToWstring(m_message);
		return m_messagew.c_str();
	}
}

Export::CellStyle* WorkbookImpl::CreateCellStyle()
{
	try
	{
		if (m_excelOperator)
		{
			short cellStyleIndex = -1;
			bool bRes = (*m_excelOperator)->CreateCellStyle(cellStyleIndex);
			if (bRes)
			{
				auto cellStyle = new CellStyleImpl(cellStyleIndex, m_excelOperator);
				return cellStyle;
			}
		}
		return nullptr;
	}
	catch (std::exception ex)
	{
		if (m_excelOperator)
		{
			String^ messageClr = Tool::ConvertString(ex.what());
			(*m_excelOperator)->SetMessage(messageClr);
		}
		else
		{
			m_message = ex.what();
		}
		return nullptr;
	}
}

CellImpl::CellImpl(const int& sheetIndex, const int& row, const int& column, gcroot<ExcelOperator::ExcelOperator^>* excelOperator)
{
	m_sheetIndex = sheetIndex;
	m_row = row;
	m_column = column;
	m_excelOperator = excelOperator;
}

int CellImpl::Row() const
{
	return m_row;
}

int CellImpl::Column() const
{
	return m_column;
}

Export::CellType CellImpl::GetCellType() const
{
	if (m_excelOperator)
	{
		int result = -1;
		if ((*m_excelOperator)->GetCellType(m_sheetIndex, m_row, m_column, result))
		{
			return Export::CellType(result);
		}
		else
		{
			return Export::CellType::Error;
		}

	}
	else
	{
		return Export::CellType::Error;
	}
}

Export::CellStyle* CellImpl::GetCellStyle() const
{
	if (m_excelOperator)
	{
		short styleIndex = -1;
		if ((*m_excelOperator)->GetCellStyle(m_sheetIndex, m_row, m_column, styleIndex))
		{
			return  new CellStyleImpl(styleIndex, m_excelOperator);
		}
		else
		{
			return nullptr;
		}

	}
	else
	{
		return nullptr;
	}
}

bool CellImpl::SetValue(const int& value)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellValue(m_sheetIndex, m_row, m_column, (int)value);
		return bRes;
	}
	return false;
}

bool CellImpl::SetValue(const double& value)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellValue(m_sheetIndex, m_row, m_column, value);
		return bRes;
	}
	return false;
}

bool CellImpl::SetValue(const bool& value)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellValue(m_sheetIndex, m_row, m_column, value);
		return bRes;
	}
	return false;
}

bool CellImpl::SetValue(const char* value)
{
	if (m_excelOperator)
	{
		String^ valueClr = Tool::ConvertString(value);
		bool bRes = (*m_excelOperator)->SetCellValue(m_sheetIndex, m_row, m_column, valueClr);
		return bRes;
	}
	return false;
}

bool CellImpl::SetValue(const wchar_t* value)
{
	if (m_excelOperator)
	{
		String^ valueClr = Tool::ConvertWString(value);
		bool bRes = (*m_excelOperator)->SetCellValue(m_sheetIndex, m_row, m_column, valueClr);
		return bRes;
	}
	return false;
}

bool CellImpl::SetValue(const char* value, Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment)
{
	if (m_excelOperator)
	{
		String^ valueClr = Tool::ConvertString(value);
		bool bRes = (*m_excelOperator)->SetCellValue(m_sheetIndex, m_row, m_column, valueClr, horizontalAlignment, verticalAlignment);
		return bRes;
	}
	return false;
}

bool CellImpl::SetValue(const wchar_t* value, Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment)
{
	if (m_excelOperator)
	{
		String^ valueClr = Tool::ConvertWString(value);
		bool bRes = (*m_excelOperator)->SetCellValue(m_sheetIndex, m_row, m_column, valueClr, horizontalAlignment, verticalAlignment);
		return bRes;
	}
	return false;
}

bool CellImpl::SetValue(const int& value, Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellValue(m_sheetIndex, m_row, m_column, value, horizontalAlignment, verticalAlignment);
		return bRes;
	}
	return false;
}

bool CellImpl::SetValue(const double& value, Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellValue(m_sheetIndex, m_row, m_column, value, horizontalAlignment, verticalAlignment);
		return bRes;
	}
	return false;
}

bool CellImpl::SetValue(const bool& value, Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellValue(m_sheetIndex, m_row, m_column, value, horizontalAlignment, verticalAlignment);
		return bRes;
	}
	return false;
}

bool CellImpl::SetBorderStyle(Export::BorderStyle borderStyle)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellBorderStyle(m_sheetIndex, m_row, m_column, borderStyle);
		return bRes;
	}
	return false;
}

bool CellImpl::SetBorderColor(Export::ColorIndex colorIndex)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellBorderColor(m_sheetIndex, m_row, m_column, colorIndex);
		return bRes;
	}
	return false;
}

bool CellImpl::SetBackgroundColor(Export::ColorIndex colorIndex)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellBackgroundColor(m_sheetIndex, m_row, m_column, colorIndex);
		return bRes;
	}
	return false;
}

bool CellImpl::SetFontColor(Export::ColorIndex colorIndex)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellFontColor(m_sheetIndex, m_row, m_column, colorIndex);
		return bRes;
	}
	return false;
}

bool CellImpl::SetFontSize(double size)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellFontSize(m_sheetIndex, m_row, m_column, size);
		return bRes;
	}
	return false;
}

bool CellImpl::SetHorizontalAlignment(Export::HorizontalAlignment align)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellHorizontalAlign(m_sheetIndex, m_row, m_column, align);
		return bRes;
	}
	return false;
}

bool CellImpl::SetVerticalAlignment(Export::VerticalAlignment align)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellVerticalAlign(m_sheetIndex, m_row, m_column, align);
		return bRes;
	}
	return false;
}

bool CellImpl::SetAlignment(Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellAlign(m_sheetIndex, m_row, m_column, horizontalAlignment, verticalAlignment);
		return bRes;
	}
	return false;
}

bool CellImpl::SetCellStyle(const short& styleIndex)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellStyle(m_sheetIndex, m_row, m_column, styleIndex);
		return bRes;
	}
	return false;
}

bool CellImpl::GetValue(int& result)const
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->GetCellValue(m_sheetIndex, m_row, m_column, result);
		return bRes;
	}
	return false;
}

bool CellImpl::GetValue(double& result)const
{
	if (m_excelOperator != nullptr)
	{
		bool bRes = (*m_excelOperator)->GetCellValue(m_sheetIndex, m_row, m_column, result);
		return bRes;
	}
	return false;
}

bool CellImpl::GetValue(bool& result)const
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->GetCellValue(m_sheetIndex, m_row, m_column, result);
		return bRes;
	}
	return false;
}

const char* CellImpl::GetValue()
{
	if (m_excelOperator)
	{
		String^ resultClr;
		bool bRes = (*m_excelOperator)->GetCellValue(m_sheetIndex, m_row, m_column, resultClr);
		if (bRes)
		{
			m_content = Tool::ConvertString(resultClr);
			return m_content.c_str();
		}
	}
	return NULL;
}

const wchar_t* CellImpl::GetValueW()
{
	if (m_excelOperator)
	{
		String^ resultClr;
		bool bRes = (*m_excelOperator)->GetCellValue(m_sheetIndex, m_row, m_column, resultClr);
		if (bRes)
		{
			m_contentw = Tool::ConvertWString(resultClr);
			return m_contentw.c_str();
		}
	}
	return NULL;
}

SheetImpl::SheetImpl(const int& index, gcroot<ExcelOperator::ExcelOperator^>* excelOperator)
{
	m_index = index;
	m_excelOperator = excelOperator;
}

Export::Cell* SheetImpl::GetCell(const int& row, const int& column)
{
	if (m_excelOperator && row >= 0 && column >= 0)
	{
		auto cell = new CellImpl(m_index, row, column, m_excelOperator);
		return cell;
	}
	else
	{
		return nullptr;
	}
}

int SheetImpl::Index() const
{
	return m_index;
}

int SheetImpl::RowCount() const
{
	return (*m_excelOperator)->RowCount(m_index);
}

int SheetImpl::ColumnCount() const
{
	return (*m_excelOperator)->ColumnCount(m_index);
}

CellStyleImpl::CellStyleImpl(const short& styleIndex, gcroot<ExcelOperator::ExcelOperator^>* excelOperator)
{
	m_styleIndex = styleIndex;
	m_excelOperator = excelOperator;
}

short CellStyleImpl::Index() const
{
	return m_styleIndex;
}

bool CellStyleImpl::SetBorderStyle(Export::BorderStyle borderStyle)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellStyleBorderStyle(m_styleIndex, borderStyle);
		return bRes;
	}
	return false;
}

bool CellStyleImpl::SetBorderColor(Export::ColorIndex colorIndex)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellStyleBorderColor(m_styleIndex, colorIndex);
		return bRes;
	}
	return false;
}

bool CellStyleImpl::SetBackgroundColor(Export::ColorIndex colorIndex)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellStyleBackgroundColor(m_styleIndex, colorIndex);
		return bRes;
	}
	return false;
}

bool CellStyleImpl::SetFontColor(Export::ColorIndex colorIndex)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellStyleFontColor(m_styleIndex, colorIndex);
		return bRes;
	}
	return false;
}

bool CellStyleImpl::SetFontSize(double size)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellStyleFontSize(m_styleIndex, size);
		return bRes;
	}
	return false;
}

bool CellStyleImpl::SetHorizontalAlignment(Export::HorizontalAlignment align)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellStyleHorizontalAlign(m_styleIndex, align);
		return bRes;
	}
	return false;
}

bool CellStyleImpl::SetVerticalAlignment(Export::VerticalAlignment align)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellStyleVerticalAlign(m_styleIndex, align);
		return bRes;
	}
	return false;
}

bool CellStyleImpl::SetAlignment(Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment)
{
	if (m_excelOperator)
	{
		bool bRes = (*m_excelOperator)->SetCellStyleAlign(m_styleIndex, horizontalAlignment, verticalAlignment);
		return bRes;
	}
	return false;
}
