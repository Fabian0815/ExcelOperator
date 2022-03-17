#pragma once
#include "Export.h"
#include <boost/optional.hpp>
using boost::intrusive_ptr;

/*
* @Type   类
* @Author ZhangZM
* @Date   2022/03/17
* @Brief  单元格样式
* @Other
*/
class CellStyle
{
public:
	CellStyle(Export::CellStyle* ptr) :m_internal(ptr) {}

	//获取接口
	short Index()const
	{
		return m_internal->Index();
	}

	//设置接口
	bool SetBorderStyle(Export::BorderStyle borderStyle)
	{
		return m_internal->SetBorderStyle(borderStyle);
	}
	bool SetBorderColor(Export::ColorIndex colorIndex)
	{
		return m_internal->SetBorderColor(colorIndex);
	}
	bool SetBackgroundColor(Export::ColorIndex colorIndex)
	{
		return m_internal->SetBackgroundColor(colorIndex);
	}
	bool SetFontColor(Export::ColorIndex colorIndex)
	{
		return m_internal->SetFontColor(colorIndex);
	}
	bool SetFontSize(double size)
	{
		return m_internal->SetFontSize(size);
	}
	bool SetHorizontalAlignment(Export::HorizontalAlignment align)
	{
		return m_internal->SetHorizontalAlignment(align);
	}
	bool SetVerticalAlignment(Export::VerticalAlignment align)
	{
		return m_internal->SetVerticalAlignment(align);
	}
	bool SetAlignment(Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment)
	{
		return m_internal->SetAlignment(horizontalAlignment, verticalAlignment);
	}
private:
	intrusive_ptr<Export::CellStyle> m_internal;
};

/*
* @Type   类
* @Author ZhangZM
* @Date   2022/03/17
* @Brief  单元格对象
* @Other
*/
class Cell
{
public:
	Cell(Export::Cell* ptr) :m_internal(ptr) {}

	/**
	* @brief 获取单元格所在行数
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	int Row()const
	{
		return m_internal->Row();
	}

	/**
	* @brief 获取单元格所在列数
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	int Column()const
	{
		return m_internal->Column();
	}

	/**
	* @brief 获取单元格内容类型
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	Export::CellType GetCellType()const
	{
		return m_internal->GetCellType();
	}

	/**
	* @brief 获取单元格样式
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	Export::CellStyle* GetCellStyle()const
	{
		Export::CellStyle* cellStyle = m_internal->GetCellStyle();
		return cellStyle;
	}

	/**
	* @brief 获取单元格内容
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	const char* GetValue()
	{
		return m_internal->GetValue();
	}
	const wchar_t* GetValueW()
	{
		return m_internal->GetValueW();
	}
	bool GetValue(int& result)const
	{
		return m_internal->GetValue(result);
	}
	bool GetValue(double& result)const
	{
		return m_internal->GetValue(result);
	}
	bool GetValue(bool& result)const
	{
		return m_internal->GetValue(result);
	}

	/**
	* @brief 设置单元格内容
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	bool SetValue(const char* value)
	{
		return m_internal->SetValue(value);
	}
	bool SetValue(const wchar_t* value)
	{
		return m_internal->SetValue(value);
	}
	bool SetValue(const int& value)
	{
		return m_internal->SetValue(value);
	}
	bool SetValue(const double& value)
	{
		return m_internal->SetValue(value);
	}
	bool SetValue(const bool& value)
	{
		return m_internal->SetValue(value);
	}

	/**
	* @brief 设置单元格内容，同时设置对齐方式
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	bool SetValue(const char* value, Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment)
	{
		return m_internal->SetValue(value, horizontalAlignment, verticalAlignment);
	}
	bool SetValue(const wchar_t* value, Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment)
	{
		return m_internal->SetValue(value, horizontalAlignment, verticalAlignment);
	}
	bool SetValue(const int& value, Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment)
	{
		return m_internal->SetValue(value, horizontalAlignment, verticalAlignment);
	}
	bool SetValue(const double& value, Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment)
	{
		return m_internal->SetValue(value, horizontalAlignment, verticalAlignment);
	}
	bool SetValue(const bool& value, Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment)
	{
		return m_internal->SetValue(value, horizontalAlignment, verticalAlignment);
	}

	/**
	* @brief 设置边框样式
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	bool SetBorderStyle(Export::BorderStyle borderStyle)
	{
		return m_internal->SetBorderStyle(borderStyle);
	}

	/**
	* @brief  设置边框颜色
	* @param
	* @return
	* @other  设置边框颜色前需要先设置边框样式
	* @author ZhangZM
	*/
	bool SetBorderColor(Export::ColorIndex colorIndex)
	{
		return m_internal->SetBorderColor(colorIndex);
	}

	/**
	* @brief 设置背景颜色
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	bool SetBackgroundColor(Export::ColorIndex colorIndex)
	{
		return m_internal->SetBackgroundColor(colorIndex);
	}

	/**
	* @brief 设置字体颜色
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	bool SetFontColor(Export::ColorIndex colorIndex)
	{
		return m_internal->SetFontColor(colorIndex);
	}

	/**
	* @brief 设置字体尺寸
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	bool SetFontSize(double size)
	{
		return m_internal->SetFontSize(size);
	}

	/**
	* @brief 设置水平对齐方式
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	bool SetHorizontalAlignment(Export::HorizontalAlignment align)
	{
		return m_internal->SetHorizontalAlignment(align);
	}

	/**
	* @brief 设置垂直对齐方式
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	bool SetVerticalAlignment(Export::VerticalAlignment align)
	{
		return m_internal->SetVerticalAlignment(align);
	}

	/**
	* @brief 设置水平与垂直对齐方式
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	bool SetAlignment(Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment)
	{
		return m_internal->SetAlignment(horizontalAlignment, verticalAlignment);
	}

	/**
	* @brief 设置单元格样式
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	bool SetCellStyle(const short& styleIndex)const
	{
		return m_internal->SetCellStyle(styleIndex);
	}
	bool SetCellStyle(CellStyle cellStyle)const
	{
		return m_internal->SetCellStyle(cellStyle.Index());
	}
private:
	intrusive_ptr<Export::Cell> m_internal;
};

class Sheet
{
public:
	Sheet(Export::Sheet* ptr) :m_internal(ptr) {}

	/**
	* @brief 获取工作表序号
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	int Index() const
	{
		return m_internal->Index();
	}

	/**
	* @brief 获取工作表当前的最大行数
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	int RowCount() const
	{
		return m_internal->RowCount();
	}

	/**
	* @brief 获取工作表当前的最大列数
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	int ColumnCount()const
	{
		return m_internal->ColumnCount();
	}

	/**
	* @brief 获取Sheet中某一个单元格
	* @param
	* @return
	* @other 单元格不存在时会主动创建
	* @author ZhangZM
	*/
	boost::optional<Cell> GetCell(const int& row, const int& column)
	{
		Export::Cell* ptr = m_internal->GetCell(row, column);
		if (ptr == NULL)
		{
			return boost::none;
		}
		else
		{
			return Cell(ptr);
		}
	}

private:
	intrusive_ptr<Export::Sheet> m_internal;
};

class Workbook
{
public:
	Workbook() :m_internal(CreateWorkbook()) {}

	/**
	* @brief 读取Excel文件
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	bool Parse(const char* filePath)
	{
		return m_internal->Parse(filePath);
	}
	bool Parse(const wchar_t* filePath)
	{
		return m_internal->Parse(filePath);
	}

	/**
	* @brief 获取某一个工作表
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	boost::optional<Sheet> GetSheet(const int& index)
	{
		Export::Sheet* ptr = m_internal->GetSheet(index);
		if (ptr == NULL)
		{
			return boost::none;
		}
		else
		{
			return Sheet(ptr);
		}
	}

	/**
	* @brief 获取工作表数量
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	int SheetCount()const
	{
		return m_internal->SheetCount();
	}

	/**
	* @brief 保存Excel文件
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	bool Save()
	{
		return m_internal->Save();
	}

	/**
	* @brief 获取上一失败操作的异常信息
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	const char* Message()
	{
		return m_internal->Message();
	}

	/**
	* @brief 创建单元格样式
	* @param
	* @return
	* @other
	* @author ZhangZM
	*/
	boost::optional<CellStyle> CreateCellStyle()
	{
		Export::CellStyle* ptr = m_internal->CreateCellStyle();
		if (ptr == NULL)
		{
			return boost::none;
		}
		else
		{
			return CellStyle(ptr);
		}
	}

private:
	intrusive_ptr<Export::Workbook> m_internal;
};
