#pragma once
#include "Build.h"
#include <boost/intrusive_ptr.hpp>

namespace Export {
	class DllResource;
}
extern "C" EXCELOPERATOR_API void release_resource(Export::DllResource * resource);
extern "C" EXCELOPERATOR_API void release_resource_const(const Export::DllResource * resource);

namespace Export {
	enum BorderStyle
	{
		None = 0,
		Thin = 1,
		Medium = 2,
		Dashed = 3,
		Dotted = 4,
		Thick = 5,
		Double = 6,
		Hair = 7,
		MediumDashed = 8,
		DashDot = 9,
		MediumDashDot = 10,
		DashDotDot = 11,
		MediumDashDotDot = 12,
		SlantedDashDot = 13
	};
	enum CellType
	{
		Unknown = -1,
		Numeric = 0,
		String = 1,
		Formula = 2,
		Blank = 3,
		Boolean = 4,
		Error = 5
	};
	enum ColorIndex
	{
		Black = 8,
		White = 9,
		Red = 10,
		Blue = 12,
		Yellow = 13,
		Pink = 14,//粉色
		Green = 17,
		Orange = 53,
		//后续需要用到其他颜色再参照NPOI颜色对照表增加
	};
	enum HorizontalAlignment
	{
		HGeneral = 0,
		HLeft = 1,
		HCenter = 2,
		HRight = 3,
		HFill = 4,
		HJustify = 5,
		HCenterSelection = 6,
		HDistributed = 7
	};
	enum VerticalAlignment
	{
		VNone = -1,
		VTop = 0,
		VCenter = 1,
		VBottom = 2,
		VJustify = 3,
		VDistributed = 4
	};

	class DllResource
	{
	public:
		DllResource() :m_refCount(0) {}
		virtual ~DllResource() {}
	private:
		friend void intrusive_ptr_add_ref(DllResource* p);
		friend void intrusive_ptr_release(DllResource* p);
		friend void intrusive_ptr_add_ref(const DllResource* p);
		friend void intrusive_ptr_release(const DllResource* p);

	private:
		mutable size_t m_refCount;
	};

	inline void intrusive_ptr_add_ref(DllResource* p) {
		++p->m_refCount;
	}

	inline void intrusive_ptr_release(DllResource* p) {
		if (--p->m_refCount == 0) {
			release_resource(p);
		}
	}

	inline void intrusive_ptr_add_ref(const DllResource* p) {
		++p->m_refCount;//mutable关键字起作用
	}

	inline void intrusive_ptr_release(const DllResource* p) {
		if (--p->m_refCount == 0) {
			release_resource_const(p);
		}
	}

	class CellStyle :public DllResource
	{
	public:
		//获取接口
		virtual short Index()const = 0;

		//设置接口
		virtual bool SetBorderStyle(BorderStyle borderStyle) = 0;
		virtual bool SetBorderColor(ColorIndex colorIndex) = 0;
		virtual bool SetBackgroundColor(ColorIndex colorIndex) = 0;
		virtual bool SetFontColor(ColorIndex colorIndex) = 0;
		virtual bool SetFontSize(double size) = 0;
		virtual bool SetHorizontalAlignment(HorizontalAlignment align) = 0;
		virtual bool SetVerticalAlignment(VerticalAlignment align) = 0;
		virtual bool SetAlignment(Export::HorizontalAlignment horizontalAlignment, Export::VerticalAlignment verticalAlignment) = 0;
	};

	class Cell :public DllResource
	{
	public:
		//获取其他属性接口
		virtual int Row()const = 0;
		virtual int Column()const = 0;
		virtual CellType GetCellType()const = 0;
		virtual CellStyle* GetCellStyle()const = 0;

		//获取值接口
		virtual const char* GetValue() = 0;
		virtual const wchar_t* GetValueW() = 0;
		virtual bool GetValue(int& result)const = 0;
		virtual bool GetValue(double& result)const = 0;
		virtual bool GetValue(bool& result)const = 0;

		//设置值接口
		virtual bool SetValue(const char* result) = 0;
		virtual bool SetValue(const wchar_t* result) = 0;
		virtual bool SetValue(const int& result) = 0;
		virtual bool SetValue(const double& result) = 0;
		virtual bool SetValue(const bool& result) = 0;
		virtual bool SetValue(const char* result, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment) = 0;
		virtual bool SetValue(const wchar_t* result, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment) = 0;
		virtual bool SetValue(const int& result, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment) = 0;
		virtual bool SetValue(const double& result, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment) = 0;
		virtual bool SetValue(const bool& result, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment) = 0;

		//设置样式接口
		virtual bool SetBorderStyle(BorderStyle borderStyle) = 0;
		virtual bool SetBorderColor(ColorIndex colorIndex) = 0;
		virtual bool SetBackgroundColor(ColorIndex colorIndex) = 0;
		virtual bool SetFontColor(ColorIndex colorIndex) = 0;
		virtual bool SetFontSize(double size) = 0;
		virtual bool SetHorizontalAlignment(HorizontalAlignment align) = 0;
		virtual bool SetVerticalAlignment(VerticalAlignment align) = 0;
		virtual bool SetAlignment(HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment) = 0;
		virtual bool SetCellStyle(const short& styleIndex) = 0;
	};
	class Sheet :public DllResource
	{
	public:
		virtual int Index()const = 0;
		virtual int RowCount()const = 0;
		virtual int ColumnCount()const = 0;
		virtual Cell* GetCell(const int& row, const int& column) = 0;
	};
	class Workbook :public DllResource
	{
	public:
		virtual ~Workbook() { }

		//解析获取接口
		virtual bool Parse(const char* filePath) = 0;
		virtual bool Parse(const wchar_t* filePath) = 0;
		virtual Sheet* GetSheet(const int& index) = 0;
		virtual int SheetCount()const = 0;

		//保存写入接口
		virtual bool Save() = 0;
		virtual const char* Message() = 0;
		virtual const wchar_t* MessageW() = 0;

		//创建接口
		virtual CellStyle* CreateCellStyle() = 0;
	};
}

extern "C" EXCELOPERATOR_API Export::Workbook * CreateWorkbook();
