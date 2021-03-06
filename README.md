# ExcelOperator

**操作 Excel 文件的 C++ 库**：使用 CLR 技术在 C++ 项目中调用 C# 项目封装好的操作接口，C#项目使用 [NPOI](https://github.com/nissl-lab/npoi) 库操作 Excel 文件，同时封装 C++ 接口为纯 C 接口，以便在低版本项目中使用。**项目会自动释放 dll 内部通过 new 创建的内存，确保了内存安全。**

## 项目介绍

### 1.ExcelOperator

C# 项目，使用 [NPOI](https://github.com/nissl-lab/npoi) 库封装操作 Excel 文件的接口。

- ExcelOperator.cs：项目的唯一代码文件，定义操作 Excel 文件的接口。

### 2.ExcelOperatorCpp

C++ 项目，使用 CLR 调用 **ExcelOperator** 封装的接口，在此基础上封装成 C++ 接口，同时为了给低版本项目使用，将 C++ 接口封装成纯 C 接口，这是 **ExcelOpeatorCpp** 项目最终生成的接口形式。

- Build.h：定义导入导出宏。
- Tool.h/cpp：定义一些工具函数，如实现 System::String^ 与 stirng 的互转函数。
- Export.h/cpp：定义虚基类，这些类型在 Workbook.h 中被继承实现。
- Workbook.h/cpp：使用 CLR 调用 C# 项目接口，封装成 C++ 接口。
- WorkbookWrapper.h：Workbook 的包装类，将 Workbook 中的接口封装成纯 C 接口，提供给低版本项目使用。
- main.cpp：定义相关的测试代码。

## 项目配置

两个项目的文件夹下面已经包含了包配置文件，同时解决方案的 x64 生成配置已经配置完成，所以只需要依次进行项目生成就可以了，最终会在 **ExcelOperator\bin\Debug** 或者 **ExcelOperator\bin\Release** 路径下生成 dll 文件。

**注意：** ExcelOperatorCpp 项目依赖 ExcelOperator 项目生成的 dll 文件，所以需要先进行 ExcelOperator 项目的生成工作。

### 1.ExcelOperator

- 项目需要引用 NPOI，如果提示包缺失的话，按照 **右键项目-管理 NuGet 程序包-浏览** 操作，在弹出的页面中搜索 NPOI 进行安装。
- 项目默认配置中 .NET 运行时是 .NET Framework 4.8 ，如果本机没有安装该版本，则需要自行安装或者换成其他版本进行生成。

### 2.ExcelOperatorCpp

- 项目需要引用 boost 库，如果提示包缺失的话，按照 **右键项目-管理 NuGet 程序包-浏览** 操作，在弹出的页面中搜索 boost 进行安装。
- 如果想调试项目，可以修改项目配置为应用程序(.exe)，然后运行项目，没有错误的话，在控制台上会打出 “读取字符串数据成功....写入字符串数据成功” 等内容，测试代码在项目下的 main.cpp 文件中。

## 实际使用

### 1.配置工作

两个项目都运行成功后，根据运行配置不同，debug 版本文件在 **ExcelOperator\bin\Debug** 路径下生成，release 版本文件在 **ExcelOperator\bin\Release** 路径下生成。需要将生成的文件拷贝到你所开发的应用程序的生成目录下，接着将 ExcelOperatorCpp 项目下的 Build.h、Export.h、WorkbookWrap.h 文件拷贝到你所开发的项目中去，最后在你所开发的项目配置中引入 ExcelOperatorCpp.lib，此时，你就可以开始在 C++项目中自由地操作 Excel 文件了，也可以随时根据项目需求进行功能扩展。

(1)需要拷贝的库文件:

- BouncyCastle.Crypto.dll
- ICSharpCode.SharpZipLib.dll
- NPOI.dll
- NPOI.OOXML.dll
- NPOI.OpenXml4Net.dll
- NPOI.OpenXmlFormats.dll
- ExcelOperator.dll
- ExcelOperatorCpp.dll
- ExcelOperatorCpp.lib

(2)需要拷贝的头文件：

- Build.h
- Export.h
- WorkbookWrap.h

### 2.已实现的接口

```cpp
//Cell对象
int Row() const;    //获取行
int Column() const; //获取列
const char* GetValue(); //获取字符串类型值
bool SetValue(const char* value);   //设置字符串类型值
bool GetValue(int& result) const;   //获取整型值
bool SetValue(const int& value);    //设置整型值
bool GetValue(double& result) const;    //获取浮点型值
bool SetValue(const double& value); //设置浮点型值
bool GetValue(bool& result) const;  //获取布尔型值
bool SetValue(const bool& value);   //设置布尔型值
bool SetBorderStyle(Export::BorderStyle borderStyle);   //设置边框样式
bool SetFontSize(double size);  //设置字体大小
//...

//Sheet对象
int RowCount() const;   //获取工作表最大行数
int ColumnCount() const;    //获取工作表最大列数
boost::optional<Cell> GetCell(const int& row, const int& column);   //获取单元格
//...

//Workbook对象
bool Parse(const char* filePath);   //解析读取Excel文件
boost::optional<Sheet> GetSheet(const int& index);  //获取工作表
bool Save();    //保存Excel文件
//...

```

### 3.使用例子

```cpp
Workbook workbook;
if (!workbook.Parse(filePath.c_str()))
{
    string error = workbook.Message();
    return error;
}
boost::optional<Sheet> sheet = workbook.GetSheet(0);
if (sheet != boost::none)
{
    boost::optional<YFCell> cell = boost::none;
    cell = sheet->GetCell(0, 0);
    if (cell)
    {
        std::wstring str = cell->GetValueW();
    }
}
```

## 最后

作为 C++ 使用 CLR 技术调用 C# 库的实际例子，此项目除了作为自己后续在做相似功能需求时的参考资料，最重要的是给那些没有找到操作 Excel 文件的现成 C++ 库的程序员提供一种解决方案、一个现成库。自己当时也找了一段时间，找到的库要么是要求电脑上必须装有 Microsoft Office ，要么是操作不了 WPS Office 创建的 Excel 文件，所以最终转换解决问题的思路为寻找其他语言这方面的库，然后再集成进项目。最后，关于这种解决方案，因为我们项目实际要求（v90 工具集），所以我才增加纯 C 接口这一层封装，不然项目没办法调用，这增加了库的适用性。另外，我没有对这种解决方案的性能方面做任何测试，所以如果你对性能有要求的话，可以先去检索看有没有其他现成的 C++ 库。最后，如果你有任何问题或者建议的话，欢迎在 Issue 中提 bug 和 suggest 。
