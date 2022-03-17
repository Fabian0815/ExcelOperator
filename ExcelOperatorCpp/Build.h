#pragma once

#ifdef EXCELOPERATOR_EXPORT
#if defined(_MSC_VER)
#define EXCELOPERATOR_API __declspec(dllexport)
#else
#define EXCELOPERATOR_API
#endif
#else
#if defined(_MSC_VER)
#define EXCELOPERATOR_API __declspec(dllimport)
#else
#define EXCELOPERATOR_API
#endif
#endif
