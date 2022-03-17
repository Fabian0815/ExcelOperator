#include "Export.h"
#include "Workbook.h"
using namespace System::Reflection;

Export::Workbook* CreateWorkbook()
{
	return new WorkbookImpl();
}
void release_resource(Export::DllResource* resource)
{
	delete resource;
}

void release_resource_const(const Export::DllResource* resource)
{
	delete resource;
}
