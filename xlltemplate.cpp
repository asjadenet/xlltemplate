// xlltemplate.cpp
#include <cmath>
#include "xlltemplate.h"

using namespace xll;

AddIn xai_template(
	Documentation(LR"(
This object will generate a Sandcastle Helpfile Builder project file.
)"));

// Information Excel needs to register add-in.
AddIn xai_function(
	// Function returning a pointer to an OPER with C name xll_function and Excel name XLL.FUNCTION.
	// Don't forget prepend a question mark to the C name.
	//                     v
    Function(XLL_LPOPER, L"?xll_function", L"XLL.FUNCTION")
	// First argument is a double called x with an argument description.
    .Arg(XLL_DOUBLE, L"x", L"is the first double argument.")
	// Paste function category.
    .Category(CATEGORY)
    .FunctionHelp(L"Help on XLL.FUNCTION goes here.")
	.Documentation(LR"(
Documentation on XLL.FUNCTION goes here.
    )")
);

void set_cell_address(XLOPER12& cell_ref, RW rw, COL col)
{
	cell_ref.xltype = xltypeSRef;
	cell_ref.val.sref.count = 1;
	cell_ref.val.sref.ref.rwFirst = rw;
	cell_ref.val.sref.ref.rwLast = rw;
	cell_ref.val.sref.ref.colFirst = col;
	cell_ref.val.sref.ref.colLast = col;
}

// Calling convention *must* be WINAPI (aka __stdcall) for Excel.
LPOPER WINAPI xll_function(double x)
{
// Be sure to export your function.
#pragma XLLEXPORT
	static OPER result;

	try {
		//Read cell B1 value:

		XLOPER12 reading_parameters;
		reading_parameters.xltype = xltypeInt;
		reading_parameters.val.w = 5; // contents of cell as number

		XLOPER12 reading_cell_ref;
		set_cell_address(reading_cell_ref, 1, 1);

		XLOPER12 result_value;
		Excel12(xlfGetCell, &result_value, 2, &reading_parameters, &reading_cell_ref);

		//Outputting result to cell F6

		XLOPER12 out_cell_ref;
		std::wstring hello_msg = L"Hello, world. Cell B2 value is:";
		if (result_value.xltype == xltypeStr)
		{
			auto len = result_value.val.str[0];
			std::wstring value(result_value.val.str + 1, len);
			hello_msg += value;
		}
		if (result_value.xltype == xltypeNum)
			hello_msg += std::to_wstring(result_value.val.num);
		OPER12 xValue(hello_msg);
		set_cell_address(out_cell_ref, 5, 5);
		Excel12(xlSet, nullptr, 2, &out_cell_ref, &xValue);

		result = x;
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		result = OPER(xlerr::Num);
	}

	return &result;
}
