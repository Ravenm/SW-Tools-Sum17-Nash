/***********************************************************************
The program below is taken the document "A brief introduction 
to C++ and Interfacing with Excel" by Andrew L. Hazel.  Very few
modifications were made to the code.  The complete document can be
found at http://www.maths.manchester.ac.uk/~ahazel/EXCEL_C++.pdf
Additional steps are needed to use Visual Studio Express Edition.
See footnote on page 58 of text.

***********************************************************************/

// Include standard libraries

#include<iostream>
#include<cmath>
#include<stdlib.h>
#include<time.h>

// Import necessary Excel libraries.  Adjust the paths as necessary.

#import "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE15\MSO.DLL" \
	rename("DocumentProperties", "DocumentPropertiesXL") \
	rename("RGB", "RBGXL")

#import  "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA6\VBE6EXT.OLB"

#import  "C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE" \
	rename("DialogBox", "DialogBoxXL") \
	rename("RGB", "RBGXL") \
	rename("DocumentProperties", "DocumentPropertiesXL") \
	rename("ReplaceText", "ReplaceTextXL")	 \
	rename("CopyFile", "CopyFileXL") \
	exclude("IFont", "IPicture") no_dual_interfaces

using namespace std;

// Simple function to graph as an example of using Excel charting
// tools from within C++

double f(const double &x) { return (sin(x)*exp(-x)); }

int main()
{
	//Surround the entire interfacing code with a try block
	try
	{
		//Initialise the COM interface
		CoInitialize(NULL);
		//Define a pointer to the Excel	application
		Excel::_ApplicationPtr xl;
		//Start	one instance of Excel
		xl.CreateInstance(L"Excel.Application");
		//Make the Excel application visible
		xl->Visible = true;
		//Add a(new)workbook
		xl->Workbooks->Add(Excel::xlWorksheet);
		//Get a pointer	to the active worksheet
		Excel::_WorksheetPtr pSheet = xl->ActiveSheet;
		//Set the name of the sheet
		pSheet->Name = "Excel Assignment 2";
		//Get a pointer to the cells on the active worksheet
		Excel::RangePtr pRange = pSheet->Cells;
		//Create two columns of data in the worksheet
		//We put labels at the top of each column to say what it contains
		pRange->Item[1][1] = "Andrew Nash";
		pRange->Item[1][2] = "Column A";
		pRange->Item[1][3] = "Column B";
		//Now we fill in the rest of the actual data by
		//using a single for loop
		//initialize random seed:
  		srand (time(NULL));
		for (unsigned i = 0; i<50; i++)
		{
			//The first column is our 1-50 values
			pRange->Item[i + 2][2] = i+1;
			//The second column is random
			pRange->Item[i + 2][3] = rand() % 50 + 1;
		}

	}
	//If there has been an error, say so
	catch (_com_error &error)
	{
		cout << "COM ERROR" << endl;
	}
	//Finally Uninitialise the COM interface
	CoUninitialize();
	//Finish the C++ program
	return 0;
}
