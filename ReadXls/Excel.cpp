
//Excel.cpp
#include "stdafx.h"
#include <tchar.h>
#include "Excel.h"


COleVariant
covTrue((short)TRUE),
covFalse((short)FALSE),
covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

CApplication Excel::application;

Excel::Excel() :isLoad(false)
{
}


Excel::~Excel()
{
	//close();
}


bool Excel::initExcel()
{
	//创建Excel 2000服务器(启动Excel)   
	if (!application.CreateDispatch(_T("Excel.application"), nullptr))
	{
		MessageBox(nullptr, _T("创建Excel服务失败,你可能没有安装EXCEL，请检查!"), _T("错误"), MB_OK);
		return FALSE;
	}

	application.put_DisplayAlerts(FALSE);
	return true;
}


void Excel::release()
{
	application.Quit();
	application.ReleaseDispatch();
	application = nullptr;
}

bool Excel::open(const char* fileName)
{

	//先关闭文件
	close();

	//利用模板建立新文档
	books.AttachDispatch(application.get_Workbooks(), true);


	LPDISPATCH lpDis = nullptr;
	try{
	lpDis = books.Add(COleVariant(CString(fileName)));

	}
	catch(...)
    {
        /*增加一个新的工作簿*/
        lpDis = books.Add(vtMissing);
       
    }


	if (lpDis)
	{
		workBook.AttachDispatch(lpDis);

		sheets.AttachDispatch(workBook.get_Worksheets());

		openFileName = fileName;
		return true;
	}

	return false;
}

void Excel::close(bool ifSave)
{
	//如果文件已经打开，关闭文件
	if (!openFileName.IsEmpty())
	{
		//如果保存,交给用户控制,让用户自己存，如果自己SAVE，会出现莫名的等待  
		if (ifSave)
		{
			//show(true);
		}
		else
		{
			workBook.Close(COleVariant(short(FALSE)), COleVariant(openFileName), covOptional);
			books.Close();
		}

		//清空打开文件名称
		openFileName.Empty();
	}


	sheets.ReleaseDispatch();
	workSheet.ReleaseDispatch();
	currentRange.ReleaseDispatch();
	workBook.ReleaseDispatch();
	books.ReleaseDispatch();
}

void Excel::saveAsXLSFile(const CString &xlsFile)
{
	workBook.SaveAs(COleVariant(xlsFile),
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		0,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional);
	return;
}


int Excel::getSheetCount()
{
	return sheets.get_Count();
}

CString Excel::getSheetName(long tableID)
{
	CWorksheet sheet;
	sheet.AttachDispatch(sheets.get_Item(COleVariant((long)tableID)));
	CString name = sheet.get_Name();
	sheet.ReleaseDispatch();
	return name;
}

void Excel::useSheet(CString strSheetName)
{
	LPDISPATCH lpDisp = NULL;
	lpDisp = sheets.get_Item(_variant_t(strSheetName));
    workSheet.AttachDispatch(lpDisp);
}

void Excel::deleteSheet(CString strSheetName)
{
	LPDISPATCH lpDisp = NULL;
	lpDisp = sheets.get_Item(_variant_t(strSheetName));
    workSheet.AttachDispatch(lpDisp);
    workSheet.Delete();
}

bool Excel::addNewSheet(LPCTSTR newSheetName)
{
	bool flag = false;
	LPDISPATCH lpDisp = NULL;
	CWorksheet sheet;
	lpDisp = sheets.Add(vtMissing, vtMissing, _variant_t((long)1), vtMissing);
    if(lpDisp){
		sheet.AttachDispatch(lpDisp);
		sheet.put_Name(newSheetName);
		sheet.ReleaseDispatch();
		flag = true;
	}
	return flag;	
}

void Excel::setSheetName(long tableID,LPCTSTR newName)
{
	CWorksheet sheet;
	sheet.AttachDispatch(sheets.get_Item(COleVariant((long)tableID)));
	sheet.put_Name(newName);
	sheet.ReleaseDispatch();
}

void Excel::setCellString(long iRow, long iColumn, CString newString)
{

	COleVariant new_value(newString);
	CRange start_range = workSheet.get_Range(COleVariant(_T("A1")), covOptional);
	CRange write_range = start_range.get_Offset(COleVariant((long)iRow - 1), COleVariant((long)iColumn - 1));
	write_range.put_Value2(new_value);
	start_range.ReleaseDispatch();
	write_range.ReleaseDispatch();

}

void Excel::setCellInt(long iRow, long iColumn, int newInt)
{
	COleVariant new_value((long)newInt);
	CRange start_range = workSheet.get_Range(COleVariant(_T("A1")), covOptional);
	CRange write_range = start_range.get_Offset(COleVariant((long)iRow - 1), COleVariant((long)iColumn - 1));
	write_range.put_Value2(new_value);
	start_range.ReleaseDispatch();
	write_range.ReleaseDispatch();
}

void Excel::preLoadSheet()
{
	CRange used_range;

	used_range = workSheet.get_UsedRange();


	VARIANT ret_ary = used_range.get_Value2();
	if (!(ret_ary.vt & VT_ARRAY))
	{
		return;
	}
	//  
	safeArray.Clear();
	safeArray.Attach(ret_ary);
}

//按照名称加载sheet表格，也可提前加载所有表格
bool Excel::loadSheet(long tableId, bool preLoaded)
{
	LPDISPATCH lpDis = nullptr;
	currentRange.ReleaseDispatch();
	currentRange.ReleaseDispatch();
	lpDis = sheets.get_Item(COleVariant((long)tableId));
	if (lpDis)
	{
		workSheet.AttachDispatch(lpDis, true);
		currentRange.AttachDispatch(workSheet.get_Cells(), true);
	}
	else
	{
		return false;
	}

	isLoad = false;
	//如果进行预先加载  
	if (preLoaded)
	{
		preLoadSheet();
		isLoad = true;
	}

	return true;
}


bool Excel::loadSheet(CString sheet, bool preLoaded)
{
	LPDISPATCH lpDis = nullptr;
	currentRange.ReleaseDispatch();
	currentRange.ReleaseDispatch();

	lpDis = sheets.get_Item(COleVariant(sheet));
	if (lpDis)
	{
		workSheet.AttachDispatch(lpDis, true);
		currentRange.AttachDispatch(workSheet.get_Cells(), true);
	}
	else
	{
		return false;
	}

	isLoad = false;
	//如果进行预先加载  
	if (preLoaded)
	{
		preLoadSheet();
		isLoad = true;
	}

	return true;
}


int Excel::getColumnCount()
{
	CRange range;
	CRange usedRange;

	usedRange.AttachDispatch(workSheet.get_UsedRange(), true);
	range.AttachDispatch(usedRange.get_Columns(), true);
	int count = range.get_Count();

	usedRange.ReleaseDispatch();
	range.ReleaseDispatch();

	return count;
}

int Excel::getRowCount()
{
	CRange range;
	CRange usedRange;

	usedRange.AttachDispatch(workSheet.get_UsedRange(), true);
	range.AttachDispatch(usedRange.get_Rows(), true);

	int count = range.get_Count();

	usedRange.ReleaseDispatch();
	range.ReleaseDispatch();

	return count;
}

bool Excel::isCellString(long iRow, long iColumn)
{
	CRange range;
	range.AttachDispatch(currentRange.get_Item(COleVariant((long)iRow), COleVariant((long)iColumn)).pdispVal, true);
	COleVariant vResult = range.get_Value2();
	//VT_BSTR标示字符串  
	if (vResult.vt == VT_BSTR)
	{
		return true;
	}
	return false;
}


bool Excel::isCellInt(long iRow, long iColumn)
{

	CRange range;
	range.AttachDispatch(currentRange.get_Item(COleVariant((long)iRow), COleVariant((long)iColumn)).pdispVal, true);
	COleVariant vResult = range.get_Value2();
	//VT_BSTR标示字符串  
	if (vResult.vt == VT_INT || vResult.vt == VT_R8)
	{
		return true;
	}
	return false;
}

CString Excel::getCellString(long iRow, long iColumn)
{
	COleVariant vResult;
	CString str;
	//字符串  
	if (isLoad == false)
	{
		CRange range;
		range.AttachDispatch(currentRange.get_Item(COleVariant((long)iRow), COleVariant((long)iColumn)).pdispVal, true);
		vResult = range.get_Value2();
		range.ReleaseDispatch();
	}
	//如果数据依据预先加载了  
	else
	{
		long read_address[2];
		VARIANT val;
		read_address[0] = iRow;
		read_address[1] = iColumn;
		safeArray.GetElement(read_address, &val);
		vResult = val;
	}

	if (vResult.vt == VT_BSTR)
	{
		str = vResult.bstrVal;
	}
	//整数  
	else if (vResult.vt == VT_INT)
	{
		str.Format(_T("%d"), vResult.pintVal);
	}
	else if (vResult.vt == VT_BOOL)
	{
		str.Format(_T("%s"), vResult.pintVal == FALSE ? "FALSE" : "TRUE");
	}

	
	//8字节的数字   
	else if (vResult.vt == VT_R8)
	{
		if ( ((int)(vResult.dblVal * 100) % 100) > 0)
		{
			str.Format(_T("%.2f"), vResult.dblVal);
		}
		else
		{
			str.Format(_T("%.0f"), vResult.dblVal);
		}
	}
	//时间格式  
	else if (vResult.vt == VT_DATE)
	{
		SYSTEMTIME st;
		VariantTimeToSystemTime(vResult.date, &st);
		CTime tm(st);
		str = tm.Format(_T("%Y-%m-%d"));

	}
	//单元格空的  
	else if (vResult.vt == VT_EMPTY)
	{
		str = "";
	}
	//未知类型
// 	{
// 		str = vResult.bstrVal;
// 	}

	return str;
}

double Excel::getCellDouble(long iRow, long iColumn)
{
	double rtn_value = 0;
	COleVariant vresult;
	//字符串  
	if (isLoad == false)
	{
		CRange range;
		range.AttachDispatch(currentRange.get_Item(COleVariant((long)iRow), COleVariant((long)iColumn)).pdispVal, true);
		vresult = range.get_Value2();
		range.ReleaseDispatch();
	}
	//如果数据依据预先加载了  
	else
	{
		long read_address[2];
		VARIANT val;
		read_address[0] = iRow;
		read_address[1] = iColumn;
		safeArray.GetElement(read_address, &val);
		vresult = val;
	}

	if (vresult.vt == VT_R8)
	{
		rtn_value = vresult.dblVal;
	}

	return rtn_value;
}

int Excel::getCellInt(long iRow, long iColumn)
{
	int num;
	COleVariant vresult;

	if (isLoad == FALSE)
	{
		CRange range;
		range.AttachDispatch(currentRange.get_Item(COleVariant((long)iRow), COleVariant((long)iColumn)).pdispVal, true);
		vresult = range.get_Value2();
		range.ReleaseDispatch();
	}
	else
	{
		long read_address[2];
		VARIANT val;
		read_address[0] = iRow;
		read_address[1] = iColumn;
		safeArray.GetElement(read_address, &val);
		vresult = val;
	}
	//  
	num = static_cast<int>(vresult.dblVal);

	return num;
}
void Excel::ATest()
{

    // TODO: Add your control notification handler code here
	CApplication ExcelApp;
	CWorkbooks books;  //多个xls文件
	CWorkbook book;  //一个xls文件
	CWorksheets sheets;  //sheet的集合
	CWorksheet sheet;     //单个sheet
    CRange range;           //一个格子
    LPDISPATCH lpDisp = NULL;
    //创建Excel 服务器(启动Excel)
    if(!ExcelApp.CreateDispatch(_T("Excel.Application"),NULL))
    {
        AfxMessageBox(_T("启动Excel服务器失败!"));
        return;
    }
	//ExcelApp.put_Visible(false);
	//ExcelApp.put_UserControl(FALSE);
    /*得到工作簿容器*/
	books.AttachDispatch(ExcelApp.get_Workbooks());
    /*打开一个工作簿，如不存在，则新增一个工作簿*/
    CString strBookPath =_T("D:\\tmpl.xls");
    try
    {
        /*打开一个工作簿*/
        lpDisp = books.Open(strBookPath, 
            vtMissing, vtMissing,vtMissing, vtMissing, vtMissing,
            vtMissing, vtMissing,vtMissing, vtMissing, vtMissing, 
            vtMissing, vtMissing,vtMissing, vtMissing);
        book.AttachDispatch(lpDisp);
    }
    catch(...)
    {
        /*增加一个新的工作簿*/
        lpDisp = books.Add(vtMissing);
        book.AttachDispatch(lpDisp);
    }
/*得到工作簿中的Sheet的容器*/
	sheets.AttachDispatch(book.get_Sheets());
    /*打开一个Sheet，如不存在，就新增一个Sheet*/
    CString strSheetName =_T("lalalalh");
    try
    {
        /*打开一个已有的Sheet*/
		lpDisp = sheets.get_Item(_variant_t(strSheetName));
        sheet.AttachDispatch(lpDisp);
    }
    catch(...)
    {
        /*创建一个新的Sheet*/
        lpDisp = sheets.Add(vtMissing, vtMissing, _variant_t((long)1), vtMissing);
        sheet.AttachDispatch(lpDisp);
		sheet.put_Name(strSheetName);
    }
    /*向Sheet中写入多个单元格,规模为4*30 */
	lpDisp = sheet.get_Range(_variant_t("C6"), _variant_t("F35"));
    range.AttachDispatch(lpDisp);
    VARTYPE vt = VT_R4; /*数组元素的类型，float*/
    SAFEARRAYBOUND sabWrite[2]; /*用于定义数组的维数和下标的起始值*/
    sabWrite[0].cElements = 30;
    sabWrite[0].lLbound = 0;
    sabWrite[1].cElements = 3;
    sabWrite[1].lLbound = 0;
    COleSafeArray olesaWrite;
    olesaWrite.Create(vt, sizeof(sabWrite)/sizeof(SAFEARRAYBOUND), sabWrite);
    /*通过指向数组的指针来对二维数组的元素进行间接赋值*/
    float (*pArray)[2]= NULL;
    olesaWrite.AccessData((void **)&pArray);
    memset(pArray, 0, sabWrite[0].cElements * sabWrite[1].cElements* sizeof(float));
    /*释放指向数组的指针*/
    olesaWrite.UnaccessData();
    pArray = NULL;
    /*对二维数组的元素进行逐个赋值*/
    long index[2]= {0, 0};
    long lFirstLBound = 0;
    long lFirstUBound = 0;
    long lSecondLBound = 0;
    long lSecondUBound = 0;
    olesaWrite.GetLBound(1, &lFirstLBound);
    olesaWrite.GetUBound(1, &lFirstUBound);
    olesaWrite.GetLBound(2, &lSecondLBound);
    olesaWrite.GetUBound(2, &lSecondUBound);
    long i = 0;
    for (i = lFirstLBound;i <= lFirstUBound; i++)
    {
        index[0] = i;
        for (long j =lSecondLBound; j <= lSecondUBound; j++)
        {
            index[1] = j;
            float lElement = (float)(i * sabWrite[1].cElements + j); 
            olesaWrite.PutElement(index, &lElement);
        }
    }
    /*把ColesaWritefeArray变量转换为VARIANT,并写入到Excel表格中*/
    VARIANT varWrite = (VARIANT)olesaWrite;
	range.put_Value2(varWrite);
	range.put_NumberFormat(COleVariant(_T("$0.00")));
    /*根据文件的后缀名选择保存文件的格式*/
    //CString strSaveAsName = _T("C:\\ew.xls");
    //CString strSuffix = strSaveAsName.Mid(strSaveAsName.ReverseFind(_T('.')));
    //XlFileFormat NewFileFormat = xlOpenXMLWorkbook;
    ////Excel::XlFileFormat NewFileFormat = xlWorkbookNormal;
    //if (0 ==strSuffix.CompareNoCase(_T(".xls")))
    //{
    //    NewFileFormat= xlExcel8;
    //}
	book.SaveAs(_variant_t(strBookPath), vtMissing, vtMissing, vtMissing, vtMissing,
        vtMissing, 0, vtMissing, vtMissing, vtMissing,
        vtMissing, vtMissing);
    //book.Save();
    /*释放资源*/
    sheet.ReleaseDispatch();
    sheets.ReleaseDispatch();
    book.ReleaseDispatch();
    books.ReleaseDispatch();
    ExcelApp.Quit();
    ExcelApp.ReleaseDispatch();
}








void Excel::show(bool bShow)
{
	application.put_Visible(bShow);
	application.put_UserControl(bShow);
}

CString Excel::getOpenFileName()
{
	return openFileName;
}

CString Excel::getOpenSheelName()
{
	return workSheet.get_Name();
}

char* Excel::getColumnName(long iColumn)
{
	static char column_name[64];
	size_t str_len = 0;

	while (iColumn > 0)
	{
		int num_data = iColumn % 26;
		iColumn /= 26;
		if (num_data == 0)
		{
			num_data = 26;
			iColumn--;
		}
		column_name[str_len] = (char)((num_data - 1) + 'A');
		str_len++;
	}
	column_name[str_len] = '\0';
	//反转  
	_strrev(column_name);

	return column_name;
}

