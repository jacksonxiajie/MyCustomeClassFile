
#include "COperatorExcel.h"
#include<QDir>
#include<QFile>
#include<QFileInfo>
#include<QAxObject>
#include<qt_windows.h>

COperatorExcel::COperatorExcel(const QString &strFilePath, QObject *parent /*nullptr*/)
	:QObject(parent)
	, m_strFilePath(strFilePath)
{
	//如果文件不存在就新建一个内容为空的Excel文件
	if (!QFile::exists(m_strFilePath))
	{
		QFile file;
		file.setFileName(m_strFilePath);
		file.open(QIODevice::WriteOnly);
		file.close();
	}
}

COperatorExcel::~COperatorExcel()
{
	Close();
}

bool COperatorExcel::open(EOperateMode mode)
{
	bool bIsOpen = false;
	m_pExcelObj = new QAxObject(this);   //连接excell 控件
	m_pExcelObj->setControl("Excel.Application");
	m_pExcelObj->dynamicCall("SetVisible (bool Visible)", false);      //不显示窗体
	m_pExcelObj->setProperty("DisplayAlerts", false);                 //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示

	do
	{
		//获取工作薄
		m_pWorkbooks = m_pExcelObj->querySubObject("Workbooks");
		if (nullptr == m_pWorkbooks)
		{
			SAFEDELETE(m_pExcelObj)
				bIsOpen = false;
			break;
		}

		//打开工作薄
		if (EOperateMode::EReadOnly == mode)
		{
			m_pWorkbooks->dynamicCall("Open(const QString &strFilePath)", m_strFilePath);
		}
		else if (EOperateMode::ECreateNewFile == mode)
		{
			SetDefaultWorkSheetNums(m_iDefaultWorkSheetNum);
			m_pWorkbooks->dynamicCall("Add");
		}

		//获取当前活动工作薄
		m_pActiveBook = m_pExcelObj->querySubObject("ActiveWorkBook");
		if (nullptr == m_pActiveBook)
		{
			bIsOpen = false;
			SAFEDELETE(m_pExcelObj)
				break;
		}

		//获取工作表
		m_pWorkSheets = m_pActiveBook->querySubObject("Sheets");
		if (nullptr == m_pWorkSheets)
		{
			bIsOpen = false;
			SAFEDELETE(m_pWorkSheets)
				break;
		}

		bIsOpen = true;
	} while (false);

	return bIsOpen;
}

bool COperatorExcel::read(const unsigned int iSheetIndex, CTypeStrVarLstPair &strVarlstPair)
{
	if (nullptr == m_pWorkSheets) { return false; }
	QAxObject *pCurSheet = m_pWorkSheets->querySubObject("Item(int)", iSheetIndex); //获取指定工作表
	if (nullptr == pCurSheet) { return false; }



	//获取工作表中的数据
	QAxObject *usedRange = pCurSheet->querySubObject("UsedRange");
	if (nullptr == usedRange) { return false; }
	QVariant varValue = usedRange->dynamicCall("Value");
	QVariantList varLst = varValue.toList();
	if (varLst.isEmpty()) { return false; }

	//读取工作表的表名
	const QString sheetName = pCurSheet->property("Name").toString();

	//存储获取到的数据
	strVarlstPair.first = sheetName;
	strVarlstPair.second = varLst;

	return true;
}


bool COperatorExcel::write(const CTypeQPairLst &sheetDataLst)
{
	//写入数据至工作表
	bool bRet = false;
	do
	{
		QAxObject *pSheetItem = m_pWorkSheets->querySubObject("Item(int)", 1);   //获取第一个工作表
		if (nullptr == pSheetItem) { bRet = false; break; }

		const unsigned int iSheetCnt = m_pWorkSheets->property("Count").toInt();   //获取工作表的总数

		QAxObject *usedRangeObj = pSheetItem->querySubObject("UsedRange");
		if (nullptr != usedRangeObj) { usedRangeObj->dynamicCall("ClearContents()"); }   //清除工作表中的已有内容

		QAxObject *pTmpSheet = nullptr;
		QAxObject *pTmpUsedRange = Q_NULLPTR;
		unsigned int   iSheetItemIndex = 0;
		foreach(auto &tmpPair, sheetDataLst)
		{
			++iSheetItemIndex;
			if (iSheetItemIndex > iSheetCnt)      //添加新的工作表
			{
				appendSheet(m_pWorkSheets);
			}

			pTmpSheet = m_pWorkSheets->querySubObject("Item(int)", iSheetItemIndex);
			if (Q_NULLPTR == pTmpSheet) { return false; }

			//设置工作表的名称
			pTmpSheet->setProperty("Name", tmpPair.first);

			pTmpUsedRange = pTmpSheet->querySubObject("UsedRange");
			if (nullptr != usedRangeObj) { usedRangeObj->dynamicCall("ClearContents()"); }

			const unsigned int iRowCnt = tmpPair.second.toList().count();   //写入数据的行数
			const unsigned int iColCnt = tmpPair.second.toList().at(0).toList().count();  //写入数据的列数

																						  //获取存储数据的范围
			pTmpUsedRange = pTmpSheet->querySubObject("Range(const QString)", QString("%1%2:%3%4")
				.arg("A")  //起始列
				.arg(1)     //起始行
				.arg(convertIntToExcelColStr(iColCnt))  //结束列
				.arg(iRowCnt)); //数据占用的总行数(结束行)

								//设置工作表的内容
			pTmpUsedRange->setProperty("Value", tmpPair.second);

			//设置工作表的样式
			SetSheetStyle(pTmpUsedRange, 30, 60, 40);
		}
	} while (false);

	return (SaveAs(m_strFilePath)); //保存工作表
}

QAxObject *COperatorExcel::appendSheet(QAxObject *pSheetObj)
{
	if (nullptr == pSheetObj) { return  nullptr; }
	const int  sheetsCnt = pSheetObj->property("Count").toInt();
	QAxObject  *pLastSheet = pSheetObj->querySubObject("Item(int)", sheetsCnt);
	if (nullptr == pLastSheet) { return nullptr; }
	pSheetObj->dynamicCall("Add(QVariant)", pLastSheet->asVariant());   //在最后一个的前面插入新sheet
	QAxObject *pNewSheet = pSheetObj->querySubObject("Item(int)", sheetsCnt);
	pLastSheet->dynamicCall("Move(QVariant)", pNewSheet->asVariant());   //将新表移到最后
	return pNewSheet;
}

void COperatorExcel::Close()
{
	//关闭之前先保存数据
	SaveAs(m_strFilePath);

	//关闭当前工作薄
	if (nullptr != m_pActiveBook)
	{
		m_pActiveBook->dynamicCall("Close()");
		SAFEDELETE(m_pActiveBook)
	}

	//退出Excel 应用程序
	if (nullptr != m_pExcelObj)
	{
		m_pExcelObj->dynamicCall("Close()");
		m_pExcelObj->dynamicCall("Quit()");
		SAFEDELETE(m_pExcelObj)
	}

}

bool COperatorExcel::SetDefaultWorkSheetNums(unsigned int sheetCnt)
{
	if (nullptr == m_pExcelObj) { return false; }
	if (0 == sheetCnt) { sheetCnt = 1; }  //设置默认工作表的最小数目为1个
	m_pExcelObj->dynamicCall("SetSheetsInNewWorkbook(const unsigned int &num)", sheetCnt);
	return true;
}

bool COperatorExcel::SetCellTextFormat(const unsigned int iSheetIndex, const QString &strCellCol,
	const unsigned int iCellRow, const QString &strTextFormat)
{
	if (nullptr == m_pWorkSheets) { return false; }
	QAxObject *pActiveSheet = m_pWorkSheets->querySubObject("Item(unsigned int))", iSheetIndex);
	if (nullptr == pActiveSheet) { return false; }
	QAxObject * pCell = pActiveSheet->querySubObject("Range(const QString)", QString("%1%2")
		.arg(strCellCol)
		.arg(iCellRow));
	if (nullptr == pCell) { return false; }
	bool bRet = pCell->setProperty("NumberFormatLocal", strTextFormat);
	return bRet;
}

bool COperatorExcel::SetSheetStyle(const unsigned int iSheetIndex
	, const unsigned int iFontSize
	, const unsigned int iDefaultColWidth
	, const unsigned int iDefaultRowHeight)
{
	if (nullptr == m_pWorkSheets) { return false; }
	QAxObject *pActiveSheet = m_pWorkSheets->querySubObject("Item(unsigned int))", iSheetIndex);
	if (nullptr == pActiveSheet) { return false; }
	QAxObject *pTmpUsedRange = pActiveSheet->querySubObject("UsedRange");
	if (nullptr == pTmpUsedRange) { return false; }

	bool bRet = pTmpUsedRange->querySubObject("Font")->setProperty("Size", iFontSize);
	bRet = pTmpUsedRange->querySubObject("Columns")->setProperty("ColumnsWidth", iDefaultColWidth);   //设置默认列宽
	bRet = pTmpUsedRange->querySubObject("Rows")->setProperty("RowHeight", iDefaultRowHeight);       //设置默认行高
	return bRet;
}

bool COperatorExcel::SaveAs(const QString &filePath)
{
	if ((nullptr == m_pActiveBook) || (filePath.isEmpty())) { return false; }
	m_pActiveBook->dynamicCall("SaveCopyAs(QString)", QDir::toNativeSeparators(filePath));
	m_pActiveBook->dynamicCall("Close()");
	return true;
}

QString COperatorExcel::convertIntToExcelColStr(unsigned int iColIndex)
{
	QString retStr = "";
	if (0 == iColIndex) {
		iColIndex = 1;
	}
	if (0 != (iColIndex / 26))
	{
		retStr = QString(char('A') + char(iColIndex / 26 - 1));
	}

	retStr += QString(char('A') + char(iColIndex % 26 - 1));
	return retStr.trimmed();
}

bool COperatorExcel::SetSheetStyle(QAxObject *pActiveSheet
	, const unsigned int iFontSize
	, const unsigned int iDefaultColWidth
	, const unsigned int iDefaultRowHeight)
{
	if (nullptr == pActiveSheet) { return false; }
	bool bRet = pActiveSheet->querySubObject("Font")->setProperty("Size", iFontSize);
	bRet = pActiveSheet->querySubObject("Columns")->setProperty("ColumnsWidth", iDefaultColWidth);   //设置默认列宽
	bRet = pActiveSheet->querySubObject("Rows")->setProperty("RowHeight", iDefaultRowHeight);       //设置默认行高
	return bRet;
}
