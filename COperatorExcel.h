#ifndef COPERATOREXCEL_H
#define COPERATOREXCEL_H

#include<QObject>
#include<QVariant>
class QAxObject;

//自定义数据类型：QPair<QString,QVariantList>-first:工作表名称，second:工作表的内容
typedef QPair<QString, QVariantList >  CTypeStrVarLstPair;

//自定义数据类型：QList< QPair<QString,QVariant>  >  存储所有当前工作薄中的所有工作表信息
typedef QList< QPair<QString, QVariant>  >  CTypeQPairLst;

//自定义Excel 操作类
class COperatorExcel : public QObject
{
	Q_OBJECT

public:

	enum EOperateMode
	{
		EReadOnly = 1,    //读取模式
		ECreateNewFile  //创建新文件模式
	};

public:
	explicit COperatorExcel(const QString &strFilePath, QObject *parent = nullptr);
	~COperatorExcel();

	//打开Excel 文件
	bool open(EOperateMode mode);

	//读取指定工作表中的数据    
	bool read(const unsigned int iSheetIndex, CTypeStrVarLstPair &strVarlstPair);

	//向工作表中写入数据
	bool write(const CTypeQPairLst &sheetDataLst);

	//关闭Excel 应用并保存数据
	void Close();

	//设置创建一个工作表时默认的创建工作表的数目，需在open之前使用
	bool SetDefaultWorkSheetNums(unsigned int sheetCnt);

	//设置指定工作表中对应单元格的数字格式
	bool SetCellTextFormat(const unsigned int iSheetIndex
		, const QString &strCellCol
		, const unsigned int iCellRow
		, const QString &strTextStyle);

	//设置指定工作表的显示样式
	bool SetSheetStyle(const unsigned int iSheetIndex
		, const unsigned int iFontSize
		, const unsigned int iDefaultColWidth
		, const unsigned int iDefaultRowHeight);
signals:

	public slots :

protected:

	//保存写入的工作表数据
	bool SaveAs(const QString &filePath);

	//将列数转为Excel 中对应的字符列
	QString convertIntToExcelColStr(unsigned int iColIndex);

	//设置当前工作活动工作表的显示样式
	bool SetSheetStyle(QAxObject *pActiveSheet
		, const unsigned int iFontSize
		, const unsigned int iDefaultColWidth
		, const unsigned int iDefaultRowHeight);

	//新增工作表
	QAxObject*  appendSheet(QAxObject *pSheetObj);

private:
	QString             m_strFilePath = "";     //文件的路径
	unsigned int        m_iDefaultWorkSheetNum = 1;  //默认工作表的个数

	QAxObject       *m_pExcelObj = nullptr;    //excel 外壳
	QAxObject       *m_pWorkbooks = nullptr;    //excel工作薄
	QAxObject       *m_pActiveBook = nullptr;   //当前活动工作薄
	QAxObject       *m_pWorkSheets = nullptr;  //工作表

};

Q_DECLARE_METATYPE(CTypeQPairLst)
Q_DECLARE_METATYPE(CTypeStrVarLstPair)

#define SAFEDELETE(ptr)\
{\
    if(nullptr != (ptr))\
    {\
        delete (ptr);\
       (ptr) = nullptr;\
    }\
}

#endif // COPERATOREXCEL_H
