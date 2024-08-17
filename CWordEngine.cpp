#include "CWordEngine.h"


#include<QAxObject>
#include<QDir>
#include<QFile>
#include<QVariantList>


CWordEngine::CWordEngine(const QString &absoluteFilePath,QObject *parent)
    :QObject(parent)
    ,m_fileName(absoluteFilePath)
{

}

CWordEngine::~CWordEngine()
{
    close();
}

bool CWordEngine::open()
{
   //未关闭就先关闭
	close();

   m_pWordWgt = new QAxObject(this);
   m_pWordWgt->setControl("Word.Application");   //设置操作的组件应用
   m_pWordWgt->dynamicCall("SetVisible (bool Visible)","false");      //不显示窗体
   m_pWordWgt->setProperty("DisplayAlerts", false);                      //不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示

   QString Text = m_pWordWgt->generateDocumentation();
    //获取文档对象
   m_pDocument = m_pWordWgt->querySubObject("Documents");  
   if(  nullptr == m_pDocument )
   {     
       return false;
   }

   //判断文件是否已存在：存在-打开该文件；不存在-新建该文件
   if( IsExist( m_fileName ) )
   {
       m_pDocument->dynamicCall("Open(const QString &)",m_fileName );
   }
   else 
   {
      //文件不存在则新建一个模板文件
       QFile file(m_fileName);
       if(!file.open(QIODevice::WriteOnly | QIODevice::Text))
       {
           close();
           return false;
       }
       file.close();

        //以当前文件为模板新建一个文件
//         m_pDocument->dynamicCall("Add(QString)", m_fileName);
	   m_pDocument->dynamicCall("Open(const QString &)", m_fileName);
   }
  
   //获取当前活动文档
   m_pActiveDoc = m_pWordWgt->querySubObject("ActiveDocument");  
   if( nullptr == m_pActiveDoc )
   {
      return false;
   }
   
   return true;
}

bool CWordEngine::read( QVariant &varData)
{
    if( nullptr ==m_pActiveDoc ){ return false; }
  
   QAxObject  *pSelectObj =  m_pActiveDoc->querySubObject("Selection");
   if( nullptr == pSelectObj ) { return false; }
  
   varData= pSelectObj->property("Value");

   return true;
}

bool CWordEngine::close()
{
    if( nullptr == m_pActiveDoc )    {  return false; }

    //关闭之前先保存数据
    save();

    m_pActiveDoc->dynamicCall("Close(bool)",true);
    delete m_pActiveDoc;
    m_pActiveDoc = nullptr;

    if(nullptr == m_pWordWgt)  { return false; }
    m_pWordWgt->dynamicCall("Quit()");
   delete m_pWordWgt;
    m_pWordWgt = nullptr;
    return true;
}

bool CWordEngine::addText(const QString &StrText, ETextStyle textStyle)
{

    QAxObject *pCurSelection = CurSelection();
    if( nullptr == pCurSelection ){ return false; }
    
    QAxObject *pRange = pCurSelection->querySubObject("Range");
    if( nullptr == pRange) { return false; }
    //设置文本
    pRange->setProperty("Text",StrText);
// 	pCurSelection->dynamicCall("TypeText(const QString&)", StrText);

    //设置文本样式
    QString strTextStyle = "";
    
    switch (textStyle) {
    case ETextStyle::title_one :
        strTextStyle = QStringLiteral("标题1");
        break;
    case ETextStyle::title_two :
        strTextStyle = QStringLiteral("标题2");
        break;
    case ETextStyle::title_three :
        strTextStyle = QStringLiteral("标题3");
        break;
    case ETextStyle::title_four :
        strTextStyle = QStringLiteral("标题4");        
        break;
    case ETextStyle::title_five :
        strTextStyle = QStringLiteral("标题5");
        break;
    case ETextStyle::explicitReference :
        strTextStyle = QStringLiteral("明显引用");
        break;
    case ETextStyle::textBody :
        strTextStyle = QStringLiteral("正文");
        break;
    default:
        strTextStyle = QStringLiteral("正文");
        break;
    }
    
    pRange->dynamicCall("SetStyle(QString &)",strTextStyle);//||
    
    return true;
}

bool CWordEngine::addPicture(const QString &strImagePath)
{
    if(!IsExist(strImagePath)){ return false;}
    if(nullptr == m_pActiveDoc){ return false;}

    QAxObject *pInLinShapes = m_pActiveDoc->querySubObject("InlineShapes");
    if(nullptr == pInLinShapes){ return false; }

    QAxObject *pCurSelection = CurSelection();
    if(nullptr == pCurSelection){ return false;}

    QVariantList  imageDataLst;
    imageDataLst<<strImagePath<<false<<true<<pCurSelection->asVariant();
	pInLinShapes->dynamicCall("AddPicture(const QString&)", strImagePath);
    return true;
}

bool CWordEngine::insertEnterKey()
{
    QAxObject *pCurSelection = CurSelection();
    if(nullptr == pCurSelection){ return false;}
    QVariant retVar = pCurSelection->dynamicCall("TypeParagraph()");
    if(retVar.isValid())
    {
        return retVar.toBool();
    }
    return true;
}

bool CWordEngine::moveCursorToEnd(CWordEngine::ECursorEndof typeOfEnd)
{
    QAxObject *pCurSelection = CurSelection();
    if(nullptr == pCurSelection){ return false;}
    QVariant var = pCurSelection->dynamicCall("Endof(QVariant,QVariant)",typeOfEnd,0);
    return true;
}

bool CWordEngine::setFontStyle(const int fontSize, bool bIsBold, const QColor &fontColor)
{
    QAxObject *pCurSelection = CurSelection();
    if(nullptr == pCurSelection){ return false;}
    QAxObject *pRange = pCurSelection->querySubObject("Range");
    if(nullptr == pRange){ return  false;}
    QAxObject *pFont = pRange->querySubObject("Font");
    if(nullptr == pFont){ return false;}

    //设置字体大小
    bool bRet = false;
    bRet = pFont->setProperty("Size",QVariant(fontSize));
    if(!bRet){ return bRet;}

    //设置字体颜色
    bRet = pFont->setProperty("Color",QVariant(ConvertQColorToInt(fontColor)));
    if(!bRet){ return bRet;}

    //设置粗体
    bRet =  pFont->setProperty("Bold",QVariant(bIsBold));
    if(!bRet){ return bRet;}

    return true;
}

bool CWordEngine::setParagraAlignFormat(CWordEngine::EParagraphAlign alignment)
{
    QAxObject *pCurSelection = CurSelection();
    if( nullptr == pCurSelection ){ return false; }
    
    QAxObject *pParagraph = pCurSelection->querySubObject("ParagraphFormat");
    if( nullptr== pParagraph ) { return false; }
    
    //设置段落对齐方式
   bool bRet = pParagraph->setProperty("Alignment",alignment);
   return bRet;    
}

bool CWordEngine::setPageAlignment(CWordEngine::EPageSetup pageSetUp)
{
    QAxObject *pSelection = CurSelection();
    if( nullptr == pSelection ) { return false; }
    QAxObject *pPageSetup = pSelection->querySubObject("PageSetup");
    if( nullptr == pPageSetup ) { return false; }
    
    //设置页面布局方式
    QString strPageSetup = "";
    switch (pageSetUp) {
    case EPageSetup::wdOrientLandscape :
        strPageSetup = QStringLiteral("wdOrientLandscape");
        break;
    case EPageSetup::wdOrientPortrait :
        strPageSetup = QStringLiteral("wdOrientPortrait");
        break;
    default:
        strPageSetup = QStringLiteral("wdOrientPortrait");
        break;
    }
  bool bRet =  pPageSetup->setProperty("Orientation",strPageSetup);
  return bRet;
}


bool CWordEngine::save()
{
    if(nullptr == m_pActiveDoc){ return false;}
     m_pActiveDoc->dynamicCall("Save()");
   return true;
}

bool CWordEngine::IsExist(const QString &strFile)
{
    return (QFile::exists(strFile));
}

QAxObject *CWordEngine::CurSelection()
{
    if(nullptr == m_pWordWgt) { return nullptr; }
    QAxObject *pSelection = m_pWordWgt->querySubObject("Selection");   
    if(nullptr == pSelection){ return nullptr; }
    return  (pSelection);
}

 int CWordEngine::ConvertQColorToInt(const QColor &color)
{
    int iRet = 0;
    if( color.isValid() )
    {
        iRet = (color.red()<<16) + (color.green()<<8) + (color.blue()<<0) + (color.alpha()<<24);
    }
    return iRet;
}
