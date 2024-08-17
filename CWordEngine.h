#ifndef CWORDENGINE_H
#define CWORDENGINE_H

#include<QVariant>
#include<QColor>
#include<QObject>

class QAxObject;


//自定义word操作类
class CWordEngine:public QObject
{
     Q_OBJECT

//公开类型
public:  
    
    //枚举类型: 段落对齐; 
    enum  EParagraphAlign
    {
        Left = 0,   //左对齐
        Center,    //居中对齐
        Right,     //右对齐
        Justify    //两端对齐
    };
    
    //枚举类型: 页面布局
    enum  EPageSetup{
        wdOrientLandscape = 0,  //水平布局
        wdOrientPortrait        //垂直布局
    };   
    
    //枚举类型：文本样式
    enum  ETextStyle{
        title_one = 0,  //标题1
        title_two ,     //标题2
        title_three ,   //标题3
        title_four ,    //标题4
        title_five ,    //标题5
        subHeading ,    //副标题
        textBody ,      //正文
        explicitReference //明显引用
    };

    //枚举类型：移动鼠标至XX末尾
    enum  ECursorEndof{
        paragraph = 4,  //移动至段落末尾
        partial   = 6,  //选中部分的末尾
        curLine   = 10, //当前行的末尾
        table     = 15, //表格的末尾
    };
    
public:
    CWordEngine(const QString &absoluteFilePath,QObject *parent = nullptr);
    ~CWordEngine();

    //打开Word,默认word文件不存在
    bool open();

    //读取word中的数据
    bool read( QVariant &varData);

    //关闭word文件
    bool close();
    
    //保存
    bool save();

    //添加文本
    bool addText(const QString &StrText,ETextStyle textStyle);

    //插入图片
    bool addPicture(const QString &strImagePath);
    
    //插入回车换行
     bool insertEnterKey();

    //移动鼠标至末尾
     bool moveCursorToEnd(ECursorEndof typeOfEnd);

    //设置选中项的字体样式：字体大小，字体颜色，是否加粗
    bool setFontStyle(const int fontSize,bool bIsBold=false,const QColor &fontColor = QColor(Qt::black));

    //设置段落对齐方式
    bool setParagraAlignFormat(EParagraphAlign alignment);
 
    //设置页面布局方式
    bool setPageAlignment(EPageSetup pageSetUp);
    

public:
    
    static bool IsExist(const QString &strFile);
    
protected:
    
    //获取当前位置
    QAxObject* CurSelection();

    //将QColor 转为整型
    inline  int ConvertQColorToInt(const QColor& color);

    
private:
    QAxObject       *m_pWordWgt = nullptr;   //word 外壳控件
    QAxObject       *m_pDocument = nullptr;  //word文档对象
    QAxObject       *m_pActiveDoc  = nullptr;  //当前活动文档对象

    QString            m_fileName = "";   //文件名
};

#endif // CWORDENGINE_H
