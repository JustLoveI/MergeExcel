//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Unit1.h"
#include <vector.h>
#include <inifiles.hpp>
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "SUIButton"
#pragma link "SUIEdit"
#pragma link "SUIProgressBar"
#pragma resource "*.dfm"
TForm1 *Form1;
TStringList *lstFile;
#define   PG   OlePropertyGet
#define   PS   OlePropertySet
#define   FN   OleFunction
#define   PR   OleProcedure
Variant vExcelAppOpen;
Variant vExcelAppSave;
Variant WorkBook1;
Variant SheetSave;
Variant SheetOpen;
Variant workbook1;
Variant Range;

struct ColumnName
{
        String sName; //列名称
        int ReadCol;
        int nLen;     //列长度
        bool bNumber;
};
vector<ColumnName> vCol;
int nColNum;  //列的数量
int nCountRow;//表行数
bool bNeedTitle;
String sTitle;
int nBeginRow;
int nCountExcel;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
        String currentdir = ExtractFileDir(Application->ExeName);
        String filepath = currentdir+"\\config.ini";
        TIniFile *ini = new TIniFile(filepath);
        String sShow = "";
        try
        {
               vCol.clear();
               nColNum = StrToInt(ini->ReadString("Column","ColumnNumber","")) ;        //读取列的数量  最后一列放文件名
               String ColName = ini->ReadString("Column","ColumnName","");
               String ColLen = ini->ReadString("Column","ColumnLen","");
               String ColBeNumber = ini->ReadString("Column","BeNumber","");
               String ReadCol = ini->ReadString("Column","ReadCol","");
               nBeginRow = StrToInt(ini->ReadString("BeginRow","BeginRow",""));
               ColumnName Col;
               for(int i = 1; i <= nColNum; i++)
               {
                        // 分割每个列
                        Col.sName = ColName.SubString(1,ColName.Pos(":")-1);
                        Col.nLen =  StrToInt(ColLen.SubString(1,ColLen.Pos(":")-1));
                        Col.bNumber = StrToInt(ColBeNumber.SubString(1,ColBeNumber.Pos(":")-1));
                        Col.ReadCol = StrToInt(ReadCol.SubString(1,ReadCol.Pos(":")-1));

                        ColName = ColName.SubString(ColName.Pos(":")+1,ColName.Length());
                        ColLen =  ColLen.SubString(ColLen.Pos(":")+1,ColLen.Length());
                        ColBeNumber = ColBeNumber.SubString(ColBeNumber.Pos(":")+1, ColBeNumber.Length());
                        ReadCol = ReadCol.SubString(ReadCol.Pos(":")+1, ReadCol.Length());
                        vCol.push_back(Col);
               }
               //添加从哪个文件获取的数据
               Col.sName = "文件名";
               Col.nLen =  30;
               Col.bNumber = 0;
               Col.ReadCol = 1;
               vCol.push_back(Col);
               nColNum++;
               sTitle = ini->ReadString("TiTle","TiTleName","");
               bNeedTitle = true;
               if(sTitle.IsEmpty())
               {
                        bNeedTitle = false;
               }
        }
        catch(...)
        {
                MessageBox(Application->Handle,"读取配置文件test.ini失败!","信息提示!",MB_OK+MB_ICONINFORMATION+MB_SYSTEMMODAL);
                delete ini;
        }
        suiProgressBar1->Position = 0;
        sShow += "标题\t\t";
        if(sTitle.IsEmpty()) sShow += "无\n";
        else sShow += sTitle + "\n";
        sShow += "列名\t\t";
        for(int i = 0; i < vCol.size(); i++)
        {
                sShow += vCol[i].sName + "\t";
        }
        sShow+="\n列长\t\t";
        for(int i = 0; i <vCol.size(); i++)
        {
                sShow += IntToStr(vCol[i].nLen) + "\t";
        }
        sShow += "\n读取字段\t";
        for(int i = 0; i <vCol.size(); i++)
        {
                if(i < vCol.size() -1)
                        sShow += IntToStr(vCol[i].ReadCol) + "\t";
                else
                        sShow += "文件名";
        }
        sShow += "\n读取从\t第" + IntToStr(nBeginRow) + "行开始";

        Label4->Caption = sShow;
}
//---------------------------------------------------------------------------

void __fastcall TForm1::suiButton2Click(TObject *Sender)
{
         if (OpenDialog1->Execute())
         {
              String path = OpenDialog1->FileName;
              path = path.SubString(1,path.LastDelimiter('\\')-1);
              suiEdit1->Text = path;
         }
}
//---------------------------------------------------------------------------
void __fastcall TForm1::suiButton1Click(TObject *Sender)
{
        if (dlgSave1->Execute())
        {
                String path = dlgSave1->FileName;
                suiEdit2->Text = path;
        }

}
//---------------------------------------------------------------------------
void __fastcall TForm1::ExportExcel(String Path)
{
        try
        {
                vExcelAppOpen = Variant::CreateObject("Excel.Application");
        }
         catch(...)
        {
                ShowMessage("启动 Excel 出错, 可能是没有安装Excel.");
                vExcelAppOpen = Unassigned;
                return;
        }

        try
        {
                vExcelAppOpen.PG("WorkBooks").PR("Open", Path.c_str());
                SheetOpen = vExcelAppOpen.PG("ActiveSheet");
         
                String sWjj = Path.SubString(Path.LastDelimiter('\\')+1,Path.LastDelimiter('.')-1);//文件夹名称
                sWjj = sWjj.SubString(1,sWjj.LastDelimiter('.')-1);
       
                for(int i = nBeginRow; ; i++)
                {

                        String  ContentTmp = "";
                        ContentTmp =  SheetOpen.PG("Cells",i,1).PG("Value");
                        if(ContentTmp == "" )
                        {
                                break;
                        }
                        for(int j = 0; j < vCol.size() - 1; j++)
                        {
                                ContentTmp = SheetOpen.PG("Cells",i,vCol[j].ReadCol).PG("Value");
                                SheetSave.PG("Cells",nCountRow,j+1).PS("Value",ContentTmp.c_str());
                        }
                        SheetSave.PG("Cells",nCountRow,nColNum).PS("Value",sWjj.c_str());
                        nCountRow++;
                }
        }
        catch(...)
        {
                String sShowMessage = "合并文件" + Path + " 出错！";
                ShowMessage(sShowMessage);
                vExcelAppOpen.OleFunction("Quit");
                vExcelAppOpen = Unassigned;
        }
        //关闭vExcelAppOpen
        vExcelAppOpen.OleFunction("Quit");
        vExcelAppOpen = Unassigned;
}
void __fastcall TForm1::suiButton3Click(TObject *Sender)
{

        try
        {
                vExcelAppSave = Variant::CreateObject("Excel.Application");
        }
        catch(...)
        {
                 ShowMessage("启动 Excel 出错, 可能是没有安装Excel.");
                 vExcelAppSave = Unassigned;
                 return;
        }
        suiProgressBar1->Position = 0;//进度条
        nCountRow = 0;
        vExcelAppSave.OlePropertyGet("Workbooks").OleFunction("Add", 1); // 工作表
        SheetSave = vExcelAppSave.PG("ActiveSheet");
        nCountRow = 1;
        char Buffer[10];
        if(bNeedTitle)
        {
                String strRang = "";       //合并单元格
                Buffer[0] = 'A'-1;
                for(int i = 1; i <= nColNum; i++)
                {
                        if(i != 1) strRang += ":";
                        Buffer[0]++;
                        strRang += String::StringOfChar(Buffer[0],1) + "1";
                        
                }
                Range = SheetSave.PG("Range",strRang.c_str()).OleFunction("Merge",false);
                vExcelAppSave.PG("Cells",1,1).PS("HorizontalAlignment",-4108);
                SheetSave.PG("Cells",1,1).PS("Value",AnsiString(sTitle).c_str());
                nCountRow ++;
        }
        //设置好列名，长度
        for(int i = 0; i <vCol.size()-1; i++)
        {
                SheetSave.PG("Cells",nCountRow, i+1).PS("Value", vCol[i].sName.c_str());
                if(vCol[i].bNumber)
                        vExcelAppSave.PG("Columns",i+1).PS("NumberFormat","@");
                vExcelAppSave.PG("Cells",nCountRow, i+1).PS("ColumnWidth", StrToInt(vCol[i].nLen));    //设置行宽
        }
        SheetSave.PG("Cells",nCountRow, nColNum).PS("Value", "文件名");
        vExcelAppSave.PG("Cells",nCountRow, nColNum).PS("ColumnWidth", 40);   
        nCountRow ++;
        String path = suiEdit1->Text;
        TSearchRec sr;
        int nExcelNum = 0;
        if (FindFirst(path + "\\*.xls", faAnyFile, sr) == 0)      //查询目录下的excel数目
        {
                do{
                        if(sr.Name[1] != '~')
                        {
                            nExcelNum ++;
                        }
                }
                while(FindNext(sr) == 0);
                FindClose(sr);
        }
        nCountExcel = 0;
        if (FindFirst(path + "\\*.xls", faAnyFile, sr) == 0)      //查询目录下的excel
        {
                do{
                        String Path = path + "\\" + sr.Name;
                        if(sr.Name[1] != '~')
                        {
                            nCountExcel ++;
                            ExportExcel(Path);
                            suiProgressBar1->Position = int(nCountExcel*1.0/nExcelNum*100);
                        }
                }
                while(FindNext(sr) == 0);
                FindClose(sr);
        }
        Variant ERange,EBorders;
        AnsiString strRange;
        strRange = "A"+IntToStr(1)+":"+String::StringOfChar(Buffer[0],1)+IntToStr(nCountRow-1); //获取操作范围    第 A列 - 第F列   第1行 - 第count +2 行
        ERange = vExcelAppSave.OlePropertyGet("Range",strRange.c_str());
        EBorders = ERange.OlePropertyGet("Borders");
        EBorders.OlePropertySet("linestyle",1); //线型
        EBorders.OlePropertySet("weight",2);    //粗细 值<=5
        EBorders.OlePropertySet("colorindex",0);
        String strXlsFile = suiEdit2->Text ;
        // 保存这个Excel文件
        vExcelAppSave.OlePropertyGet("ActiveWorkbook")
        .OleFunction("SaveAs", strXlsFile.c_str());
        vExcelAppSave.OleFunction("Quit");
        vExcelAppSave = Unassigned;
        suiProgressBar1->Position = 100;
        //lstFile->SaveToFile(suiEdit1->Text.c_str());
        ShowMessage("合并完成");
}
//---------------------------------------------------------------------------
