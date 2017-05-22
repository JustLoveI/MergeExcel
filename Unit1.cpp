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
        String sName; //������
        int ReadCol;
        int nLen;     //�г���
        bool bNumber;
};
vector<ColumnName> vCol;
int nColNum;  //�е�����
int nCountRow;//������
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
               nColNum = StrToInt(ini->ReadString("Column","ColumnNumber","")) ;        //��ȡ�е�����  ���һ�з��ļ���
               String ColName = ini->ReadString("Column","ColumnName","");
               String ColLen = ini->ReadString("Column","ColumnLen","");
               String ColBeNumber = ini->ReadString("Column","BeNumber","");
               String ReadCol = ini->ReadString("Column","ReadCol","");
               nBeginRow = StrToInt(ini->ReadString("BeginRow","BeginRow",""));
               ColumnName Col;
               for(int i = 1; i <= nColNum; i++)
               {
                        // �ָ�ÿ����
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
               //��Ӵ��ĸ��ļ���ȡ������
               Col.sName = "�ļ���";
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
                MessageBox(Application->Handle,"��ȡ�����ļ�test.iniʧ��!","��Ϣ��ʾ!",MB_OK+MB_ICONINFORMATION+MB_SYSTEMMODAL);
                delete ini;
        }
        suiProgressBar1->Position = 0;
        sShow += "����\t\t";
        if(sTitle.IsEmpty()) sShow += "��\n";
        else sShow += sTitle + "\n";
        sShow += "����\t\t";
        for(int i = 0; i < vCol.size(); i++)
        {
                sShow += vCol[i].sName + "\t";
        }
        sShow+="\n�г�\t\t";
        for(int i = 0; i <vCol.size(); i++)
        {
                sShow += IntToStr(vCol[i].nLen) + "\t";
        }
        sShow += "\n��ȡ�ֶ�\t";
        for(int i = 0; i <vCol.size(); i++)
        {
                if(i < vCol.size() -1)
                        sShow += IntToStr(vCol[i].ReadCol) + "\t";
                else
                        sShow += "�ļ���";
        }
        sShow += "\n��ȡ��\t��" + IntToStr(nBeginRow) + "�п�ʼ";

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
                ShowMessage("���� Excel ����, ������û�а�װExcel.");
                vExcelAppOpen = Unassigned;
                return;
        }

        try
        {
                vExcelAppOpen.PG("WorkBooks").PR("Open", Path.c_str());
                SheetOpen = vExcelAppOpen.PG("ActiveSheet");
         
                String sWjj = Path.SubString(Path.LastDelimiter('\\')+1,Path.LastDelimiter('.')-1);//�ļ�������
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
                String sShowMessage = "�ϲ��ļ�" + Path + " ����";
                ShowMessage(sShowMessage);
                vExcelAppOpen.OleFunction("Quit");
                vExcelAppOpen = Unassigned;
        }
        //�ر�vExcelAppOpen
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
                 ShowMessage("���� Excel ����, ������û�а�װExcel.");
                 vExcelAppSave = Unassigned;
                 return;
        }
        suiProgressBar1->Position = 0;//������
        nCountRow = 0;
        vExcelAppSave.OlePropertyGet("Workbooks").OleFunction("Add", 1); // ������
        SheetSave = vExcelAppSave.PG("ActiveSheet");
        nCountRow = 1;
        char Buffer[10];
        if(bNeedTitle)
        {
                String strRang = "";       //�ϲ���Ԫ��
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
        //���ú�����������
        for(int i = 0; i <vCol.size()-1; i++)
        {
                SheetSave.PG("Cells",nCountRow, i+1).PS("Value", vCol[i].sName.c_str());
                if(vCol[i].bNumber)
                        vExcelAppSave.PG("Columns",i+1).PS("NumberFormat","@");
                vExcelAppSave.PG("Cells",nCountRow, i+1).PS("ColumnWidth", StrToInt(vCol[i].nLen));    //�����п�
        }
        SheetSave.PG("Cells",nCountRow, nColNum).PS("Value", "�ļ���");
        vExcelAppSave.PG("Cells",nCountRow, nColNum).PS("ColumnWidth", 40);   
        nCountRow ++;
        String path = suiEdit1->Text;
        TSearchRec sr;
        int nExcelNum = 0;
        if (FindFirst(path + "\\*.xls", faAnyFile, sr) == 0)      //��ѯĿ¼�µ�excel��Ŀ
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
        if (FindFirst(path + "\\*.xls", faAnyFile, sr) == 0)      //��ѯĿ¼�µ�excel
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
        strRange = "A"+IntToStr(1)+":"+String::StringOfChar(Buffer[0],1)+IntToStr(nCountRow-1); //��ȡ������Χ    �� A�� - ��F��   ��1�� - ��count +2 ��
        ERange = vExcelAppSave.OlePropertyGet("Range",strRange.c_str());
        EBorders = ERange.OlePropertyGet("Borders");
        EBorders.OlePropertySet("linestyle",1); //����
        EBorders.OlePropertySet("weight",2);    //��ϸ ֵ<=5
        EBorders.OlePropertySet("colorindex",0);
        String strXlsFile = suiEdit2->Text ;
        // �������Excel�ļ�
        vExcelAppSave.OlePropertyGet("ActiveWorkbook")
        .OleFunction("SaveAs", strXlsFile.c_str());
        vExcelAppSave.OleFunction("Quit");
        vExcelAppSave = Unassigned;
        suiProgressBar1->Position = 100;
        //lstFile->SaveToFile(suiEdit1->Text.c_str());
        ShowMessage("�ϲ����");
}
//---------------------------------------------------------------------------
