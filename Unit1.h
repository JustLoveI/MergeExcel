//---------------------------------------------------------------------------

#ifndef Unit1H
#define Unit1H
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "SUIButton.hpp"
#include <Dialogs.hpp>
#include "SUIEdit.hpp"
#include "SUIProgressBar.hpp"
#include <ExtCtrls.hpp>
#include <Graphics.hpp>
//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
        TsuiButton *suiButton2;
        TOpenDialog *OpenDialog1;
        TLabel *Label1;
        TsuiEdit *suiEdit1;
        TLabel *Label2;
        TsuiEdit *suiEdit2;
        TSaveDialog *dlgSave1;
        TsuiButton *suiButton1;
        TsuiProgressBar *suiProgressBar1;
        TLabel *Label3;
        TsuiButton *suiButton3;
        TLabel *Label4;
        void __fastcall suiButton2Click(TObject *Sender);
        void __fastcall suiButton1Click(TObject *Sender);
        void __fastcall suiButton3Click(TObject *Sender);
private:	// User declarations
public:		// User declarations
        __fastcall TForm1(TComponent* Owner);
        void __fastcall ExportExcel(String Path);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
 