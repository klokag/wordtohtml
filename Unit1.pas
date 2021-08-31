unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, ComObj, Vcl.OleCtrls,
  SHDocVw;


type
  TForm1 = class(TForm)
    OpenFile_B1: TButton;
    FileName1: TEdit;
    OD: TOpenDialog;
    CreateHTML_B: TButton;
    PreView_WB: TWebBrowser;
    Memo1: TMemo;
    procedure FormCreate(Sender: TObject);
    procedure OpenFile_B1Click(Sender: TObject);
    procedure CreateHTML_BClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
    // ñîçäàíèå html äîêóìåíòà
    procedure CreateHTMLFile;
    procedure WordParsing;
    procedure ParagraphParsing(paragraph: variant);
    procedure TableFormatting();
  end;

var
  Form: TForm1;
  // html-ôàéë
  HTMLFile: TStringList;
  // word-ôàéë
  W: variant;
  // ñ÷åò÷èê òàáëèöû
  TableCount: integer;

implementation

{$R *.dfm}

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  W.activedocument.close;
  W.quit;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  // çàäàåì íà÷àëüíóþ äèðåêòîðèþ äëÿ OpenDialog
  OD.InitialDir := ExtractFilePath(Application.ExeName);
  // ñîçäàåì äîêóìåíò, html
  HTMLFile := TStringList.Create;
end;

procedure TForm1.OpenFile_B1Click(Sender: TObject);
begin
  // åñëè ôàéë íå âûáðàí, òî âûõîäèì
  { if (not OD.Execute) or (OD.FileName = '') then
    exit; }

  // çàãðóæàåì òåñòîâûé äîêóìåíò
  W := CreateOleObject('Word.Application');
  // W.Documents.Open(OD.FileName, EmptyParam, True);
  W.Documents.Open('C:\Users\Public\Documents\Ïðàêòèêà\test doc.docx',
    EmptyParam, True);
  W.activedocument.SaveAs
    ('C:\Users\Public\Documents\Ïðàêòèêà\test doc_dublicate.docx');
  W.activedocument.close;
  W.Documents.Open
    ('C:\Users\Public\Documents\Ïðàêòèêà\test doc_dublicate.docx');
  // W.Visible := True;
  // 'C:\Users\Public\Documents\Ïðàêòèêà\Win32\Debug\test doc.docx'
  // ïèøåì ïóòü
  // FileName1.Text := OD.FileName;
  FileName1.Text :=
    'C:\Users\Public\Documents\Ïðàêòèêà\Win32\Debug\test doc.docx';
end;

procedure TForm1.CreateHTML_BClick(Sender: TObject);
begin
  // ñîçäàåì html-ôàéë
  CreateHTMLFile;
  // ñîõðàíÿåì åãî
  HTMLFile.SaveToFile('C:\Users\Public\Documents\Ïðàêòèêà\test.html');
  // îòêðûâàåì â áðàóçåðå
  PreView_WB.Navigate('file://' +
    'C:\Users\Public\Documents\Ïðàêòèêà\test.html');
end;

procedure TForm1.CreateHTMLFile;
begin
  // î÷èùàåì
  HTMLFile.Clear;
  // îáðàùàåìñÿ ê "html"
  with HTMLFile do
  begin
    // ïèøåì çàãîëîâîê
    Add('<html>');
    Add('<head>');
    Add('<title>' + ExtractFileName(FileName1.Text) + '</title>' + #10#13 +
      '</head>');
    Add('<body>');
    // íàïîëíÿåì body
    WordParsing;
    Memo1.lines.Add('thats all');
    Add('</body>');
    Add('</html>');
  end;
end;

procedure TForm1.WordParsing;
var
  i: integer;
  wordrange: variant;
  FontName, curFontName, AlignName: string;
  listFlag: boolean;
  CurTable, AlignNumb, curFontSize, FontSize: integer;
begin
  listFlag := False; // ôëàã äëÿ ïðîâåðêè ñïèñêîâ
  TableCount := 1; // èíèöèàëèçèðóåì ñ÷åò÷èê òàáëèö
  CurTable := 1; // òåêóùàÿ òàáëèöà

  // öèêë ïî àáçàöàì
  for i := 1 to W.activedocument.Paragraphs.Count do
  begin
    wordrange := W.activedocument.Paragraphs.Item(i).range; // àáçàö

    AlignNumb := W.activedocument.Paragraphs.Item(i).Alignment;
    case AlignNumb of
    0:
    AlignName := 'Left';
    1:
    AlignName := 'Center';
    2:
    AlignName := 'Right';
    3:
    AlignName := 'Justify'
    end;

    wordrange := W.activedocument.Paragraphs.Item(i).range; // àáçàö

    curFontName := string('face = "' + string(wordrange.formattedText.Font.Name)
      + '"'); // íàçâàíèå øðèôòà
    curFontSize := strTOInt(varToStr(wordrange.formattedText.Font.Size)); // íàçâàíèå øðèôòà

    case curFontSize of
    12:
    curFontSize := 3;
    14:
    curFontSize := 4;
    18:
     curFontSize := 5;
    24:
     curFontSize := 6;
    end;

    // åñëè òåêñò æèðíûé, ïåðåäàòü àáçàö â ôóíêöèþ
    if wordrange.formattedText.bold <> 0 then
      ParagraphParsing(wordrange);

    // åñëè òåêóùèé øðèôò íå ñîâïàäàåò ñ ïðåäûäóùèì
    if (curFontName <> FontName) or (curFontSize <> FontSize) then
    begin
      wordrange.insertbefore('<font ' + curFontName + '" size = "' +
        intTOstr(curFontSize) + '">');
      if i <> 1 then
        wordrange.insertbefore('</font>');
      FontName := curFontName;
      FontSize := curFontSize;
    end;

    //åñëè íûíÿøíÿÿ òàáëèöà çàêîí÷èëàñü, ìåíÿåì CurTable
    if (CurTable <> TableCount) and (wordrange.Tables.Count = 0) then
    begin
      CurTable := CurTable + 1;
    end;

    //åñëè ïàðàãðàô ïðèíàäëåæèò òàáëèöå
    if (wordrange.Tables.Count > 0) then
    begin
      //åñëè òåêóùàÿ òàáëèöà ñîâïàäàåò ñ TableCount
      if CurTable = TableCount then
        TableFormatting();
      continue;
    end;

    // åñëè àáçàö íàõîäèòñÿ â ñïèñêå
    if (wordrange.listformat.listtype > 0) and
      (wordrange.listformat.listtype < 6) then
    begin
      if listFlag = False then
        HTMLFile.Append('<ul>');
      HTMLFile.Append('<li>' + string(wordrange.Text) + '</li>');
      listFlag := True;
      continue;
    end;
    if listFlag = True then
    begin
      HTMLFile.Append('</ul>');
      listFlag := False;
    end;


    HTMLFile.Append('<p ALIGN = "' + AlignName + '">' + string(wordrange.Text) + '</p>');

  end;
  HTMLFile.Append('</font>');
end;

procedure TForm1.ParagraphParsing(paragraph: variant);
var
  Flag: boolean;
  i: integer;
  isBold: integer;
begin
  Memo1.lines.Add('gogo');
  // îáðàòíûé öèêë ïî ñëîâàì â ïàðàãðàôå
  for i := paragraph.words.Count downto 1 do
  begin
    isBold := paragraph.words.Item(i).formattedText.bold;
    // ïðîâåðêà æèðíîñòè ñëîâà
    // åñëè òåêñò æèðíûé
    if isBold = -1 then
    begin
      if Flag = False then
        paragraph.words.Item(i).insertafter('</b>');
      Flag := True;
    end;
    // åñëè òåêñò íå æèðíûé
    if (isBold = 0) and (Flag = True) then
    begin
      paragraph.words.Item(i).insertafter('<b>');
      Flag := False;
    end;
  end;
  // ïðîñòàâëÿåì îòêðûâàþùèé òåã, åñëè ïåðâîå ñëîâî æèðíîå
  if Flag = True then
    paragraph.words.Item(1).insertbefore('<b>');
end;

procedure TForm1.TableFormatting();
var
  TableColumnsCount, TableRowsCount, CurrentRow, CurrentColumn: integer;
  text: string;
begin
  // Îïðåäåëÿåì êîëè÷åñòâî ñòîëáöîâ
  TableColumnsCount := W.activedocument.Tables.Item(TableCount).Columns.Count;
  // Îïðåäåëÿåì êîëè÷åñòâî ñòðîê
  TableRowsCount := W.activedocument.Tables.Item(TableCount).Rows.Count;
  HTMLFile.Append('<table border="4" bordercolor="#000000">');
  //ïåðåáèðàåì ïî ñòðîêàì/êîëîíêàì
  for CurrentRow := 1 to TableRowsCount do
  begin
    HTMLFile.Append('<tr>');
    for CurrentColumn := 1 to TableColumnsCount do
    begin
      text := W.activedocument.Tables.Item(TableCount)
        .Cell(CurrentRow, CurrentColumn).range.Text;
        //memo1.Lines.Add(intToStr(length(text)));

      if length(text) > 2 then
      begin
         HTMLFile.Append('<th>' + Copy(text, 1,Length(text) - 1) + '</th>');
      end;
      if length(text) = 2 then HTMLFile.Append('<th>' + '&nbsp;' +'</th>');

    end;
    HTMLFile.Append('</tr>');
  end;

  HTMLFile.Append('</table>');

  TableCount := TableCount + 1;
end;

end.

{
  Íàäî ñäåëàòü:
  1. îãëàâëåíèå
  2. òèïû ñïèñêîâ
  3. òàáëèöû
  4. êàðòèíêè
  5. êîñìåòèêà (âûðàâíèâàíèå)
}

{
hyperlinks(item)
.follow
.range.text
.add anchor
}
