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
    // �������� html ���������
    procedure CreateHTMLFile;
    procedure WordParsing;
    procedure ParagraphParsing(paragraph: variant);
    procedure TableFormatting();
  end;

var
  Form: TForm1;
  // html-����
  HTMLFile: TStringList;
  // word-����
  W: variant;
  // ������� �������
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
  // ������ ��������� ���������� ��� OpenDialog
  OD.InitialDir := ExtractFilePath(Application.ExeName);
  // ������� ��������, html
  HTMLFile := TStringList.Create;
end;

procedure TForm1.OpenFile_B1Click(Sender: TObject);
begin
  // ���� ���� �� ������, �� �������
  { if (not OD.Execute) or (OD.FileName = '') then
    exit; }

  // ��������� �������� ��������
  W := CreateOleObject('Word.Application');
  // W.Documents.Open(OD.FileName, EmptyParam, True);
  W.Documents.Open('C:\Users\Public\Documents\��������\test doc.docx',
    EmptyParam, True);
  W.activedocument.SaveAs
    ('C:\Users\Public\Documents\��������\test doc_dublicate.docx');
  W.activedocument.close;
  W.Documents.Open
    ('C:\Users\Public\Documents\��������\test doc_dublicate.docx');
  // W.Visible := True;
  // 'C:\Users\Public\Documents\��������\Win32\Debug\test doc.docx'
  // ����� ����
  // FileName1.Text := OD.FileName;
  FileName1.Text :=
    'C:\Users\Public\Documents\��������\Win32\Debug\test doc.docx';
end;

procedure TForm1.CreateHTML_BClick(Sender: TObject);
begin
  // ������� html-����
  CreateHTMLFile;
  // ��������� ���
  HTMLFile.SaveToFile('C:\Users\Public\Documents\��������\test.html');
  // ��������� � ��������
  PreView_WB.Navigate('file://' +
    'C:\Users\Public\Documents\��������\test.html');
end;

procedure TForm1.CreateHTMLFile;
begin
  // �������
  HTMLFile.Clear;
  // ���������� � "html"
  with HTMLFile do
  begin
    // ����� ���������
    Add('<html>');
    Add('<head>');
    Add('<title>' + ExtractFileName(FileName1.Text) + '</title>' + #10#13 +
      '</head>');
    Add('<body>');
    // ��������� body
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
  listFlag := False; // ���� ��� �������� �������
  TableCount := 1; // �������������� ������� ������
  CurTable := 1; // ������� �������

  // ���� �� �������
  for i := 1 to W.activedocument.Paragraphs.Count do
  begin
    wordrange := W.activedocument.Paragraphs.Item(i).range; // �����

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

    wordrange := W.activedocument.Paragraphs.Item(i).range; // �����

    curFontName := string('face = "' + string(wordrange.formattedText.Font.Name)
      + '"'); // �������� ������
    curFontSize := strTOInt(varToStr(wordrange.formattedText.Font.Size)); // �������� ������

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

    // ���� ����� ������, �������� ����� � �������
    if wordrange.formattedText.bold <> 0 then
      ParagraphParsing(wordrange);

    // ���� ������� ����� �� ��������� � ����������
    if (curFontName <> FontName) or (curFontSize <> FontSize) then
    begin
      wordrange.insertbefore('<font ' + curFontName + '" size = "' +
        intTOstr(curFontSize) + '">');
      if i <> 1 then
        wordrange.insertbefore('</font>');
      FontName := curFontName;
      FontSize := curFontSize;
    end;

    //���� �������� ������� �����������, ������ CurTable
    if (CurTable <> TableCount) and (wordrange.Tables.Count = 0) then
    begin
      CurTable := CurTable + 1;
    end;

    //���� �������� ����������� �������
    if (wordrange.Tables.Count > 0) then
    begin
      //���� ������� ������� ��������� � TableCount
      if CurTable = TableCount then
        TableFormatting();
      continue;
    end;

    // ���� ����� ��������� � ������
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
  // �������� ���� �� ������ � ���������
  for i := paragraph.words.Count downto 1 do
  begin
    isBold := paragraph.words.Item(i).formattedText.bold;
    // �������� �������� �����
    // ���� ����� ������
    if isBold = -1 then
    begin
      if Flag = False then
        paragraph.words.Item(i).insertafter('</b>');
      Flag := True;
    end;
    // ���� ����� �� ������
    if (isBold = 0) and (Flag = True) then
    begin
      paragraph.words.Item(i).insertafter('<b>');
      Flag := False;
    end;
  end;
  // ����������� ����������� ���, ���� ������ ����� ������
  if Flag = True then
    paragraph.words.Item(1).insertbefore('<b>');
end;

procedure TForm1.TableFormatting();
var
  TableColumnsCount, TableRowsCount, CurrentRow, CurrentColumn: integer;
  text: string;
begin
  // ���������� ���������� ��������
  TableColumnsCount := W.activedocument.Tables.Item(TableCount).Columns.Count;
  // ���������� ���������� �����
  TableRowsCount := W.activedocument.Tables.Item(TableCount).Rows.Count;
  HTMLFile.Append('<table border="4" bordercolor="#000000">');
  //���������� �� �������/��������
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
  ���� �������:
  1. ����������
  2. ���� �������
  3. �������
  4. ��������
  5. ��������� (������������)
}

{
hyperlinks(item)
.follow
.range.text
.add anchor
}
