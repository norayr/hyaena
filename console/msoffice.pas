unit msoffice;

interface
Uses ComObj, ActiveX, SysUtils;
  type
  TWordReplaceFlags = set of (wrfReplaceAll, wrfMatchCase, wrfMatchWildcards);
  function MSWordIsInstalled: Boolean;
  function Word_StringReplace(ADocument, ODocument: TFileName; SearchString, ReplaceString: Widestring; Flags: TWordReplaceFlags): Boolean;
  function Word_CharsReplace(ADocument, ODocument: TFileName; Search : Array of String; Replace : Array of WideString; Flags: TWordReplaceFlags): Boolean;
implementation
   function Word_StringReplace(ADocument, ODocument: TFileName; SearchString, ReplaceString: Widestring; Flags: TWordReplaceFlags): Boolean;
const
  wdFindContinue = 1;
  wdReplaceOne = 1;
  wdReplaceAll = 2;
  wdDoNotSaveChanges = 0;
var
  WordApp: OLEvariant;
  overwrite : boolean;
begin
  Result := False;
  if Odocument = '' then overwrite := true else overwrite := false;
  OleCheck (CoInitialize (nil));
  { Check if file exists }
  if not FileExists(ADocument) then
  begin
    WriteLn('Specified Document not found.');
    Exit;
  end;

  { Create the OLE Object }
  try
    WordApp := ComObj.CreateOLEObject('Word.Application');
  except
    on E: Exception do
    begin
      E.Message := 'Word is not available.';
      raise;
    end;
  end;

  try
    { Hide Word }
    WordApp.Visible := True;
    { Open the document }
    WordApp.Documents.Open(ADocument);
    { Initialize parameters}
    WordApp.Selection.Find.ClearFormatting;
    WordApp.Selection.Find.Text := SearchString;
    WordApp.Selection.Find.Replacement.Text := ReplaceString;
    WordApp.Selection.Find.Forward := True;
    WordApp.Selection.Find.Wrap := wdFindContinue;
    WordApp.Selection.Find.Format := False;
    WordApp.Selection.Find.MatchCase := wrfMatchCase in Flags;
    WordApp.Selection.Find.MatchWholeWord := False;
    WordApp.Selection.Find.MatchWildcards := wrfMatchWildcards in Flags;
    WordApp.Selection.Find.MatchSoundsLike := False;
    WordApp.Selection.Find.MatchAllWordForms := False;
    { Perform the search}
    if wrfReplaceAll in Flags then
      WordApp.Selection.Find.Execute(Replace := wdReplaceAll)    else
      WordApp.Selection.Find.Execute(Replace := wdReplaceOne);
    { Save word }
    if Overwrite then WordApp.ActiveDocument.SaveAs(ADocument)
    else WordApp.ActiveDocument.SaveAs(ODocument)
    ;
    { Assume that successful }
    Result := True;
    { Close the document }
    WordApp.ActiveDocument.Close(wdDoNotSaveChanges);
  finally
    { Quit Word }
    WordApp.Quit;
    //WordApp := Unassigned;
    CoUninitialize;
  end;
end;

function Word_CharsReplace(ADocument, ODocument: TFileName; Search : Array of String; Replace : Array of WideString; Flags: TWordReplaceFlags): Boolean;
const
  wdFindContinue = 1;
  wdReplaceOne = 1;
  wdReplaceAll = 2;
  wdDoNotSaveChanges = 0;
var
  WordApp: OLEvariant;
  overwrite : boolean;
  i : longint;
 function wordrunning : boolean;
   begin
   {check if word already running}
   try
    // searching for an opened word instance, if no, then exception
    WordApp := GetActiveOleObject('Word.Application');
    wordrunning := true;
    exit;
    //  visible
    //WordApp.Visible := true;
  except
  {word is not opened at all}
  wordrunning := false;
  end;
end {wordrunning};

function opened (s : TFileName) : boolean;
var i : integer;
begin
opened := false;
 // searching for our document
    for i := 1 to WordApp.Documents.Count do begin

     if WordApp.Documents.Item(i).Name = SysUtils.ExtractFileName(Adocument) then begin
         WordApp.Documents.Item(i).Activate;
         opened := true;
         exit;
     end {if} ;
  end;
end {activate};

procedure startit;
begin
  { Create the OLE Object }
  try
    WordApp := ComObj.CreateOLEObject('Word.Application');
  except
    on E: Exception do
    begin
      E.Message := 'MS Office is not available.';
      raise;
    end;
  end;

end;
procedure open ;
begin
    { Open the document }

    WordApp.Documents.Open(Adocument);


end;
begin { start of procedure}
  Result := False;
  if Odocument = '' then overwrite := true else overwrite := false;
  OleCheck (CoInitialize (nil));
  { Check if file exists }
  if not FileExists(ADocument) then
  begin
    WriteLn('Specified Document not found.');
    Exit;
  end;
  try
   //check if word is running
   if not wordrunning then begin
        startit;
        open;
        end
       else
        begin
        if not opened(Adocument) then open;
        end;

    { Hide Word }
    WordApp.Visible := False;
    { Initialize parameters}
    i := 0;
    repeat
        WordApp.Selection.Find.ClearFormatting;
    WordApp.Selection.Find.Text := Search[i];
    WordApp.Selection.Find.Replacement.Text := Replace[i];
      WordApp.Selection.Find.Forward := True;
    WordApp.Selection.Find.Wrap := wdFindContinue;
      WordApp.Selection.Find.Format := False;
    WordApp.Selection.Find.MatchCase := wrfMatchCase in Flags;
    WordApp.Selection.Find.MatchWholeWord := False;
    WordApp.Selection.Find.MatchWildcards := wrfMatchWildcards in Flags;
    WordApp.Selection.Find.MatchSoundsLike := False;
    WordApp.Selection.Find.MatchAllWordForms := False;
    { Perform the search}
    if wrfReplaceAll in Flags then
      WordApp.Selection.Find.Execute(Replace := wdReplaceAll)
    else
      WordApp.Selection.Find.Execute(Replace := wdReplaceOne);
      inc(i);

     until i = 95;
    { Save word }

    if Overwrite then WordApp.ActiveDocument.SaveAs(ADocument)
    else WordApp.ActiveDocument.SaveAs(ODocument)
    ;

    { Assume that successful }
    Result := True;
    { Close the document }
    WordApp.ActiveDocument.Close(wdDoNotSaveChanges);
  finally
    { Quit Word }
    WordApp.Quit;
    //WordApp := Unassigned;
    CoUninitialize;
  end;
end;

function AppIsInstalled(strOLEObject: string): Boolean;
var
  ClassID: TCLSID;
begin
  Result := (CLSIDFromProgID(PWideChar(WideString(strOLEObject)), ClassID) = S_OK)
end;

   function MSWordIsInstalled: Boolean;
begin
  Result := AppIsInstalled('Word.Application');
end;




end.
