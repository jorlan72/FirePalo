unit main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.WinXPanels, Vcl.ComCtrls,
  Vcl.ExtCtrls, Vcl.StdCtrls, StrUtils, System.Generics.Collections,
  Vcl.Imaging.pngimage, Vcl.Grids, Vcl.ValEdit, Vcl.Imaging.jpeg, ShellAPI,
  Vcl.Mask, Vcl.Menus, System.ImageList, Vcl.ImgList;

type
  TfrmMain = class(TForm)
    StatusBar: TStatusBar;
    PageControl: TPageControl;
    TabSheetObjects: TTabSheet;
    TabSheetRules: TTabSheet;
    TabSheetZones: TTabSheet;
    ButtonOpenFPFile: TButton;
    ButtonManageableObjects: TButton;
    GroupBoxNetworkObjects: TGroupBox;
    SplitterServices: TSplitter;
    GroupBoxServices: TGroupBox;
    GroupBoxRules: TGroupBox;
    TabSheetPaloAltoCLI: TTabSheet;
    GroupBoxPaloCLI: TGroupBox;
    MemoPaloCLI: TMemo;
    TabSheetWelcome: TTabSheet;
    GroupBoxNetworkGroups: TGroupBox;
    SplitterObjects2: TSplitter;
    GroupBoxSecurityZones: TGroupBox;
    TabSheetServices: TTabSheet;
    GroupBoxServiceGroups: TGroupBox;
    TabSheetURL: TTabSheet;
    GroupBoxURLGroups: TGroupBox;
    SplitterURLs: TSplitter;
    GroupBoxURL: TGroupBox;
    TabSheetApplications: TTabSheet;
    GroupBoxApplications: TGroupBox;
    ImageLogo: TImage;
    LabelApp: TLabel;
    LabelCopyRight: TLabel;
    TabSheetFPCLI: TTabSheet;
    GroupBoxPolicies: TGroupBox;
    TreeViewPolicies: TTreeView;
    GroupBoxWhatToDoNext: TGroupBox;
    LabelNote02: TLabel;
    LabelNote01: TLabel;
    LabelNote03: TLabel;
    LabelNote04: TLabel;
    LabelConvert: TLabel;
    RuleNames: TValueListEditor;
    Zones: TValueListEditor;
    TabSheetManageableObjects: TTabSheet;
    ProgressBarProsessen: TProgressBar;
    GroupBoxProsessen: TGroupBox;
    LabelParsed01: TLabel;
    LabelParsed02: TLabel;
    LabelParsed03: TLabel;
    Applications: TValueListEditor;
    URLEntriesGroup: TValueListEditor;
    URLEntries: TValueListEditor;
    NetworkObjectsGroup: TValueListEditor;
    NetworkObjects: TValueListEditor;
    ServicesGroup: TValueListEditor;
    Services: TValueListEditor;
    TabSheetActions: TTabSheet;
    Action: TValueListEditor;
    LogEnd: TValueListEditor;
    LogBeginning: TValueListEditor;
    Category: TValueListEditor;
    UserGroup: TValueListEditor;
    GroupBoxAction: TGroupBox;
    LabelFPDescription: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    PanelZones: TPanel;
    LabelZone01: TLabel;
    LabelZone02: TLabel;
    LabelZone03: TLabel;
    LabelZone04: TLabel;
    LabelZone06: TLabel;
    LabelZone05: TLabel;
    PanelSecurityRules: TPanel;
    LabelRule01: TLabel;
    LabelRule02: TLabel;
    LabelRule03: TLabel;
    LabelRule04: TLabel;
    PanelApplications: TPanel;
    LabelApp01: TLabel;
    LabelApp02: TLabel;
    LabelApp03: TLabel;
    LabelApp04: TLabel;
    PanelGroupAddresses: TPanel;
    PanelGroupServices: TPanel;
    PanelGroupURL: TPanel;
    PanelAddresses: TPanel;
    LabelAdr01: TLabel;
    LabelAdr02: TLabel;
    LabelAdr03: TLabel;
    LabelAdr04: TLabel;
    LabelAdr05: TLabel;
    LabelAdr6: TLabel;
    PanelServices: TPanel;
    LabelService01: TLabel;
    LabelService02: TLabel;
    LabelService03: TLabel;
    LabelService04: TLabel;
    LabelService06: TLabel;
    LabelService05: TLabel;
    CheckBoxRuleDescriptionNumber: TCheckBox;
    PanelURLs: TPanel;
    LabelURL01: TLabel;
    LabelURL02: TLabel;
    LabelURL03: TLabel;
    LabelURL04: TLabel;
    LabelURL05: TLabel;
    LabelURL06: TLabel;
    Panel1: TPanel;
    LabelAction01: TLabel;
    LabelAction02: TLabel;
    LabelAction03: TLabel;
    ComboBoxAPPID: TComboBox;
    GroupBoxPaloAppID: TGroupBox;
    PanelGroupApplications: TPanel;
    TabSheetExtraOptions: TTabSheet;
    LabelApplipedia: TLabel;
    LabelCiscoAPPID: TLabel;
    ButtonGenerate: TButton;
    Panel2: TPanel;
    LabelOptionsHeader: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    CheckBoxUserID: TCheckBox;
    CheckBoxCategory: TCheckBox;
    LabelDone01: TLabel;
    LabelDone02: TLabel;
    GroupBoxPanorama: TGroupBox;
    EditPanorama: TEdit;
    GroupBoxCategory: TGroupBox;
    GroupBoxUserID: TGroupBox;
    GroupBoxOutputConfig: TGroupBox;
    CheckBoxAllLowerCase: TCheckBox;
    Label6: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    TreeViewCommit: TTreeView;
    ButtonUpdateTree: TButton;
    LabelDone03: TLabel;
    ProgressBarUpdate: TProgressBar;
    Label3: TLabel;
    Label7: TLabel;
    Label11: TLabel;
    LabeledEditProfileGroup: TLabeledEdit;
    LabeledEditLogSetting: TLabeledEdit;
    Label12: TLabel;
    CheckBoxLogBeginning: TCheckBox;
    CheckBoxLogEnd: TCheckBox;
    CheckBoxUseApplicationDefault: TCheckBox;
    PopupMenuValue: TPopupMenu;
    MenuShred: TMenuItem;
    MenuDelete: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    MenuLowercaseAllValues: TMenuItem;
    MenuUppercaseAllValues: TMenuItem;
    ImageList: TImageList;
    LabelAPP05: TLabel;
    NetworkObjectsValues: TValueListEditor;
    ServicesValues: TValueListEditor;
    URLEntriesValues: TValueListEditor;
    Label13: TLabel;
    N5: TMenuItem;
    MenuExportList: TMenuItem;
    MenuImportList: TMenuItem;
    Label14: TLabel;
    N6: TMenuItem;
    MenuAddTag: TMenuItem;
    N1: TMenuItem;
    MenuConvertToApplication: TMenuItem;
    LabelRule07: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure ButtonOpenFPFileClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ButtonManageableObjectsClick(Sender: TObject);
    procedure LabelCiscoAPPIDClick(Sender: TObject);
    procedure LabelApplipediaClick(Sender: TObject);
    procedure ButtonUpdateTreeClick(Sender: TObject);
    procedure ButtonGenerateClick(Sender: TObject);
    procedure PopupMenuValuePopup(Sender: TObject);
    procedure ServicesDblClick(Sender: TObject);
    procedure ServicesValuesDblClick(Sender: TObject);
    procedure NetworkObjectsDblClick(Sender: TObject);
    procedure NetworkObjectsValuesDblClick(Sender: TObject);
    procedure URLEntriesDblClick(Sender: TObject);
    procedure URLEntriesValuesDblClick(Sender: TObject);
    procedure MenuExportListClick(Sender: TObject);
    procedure MenuImportListClick(Sender: TObject);
    procedure MenuDeleteClick(Sender: TObject);
    procedure MenuAddTagClick(Sender: TObject);
    procedure MenuLowercaseAllValuesClick(Sender: TObject);
    procedure MenuUppercaseAllValuesClick(Sender: TObject);
    procedure MenuShredClick(Sender: TObject);
    procedure MenuConvertToApplicationClick(Sender: TObject);
    procedure NetworkObjectsGroupDblClick(Sender: TObject);
    procedure ServicesGroupDblClick(Sender: TObject);
    procedure URLEntriesGroupDblClick(Sender: TObject);
    procedure RuleNamesMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

Procedure ParseOutput;

var
  frmMain: TfrmMain;
  ParseText : TStringlist;
  ActiveParseLine : Integer;
  NextLine : Integer;
  CurrentRule : String;
  RuleNode : TTreeNode;
  CommandString : String;
  ClickedEditor : TValueListEditor;
  ClickedEditor2 : TValueListEditor;
  TagList : TStringList;
  GroupList : TStringList;

function FindFirstNonSpace(const S: string): Integer;
function ContainsHyphenBetweenParentheses(const s: string): Boolean;
function GetApplicationVersion: string;
Function ProcessParameter(Parameter : string) : string;
Function ProcessRule(Rule : string) : string;
procedure PalofySourceTree;
procedure SplitCommaObjects;
function FindChildNode(Node: TTreeNode; const Text: string): TTreeNode;
procedure UpdateResultTree;
procedure FillAction;
function FindNodeWithText(StartNode: TTreeNode; const Text: string): TTreeNode;
procedure FillSourceZones;
procedure FillDestinationZones;
procedure FillSourceNetworks;
procedure FillDestinationNetworks;
procedure FillDestinationPorts;
procedure FillSourcePorts;
procedure FillApplications;
procedure FillURLs;
procedure FillCategory;
procedure FillLogBeginning;
procedure FillLogEnd;
procedure FillUserGroup;
procedure LowercaseTreeView(Tree: TTreeView);
procedure AddAnyChildIfEmpty(Tree: TTreeView);
procedure AddUrlCategoryToTag(Tree: TTreeView);
procedure AddUserGroupToTag(Tree: TTreeView);
procedure AddDescriptionToRootNodes(Tree: TTreeView);
procedure GenerateCLIAction(rootNodeName, foundChildName: string);
procedure GenerateCLISourceZones(rootNodeName, foundChildName: string);
procedure GenerateCLIDestinationZones(rootNodeName, foundChildName: string);
procedure GenerateCLISourceNetworks(rootNodeName, foundChildName: string);
procedure GenerateCLIDestinationNetworks(rootNodeName, foundChildName: string);
procedure GenerateCLISourcePorts(rootNodeName, foundChildName: string);
procedure GenerateCLIDestinationPorts(rootNodeName, foundChildName: string);
procedure GenerateCLIApplications(rootNodeName, foundChildName: string);
procedure GenerateCLILogBeginning(rootNodeName, foundChildName: string);
procedure GenerateCLILogEnd(rootNodeName, foundChildName: string);
procedure GenerateCLIURLEntries(rootNodeName, foundChildName: string);
procedure GenerateCLITag(rootNodeName, foundChildName: string);
procedure GenerateCLIDescription(rootNodeName, foundChildName: string);
procedure GenerateCLICommands(Tree: TTreeView);
procedure ReplaceInValue(ValueListEditor: TValueListEditor; const SearchWord, ReplaceWord: string);
Procedure ExtractNetworkObjectsValues;
Procedure ExtractServiceValues;
Procedure ExtractURLValues;
procedure UpdateValuesFromComboBox(ValueList: TValueListEditor; ComboBox: TComboBox);
procedure DeleteSelectedItemAndRelatedNodes;
procedure AddTagToRootNode(const MatchString, TagAddString: string);
procedure CopyTagChildrenToCommit;
procedure ModifyValueListEditorValues(ValueListEditor: TValueListEditor; Operator: string);
procedure CreateAddresses;
procedure CreateServices;
procedure CreateUrls;
procedure CreateTags;
procedure TrimToMax(ValueListEditor: TValueListEditor);
procedure UpdateGroupsList(SearchString: string; GroupList: TStringList; Addresses: TValueListEditor);
procedure CreateAddressGroups;
procedure CreateServiceGroups;
procedure CreateURLGroups;
procedure MoveMatchingItemsToApplications(TreeView: TTreeView; const SearchWord: string);

implementation

{$R *.dfm}

function GetApplicationVersion: string;
// GET THE VERSION BUILD
var
  Exe: string;
  Size, Handle: DWORD;
  Buffer: TBytes;
  FixedPtr: PVSFixedFileInfo;
begin
  Exe := ParamStr(0);
  Size := GetFileVersionInfoSize(PChar(Exe), Handle);
  if Size = 0 then exit;
  SetLength(Buffer, Size);
  if not GetFileVersionInfo(PChar(Exe), Handle, Size, Buffer) then exit;
  if not VerQueryValue(Buffer, '\', Pointer(FixedPtr), Size) then exit;
  Result := Format('%d.%d.%d.%d', [LongRec(FixedPtr.dwFileVersionMS).Hi, LongRec(FixedPtr.dwFileVersionMS).Lo, LongRec(FixedPtr.dwFileVersionLS).Hi, LongRec(FixedPtr.dwFileVersionLS).Lo])
end;

function ContainsHyphenBetweenParentheses(const s: string): Boolean;
var
  startPos, endPos, hyphenPos: Integer;
begin
  startPos := Pos('(', s);
  endPos := Pos(')', s);
  if (startPos > 0) and (endPos > startPos) then
  begin
    hyphenPos := Pos('-', Copy(s, startPos + 1, endPos - startPos - 1));
    Result := hyphenPos > 0;
  end else Result := False;
end;

function FindFirstNonSpace(const S: string): Integer;
begin
  Result := 1;
  while (Result <= Length(S)) and (S[Result] = ' ') do Inc(Result);
  if Result > Length(S) then Result := 0;
end;

procedure DeleteEmptyLines(Memo: TStringList);
var
  i: Integer;
begin
  for i := Memo.Count - 1 downto 0 do
  begin
    if Trim(Memo.Strings[i]) = '' then Memo.Delete(i);
  end;
end;

function GetParentItemCount(TreeView: TTreeView): Integer;
var
  I: Integer;
begin
  Result := 0;
  for I := 0 to TreeView.Items.Count - 1 do if TreeView.Items[I].Parent = nil then Inc(Result);
end;

procedure TfrmMain.ButtonOpenFPFileClick(Sender: TObject);
var
  OpenFileDialog: TOpenDialog;
begin
  OpenFileDialog := TOpenDialog.Create(nil);
    try
      OpenFileDialog.Title := 'Open FirePower Output';
      OpenFileDialog.Options := [ofFileMustExist, ofPathMustExist];
      if OpenFileDialog.Execute then
        begin
          ParseText.LoadFromFile(OpenFileDialog.FileName);
          if pos('-[ Rule:', ParseText.Text) = 0 then
            begin
              showmessage('File does not seem to contain the correct data. Try another file');
              exit;
            end;
          frmMain.TreeViewPolicies.Items.Clear;
          ParseOutput;
          ParseText.Clear;
          SplitCommaObjects;
          frmMain.ButtonOpenFPFile.Enabled := true;
          frmMain.ButtonManageableObjects.Enabled := true;
          frmMain.PageControl.Pages[1].TabVisible := true;
          frmMain.PageControl.ActivePage := frmMain.TabSheetFPCLI;
          frmMain.GroupBoxPolicies.Caption := ' ' + IntToStr(GetParentItemCount(frmMain.TreeViewPolicies)) + ' policies found ';
          frmMain.ButtonManageableObjects.SetFocus;
        end;
    finally
      OpenFileDialog.Free;
    end;
end;

procedure TfrmMain.ButtonUpdateTreeClick(Sender: TObject);
begin
  frmMain.ButtonUpdateTree.Enabled := false;
  frmMain.TreeViewCommit.Enabled := false;
  frmMain.ProgressBarUpdate.Visible := true;
  frmMain.ProgressBarUpdate.Position := 0;
  frmMain.PageControl.Pages[2].TabVisible := false;
  frmMain.PageControl.Pages[3].TabVisible := false;
  frmMain.PageControl.Pages[4].TabVisible := false;
  frmMain.PageControl.Pages[5].TabVisible := false;
  frmMain.PageControl.Pages[6].TabVisible := false;
  frmMain.PageControl.Pages[7].TabVisible := false;
  frmMain.PageControl.Pages[8].TabVisible := false;
  frmMain.PageControl.Pages[9].TabVisible := false;
  frmMain.PageControl.Pages[10].TabVisible := false;
  frmMain.PageControl.Pages[12].TabVisible := false;
// *** begin to update tree
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating Rulenames and sub sections...';
  UpdateResultTree;
  frmMain.ProgressBarUpdate.Position := 5;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating Actions...';
  FillAction;
  frmMain.ProgressBarUpdate.Position := 10;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating Source Zones...';
  FillSourceZones;
  frmMain.ProgressBarUpdate.Position := 15;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating Destination Zones...';
  FillDestinationZones;
  frmMain.ProgressBarUpdate.Position := 20;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating Source Networks...';
  FillSourceNetworks;
  frmMain.ProgressBarUpdate.Position := 25;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating Destination Networks...';
  FillDestinationNetworks;
  frmMain.ProgressBarUpdate.Position := 30;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating Destination Ports...';
  FillDestinationPorts;
  frmMain.ProgressBarUpdate.Position := 35;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating Source Ports...';
  FillSourcePorts;
  frmMain.ProgressBarUpdate.Position := 40;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating Applications...';
  FillApplications;
  frmMain.ProgressBarUpdate.Position := 45;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating URLs...';
  FillURLs;
  frmMain.ProgressBarUpdate.Position := 50;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating Category...';
  FillCategory;
  frmMain.ProgressBarUpdate.Position := 55;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating Log Beginning...';
  FillLogBeginning;
  frmMain.ProgressBarUpdate.Position := 60;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating Log End...';
  FillLogEnd;
  frmMain.ProgressBarUpdate.Position := 65;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating User-ID...';
  FillUserGroup;
  frmMain.ProgressBarUpdate.Position := 70;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating Tags for Category...';
  if frmMain.CheckBoxCategory.Checked then AddUrlCategoryToTag(frmMain.TreeViewCommit);
  frmMain.ProgressBarUpdate.Position := 75;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating Tags for User-ID...';
  if frmMain.CheckBoxUserID.Checked then AddUserGroupToTag(frmMain.TreeViewCommit);
  frmMain.ProgressBarUpdate.Position := 80;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating Descriptions...';
  if frmMain.CheckBoxRuleDescriptionNumber.Checked then AddDescriptionToRootNodes(frmMain.TreeViewCommit);
  frmMain.ProgressBarUpdate.Position := 85;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Updating user created Tags...';
  CopyTagChildrenToCommit;
  frmMain.ProgressBarUpdate.Position := 90;
  frmMain.ButtonUpdateTree.Caption := 'Please wait while we commit all changes. Adding "Any" to empty slots...';
  AddAnyChildIfEmpty(frmMain.TreeViewCommit);
  frmMain.ProgressBarUpdate.Position := 100;
// *** end update tree
  frmMain.ButtonUpdateTree.Enabled := true;
  frmMain.ButtonUpdateTree.Caption := 'Step 3 : Commit Changes (Update)';
  frmMain.ProgressBarUpdate.Visible := false;
  frmMain.TreeViewCommit.Enabled := true;
  frmMain.PageControl.Pages[2].TabVisible := true;
  frmMain.PageControl.Pages[3].TabVisible := true;
  frmMain.PageControl.Pages[4].TabVisible := true;
  frmMain.PageControl.Pages[5].TabVisible := true;
  frmMain.PageControl.Pages[6].TabVisible := true;
  frmMain.PageControl.Pages[7].TabVisible := true;
  frmMain.PageControl.Pages[8].TabVisible := true;
  frmMain.PageControl.Pages[9].TabVisible := true;
  frmMain.PageControl.Pages[10].TabVisible := true;
  frmMain.PageControl.Pages[12].TabVisible := true;
end;

procedure TfrmMain.MenuConvertToApplicationClick(Sender: TObject);
var
  selectedKey: string;
begin
  selectedKey := ClickedEditor.Cells[0, ClickedEditor.Row];
  MoveMatchingItemsToApplications(frmMain.TreeViewPolicies, selectedKey);
  frmMain.Applications.InsertRow(SelectedKey, ClickedEditor.Cells[1, ClickedEditor.Row], true);
  ClickedEditor.DeleteRow(ClickedEditor.Row);
  ServicesValues.DeleteRow(ClickedEditor.Row);
  ShowMessage('       The ID and the Service Name moved to Application. Proceed to Applications to set correct values         ');
  frmMain.PageControl.ActivePage := frmMain.TabSheetApplications;
  frmMain.Applications.Row := frmMain.Applications.RowCount - 1;
end;

procedure TfrmMain.ButtonGenerateClick(Sender: TObject);
begin
  frmMain.ButtonGenerate.Enabled := false;
  frmMain.MemoPaloCLI.Clear;
  CreateAddresses;
  CreateServices;
  CreateUrls;
  CreateTags;
  CreateAddressGroups;
  CreateServiceGroups;
  CreateURLGroups;
  GenerateCLICommands(frmMain.TreeViewCommit);
  if frmMain.CheckBoxAllLowerCase.Checked then frmMain.MemoPaloCLI.Text := lowercase(frmMain.MemoPaloCLI.Text);
  frmMain.GroupBoxPaloCLI.Caption := 'Palo Alto Configuration (' + IntToStr(frmMain.MemoPaloCLI.Lines.Count) + ' lines)';
  if frmMain.CheckBoxAllLowerCase.Checked then frmMain.GroupBoxPaloCLI.Caption := frmMain.GroupBoxPaloCLI.Caption + ' - All configuration set to lowercase';
  frmMain.ButtonGenerate.Enabled := true;
end;

procedure TfrmMain.ButtonManageableObjectsClick(Sender: TObject);
begin
  frmMain.PageControl.Pages[0].TabVisible := false;
  frmMain.PageControl.Pages[1].TabVisible := false;
  frmMain.PageControl.Pages[2].TabVisible := true;
  frmMain.ProgressBarProsessen.Max := GetParentItemCount(frmMain.TreeViewPolicies);
  PalofySourceTree;
  ReplaceInValue(frmMain.NetworkObjectsGroup, ' (group)', '');
  ReplaceInValue(frmMain.ServicesGroup, ' (group)', '');
  ReplaceInValue(frmMain.URLEntriesGroup, ' (group)', '');
  ExtractNetworkObjectsValues;
  ExtractServiceValues;
  ExtractURLValues;
  if FileExists(ExtractFilePath(Application.ExeName) + 'PaloAppID.txt') then frmMain.ComboBoxAPPID.Items.LoadFromFile(ExtractFilePath(Application.ExeName) + 'PaloAppID.txt');
  UpdateValuesFromComboBox(frmMain.Applications, frmMain.ComboBoxAPPID);
  frmMain.PageControl.Pages[3].TabVisible := true;
  frmMain.PageControl.Pages[4].TabVisible := true;
  frmMain.PageControl.Pages[5].TabVisible := true;
  frmMain.PageControl.Pages[6].TabVisible := true;
  frmMain.PageControl.Pages[7].TabVisible := true;
  frmMain.PageControl.Pages[8].TabVisible := true;
  frmMain.PageControl.Pages[9].TabVisible := true;
  frmMain.PageControl.Pages[10].TabVisible := true;
  frmMain.PageControl.Pages[11].TabVisible := true;
  frmMain.PageControl.ActivePage := frmMain.TabSheetManageableObjects;
  frmMain.GroupBoxProsessen.Visible := false;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  ParseText := TStringList.Create;
  TagList := TStringList.Create;
  GroupList := TStringList.Create;
  LabelApp.Caption := 'FirePalo Version ' + GetApplicationVersion;
  frmMain.PageControl.Pages[1].TabVisible := false;
  frmMain.PageControl.Pages[2].TabVisible := false;
  frmMain.PageControl.Pages[3].TabVisible := false;
  frmMain.PageControl.Pages[4].TabVisible := false;
  frmMain.PageControl.Pages[5].TabVisible := false;
  frmMain.PageControl.Pages[6].TabVisible := false;
  frmMain.PageControl.Pages[7].TabVisible := false;
  frmMain.PageControl.Pages[8].TabVisible := false;
  frmMain.PageControl.Pages[9].TabVisible := false;
  frmMain.PageControl.Pages[10].TabVisible := false;
  frmMain.PageControl.Pages[11].TabVisible := false;
  frmMain.PageControl.Pages[12].TabVisible := false;
  frmMain.PageControl.ActivePage := frmMain.TabSheetWelcome;
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
begin
  ParseText.Free;
  TagList.Free;
  GroupList.Free;
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin
  frmMain.ButtonOpenFPFile.SetFocus;
end;

procedure TfrmMain.LabelApplipediaClick(Sender: TObject);
var
  URL: string;
begin
  URL := frmMain.LabelApplipedia.Caption;
  ShellExecute(0, 'open', PChar(URL), nil, nil, SW_SHOWNORMAL);
end;

procedure TfrmMain.LabelCiscoAPPIDClick(Sender: TObject);
var
  URL: string;
begin
  URL := frmMain.LabelCiscoAPPID.Caption;
  ShellExecute(0, 'open', PChar(URL), nil, nil, SW_SHOWNORMAL);
end;

procedure TfrmMain.MenuAddTagClick(Sender: TObject);
var
  TagName : String;
  selectedKey: string;
begin
  selectedKey := ClickedEditor.Cells[0, ClickedEditor.Row];
  TagName := InputBox('Add tag to rules using this object ID', 'Please enter a Tag name:  ', '');
  if TagName = '' then exit;
  AddTagToRootNode(selectedKey, TagName);
  TagList.Add(TagName);
end;

procedure TfrmMain.MenuDeleteClick(Sender: TObject);
begin
  DeleteSelectedItemAndRelatedNodes;
end;

procedure TfrmMain.MenuExportListClick(Sender: TObject);
var
  FileName : String;
  URL: string;
  MessageText : String;
begin
  FileName := ExtractFilePath(Application.ExeName) + DateToStr(Now) + ' - ' + ClickedEditor.Name + '.txt';
  ClickedEditor.Strings.SaveToFile(FileName);
  MessageText := '          File exported to ' + FileName + '          ' + #10#13 + #10#13 +
                 '          DO NOT delete or add lines in the exported file!' + #10#13 +
                 '          DO NOT edit IDs! - Only edit values';
  ShowMessage(MessageText);
  URL := ExtractFilePath(Application.ExeName);
  ShellExecute(0, 'open', PChar(URL), nil, nil, SW_SHOWNORMAL);
end;

procedure TfrmMain.MenuImportListClick(Sender: TObject);
var
  OpenFileDialog: TOpenDialog;
begin
  OpenFileDialog := TOpenDialog.Create(nil);
    try
      OpenFileDialog.Title := 'Open ID and Value File';
      OpenFileDialog.Options := [ofFileMustExist, ofPathMustExist];
      if OpenFileDialog.Execute then
        begin
          ClickedEditor.Strings.Clear;
          ClickedEditor.Strings.LoadFromFile(OpenFileDialog.FileName);
        end;
    finally
      OpenFileDialog.Free;
    end;
end;

procedure TfrmMain.MenuLowercaseAllValuesClick(Sender: TObject);
begin
ModifyValueListEditorValues(ClickedEditor, 'lower');
end;

procedure TfrmMain.MenuShredClick(Sender: TObject);
begin
  TrimToMax(ClickedEditor);
end;

procedure TfrmMain.MenuUppercaseAllValuesClick(Sender: TObject);
begin
ModifyValueListEditorValues(ClickedEditor, 'upper');
end;

procedure TfrmMain.NetworkObjectsDblClick(Sender: TObject);
var
  SelectedKey: string;
  i: Integer;
begin
  SelectedKey := NetworkObjects.Keys[NetworkObjects.Row];
  for i := 1 to NetworkObjectsValues.RowCount - 1 do
  begin
    if NetworkObjectsValues.Keys[i] = SelectedKey then
    begin
      NetworkObjectsValues.Row := i;
      NetworkObjectsValues.SetFocus;
      Break;
    end;
  end;
end;

procedure TfrmMain.NetworkObjectsGroupDblClick(Sender: TObject);
var
  SelectedKey: string;
  i: Integer;
  Members : String;
begin
  SelectedKey := NetworkObjectsGroup.Keys[NetworkObjectsGroup.Row];
  if selectedkey = '' then exit;
  UpdateGroupsList(SelectedKey, GroupList, NetworkObjects);
  for i := 0 to GroupList.Count - 1 do
    begin
      members := members + GroupList.Strings[i] + #10#13;
    end;
  ShowMessage('GROUP MEMBERS :' + #10#13 + #10#13 + members);
end;

procedure TfrmMain.NetworkObjectsValuesDblClick(Sender: TObject);
var
  SelectedKey: string;
  i: Integer;
begin
  SelectedKey := NetworkObjectsValues.Keys[NetworkObjectsValues.Row];
  for i := 1 to NetworkObjects.RowCount - 1 do
  begin
    if NetworkObjects.Keys[i] = SelectedKey then
    begin
      NetworkObjects.Row := i;
      NetworkObjects.SetFocus;
      NetworkObjectsValues.SetFocus;
      Break;
    end;
  end;
end;

procedure TfrmMain.PopupMenuValuePopup(Sender: TObject);
var
  selectedKey: string;
begin
  ClickedEditor := TValueListEditor(frmMain.PopupMenuValue.PopupComponent);
  frmMain.MenuShred.Enabled := true;
  frmMain.MenuDelete.Enabled := true;
  frmMain.MenuAddTag.Enabled := true;
  frmMain.MenuUppercaseAllValues.Enabled := true;
  frmMain.MenuLowercaseAllValues.Enabled := true;
  frmMain.MenuConvertToApplication.Enabled := false;
  frmMain.MenuExportList.Enabled := true;
  frmMain.MenuImportList.Enabled := true;
  selectedKey := ClickedEditor.Cells[0, ClickedEditor.Row];
  if selectedKey = '' then
    begin
    frmMain.MenuDelete.Enabled := false;
    frmMain.MenuShred.Enabled := false;
    frmMain.MenuLowercaseAllValues.Enabled := false;
    frmMain.MenuUppercaseAllValues.Enabled := false;
    frmMain.MenuExportList.Enabled := false;
    frmMain.MenuImportList.Enabled := false;
    frmMain.MenuAddTag.Enabled := false;
    end;
  if ClickedEditor = frmMain.Action then
    begin
      frmMain.MenuShred.Enabled := false;
      frmMain.MenuDelete.Enabled := false;
    end;
  if Pos('Values', ClickedEditor.Name) > 0 then
    begin
      frmMain.MenuShred.Enabled := false;
      frmMain.MenuDelete.Enabled := false;
    end;
  if ClickedEditor = frmMain.Services then frmMain.MenuConvertToApplication.Enabled := true;
  if ClickedEditor = frmMain.Applications then frmMain.MenuShred.Enabled := false;
  if ClickedEditor = frmMain.Applications then frmMain.MenuUppercaseAllValues.Enabled := false;

end;

procedure TfrmMain.RuleNamesMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  Col, Row: Integer;
  selectedKey : String;
begin
  ClickedEditor2 := Sender as TValueListEditor;
  if Button = mbRight then
    begin
      ClickedEditor2.MouseToCell(X, Y, Col, Row);
      if (Row >= 1) and (Row < ClickedEditor2.RowCount) then
        begin
          ClickedEditor2.Row := Row;
          frmMain.StatusBar.SetFocus;
          frmMain.PopupMenuValuePopup(ClickedEditor2);
        end;
    end;
end;

procedure TfrmMain.ServicesDblClick(Sender: TObject);
var
  SelectedKey: string;
  i: Integer;
begin
  SelectedKey := Services.Keys[Services.Row];
  for i := 1 to ServicesValues.RowCount - 1 do
  begin
    if ServicesValues.Keys[i] = SelectedKey then
    begin
      ServicesValues.Row := i;
      ServicesValues.SetFocus;
      Break;
    end;
  end;
end;

procedure TfrmMain.ServicesGroupDblClick(Sender: TObject);
var
  SelectedKey: string;
  i: Integer;
  Members : String;
begin
  SelectedKey := ServicesGroup.Keys[ServicesGroup.Row];
  if selectedkey = '' then exit;
  UpdateGroupsList(SelectedKey, GroupList, Services);
  for i := 0 to GroupList.Count - 1 do
    begin
      members := members + GroupList.Strings[i] + #10#13;
    end;
  ShowMessage('GROUP MEMBERS :' + #10#13 + #10#13 + members);
end;

procedure TfrmMain.ServicesValuesDblClick(Sender: TObject);
var
  SelectedKey: string;
  i: Integer;
begin
  SelectedKey := ServicesValues.Keys[ServicesValues.Row];
  for i := 1 to Services.RowCount - 1 do
  begin
    if Services.Keys[i] = SelectedKey then
    begin
      Services.Row := i;
      Services.SetFocus;
      ServicesValues.SetFocus;
      Break;
    end;
  end;
end;

procedure TfrmMain.URLEntriesDblClick(Sender: TObject);
var
  SelectedKey: string;
  i: Integer;
begin
  SelectedKey := URLEntries.Keys[URLEntries.Row];
  for i := 1 to URLEntriesValues.RowCount - 1 do
  begin
    if URLEntriesValues.Keys[i] = SelectedKey then
    begin
      URLEntriesValues.Row := i;
      URLEntriesValues.SetFocus;
      Break;
    end;
  end;
end;

procedure TfrmMain.URLEntriesGroupDblClick(Sender: TObject);
var
  SelectedKey: string;
  i: Integer;
  Members : String;
begin
  SelectedKey := URLEntriesGroup.Keys[URLEntriesGroup.Row];
  if selectedkey = '' then exit;
  UpdateGroupsList(SelectedKey, GroupList, URLEntries);
  for i := 0 to GroupList.Count - 1 do
    begin
      members := members + GroupList.Strings[i] + #10#13;
    end;
  ShowMessage('GROUP MEMBERS :' + #10#13 + #10#13 + members);
end;

procedure TfrmMain.URLEntriesValuesDblClick(Sender: TObject);
var
  SelectedKey: string;
  i: Integer;
begin
  SelectedKey := URLEntriesValues.Keys[URLEntriesValues.Row];
  for i := 1 to URLEntries.RowCount - 1 do
  begin
    if URLEntries.Keys[i] = SelectedKey then
    begin
      URLEntries.Row := i;
      URLEntries.SetFocus;
      URLEntriesValues.SetFocus;
      Break;
    end;
  end;
end;

Procedure CleanStringList;
var
  CurrentLine : Integer;
begin
  for CurrentLine := 0 to ParseText.Count - 1 do
    begin
      if Pos('  Intrusion Policy', ParseText.Strings[CurrentLine]) > 0 then begin ParseText.Strings[CurrentLine] := ''; Continue; end;
      if Pos('  Source ISE Metadata', ParseText.Strings[CurrentLine]) > 0 then begin ParseText.Strings[CurrentLine] := ''; Continue; end;
      if Pos('  Dest ISE Metadata', ParseText.Strings[CurrentLine]) > 0 then begin ParseText.Strings[CurrentLine] := ''; Continue; end;
      if Pos('  Intrusion Policy', ParseText.Strings[CurrentLine]) > 0 then begin ParseText.Strings[CurrentLine] := ''; Continue; end;
      if Pos('  Logging Configuration', ParseText.Strings[CurrentLine]) > 0 then begin ParseText.Strings[CurrentLine] := ''; Continue; end;
      if Pos('  Safe Search', ParseText.Strings[CurrentLine]) > 0 then begin ParseText.Strings[CurrentLine] := ''; Continue; end;
      if Pos('  Rule Hits', ParseText.Strings[CurrentLine]) > 0 then begin ParseText.Strings[CurrentLine] := ''; Continue; end;
      if Pos('  Variable Set', ParseText.Strings[CurrentLine]) > 0 then begin ParseText.Strings[CurrentLine] := ''; Continue; end;
      if Pos('  URLs', ParseText.Strings[CurrentLine]) > 0 then begin ParseText.Strings[CurrentLine] := ''; Continue; end;
      if Pos('  Variable Set', ParseText.Strings[CurrentLine]) > 0 then begin ParseText.Strings[CurrentLine] := ''; Continue; end;
      if Pos('  Files', ParseText.Strings[CurrentLine]) > 0 then begin ParseText.Strings[CurrentLine] := ''; Continue; end;
      if Pos('  DC', ParseText.Strings[CurrentLine]) > 0 then begin ParseText.Strings[CurrentLine] := ''; Continue; end;
      if Pos('  Users', ParseText.Strings[CurrentLine]) > 0 then begin ParseText.Strings[CurrentLine] := ''; Continue; end;
    end;
end;

Procedure ParseOutput;
var
  CurrentLine : Integer;
  RulesStarted : Boolean;
begin
  RulesStarted := false;
  CleanStringList;
  frmMain.TreeViewPolicies.Items.BeginUpdate;
  // start parsing firepower text from the first line until the end //
  for CurrentLine := 0 to ParseText.Count - 1 do
    begin
      ActiveParseLine := CurrentLine;
      // skip all lines that are blank //
      if ParseText.Strings[CurrentLine] = '' then Continue;
      // if we have not started a rule and the new line does not contain rule start, then skip //
      if (not RulesStarted) then if (Pos('[ Rule: ', ParseText.Strings[CurrentLine]) = 0) then
        begin
         ParseText.Strings[CurrentLine] := '';
         Continue;
        end;
      // when we hit the advanced section, we now all rules have been parsed
      if Pos('[ Advanced Settings ]', ParseText.Strings[CurrentLine]) > 0  then
        begin
          ParseText.Strings[CurrentLine] := '';
          RulesStarted := false;
        end;
      // we will not start parsing until we hit a line that indicates we are starting a rule
      if Pos('[ Rule: ', ParseText.Strings[CurrentLine]) > 0 then
        begin
          RulesStarted := true;
          ParseText.Strings[CurrentLine] := ProcessRule(ParseText.Strings[CurrentLine]);
          Continue;
        end;
      // if we are still not processing a rule, then skip
      if (not RulesStarted) then Continue;
      NextLine := CurrentLine + 1;
      // when processing a rule, check if next line is the rest of the current line broken into multiple lines
      if NextLine <= ParseText.Count - 1 then
        begin
          if ParseText.Strings[NextLine] <> '' then
            begin
              if ParseText.Strings[NextLine][1] + ParseText.Strings[NextLine][2] <> '  ' then
                begin
                  ParseText.Strings[NextLine] := ParseText.Strings[CurrentLine] + ParseText.Strings[NextLine];
                  ParseText.Strings[CurrentLine] := '';
                  Continue;
                end;
            end;
        end;
      // output values
      ParseText.Strings[CurrentLine] := ProcessParameter(ParseText.Strings[CurrentLine]);
      Application.ProcessMessages;
      ParseText.Strings[CurrentLine] := '';
    end;
  frmMain.TreeViewPolicies.Items.EndUpdate;
end;

Function ProcessParameter(Parameter : string) : string;
var
  ParamNode : TTreeNode;
  ItemsNode : TTreeNode;
  I : Integer;
  ExtraItem : String;
  Done : Boolean;
begin

  I := ActiveParseLine + 1;
  Done := false;

if Pos('  Action', Parameter) > 0 then
  begin
    Result := Copy(Parameter, Pos(':', Parameter) + 2, Length(Parameter) - Pos(':', Parameter));
    ParamNode := frmMain.TreeViewPolicies.Items.AddChild(RuleNode, 'Action');
    frmMain.TreeViewPolicies.Items.AddChild(ParamNode, Result);
    exit;
  end;

  if Pos('  Source Zones', Parameter) > 0 then
  begin
    Result := Copy(Parameter, Pos(':', Parameter) + 2, Length(Parameter) - Pos(':', Parameter));
    ParamNode := frmMain.TreeViewPolicies.Items.AddChild(RuleNode, 'SourceZones');
    frmMain.TreeViewPolicies.Items.AddChild(ParamNode, Result);
      while Pos(':', ParseText.Strings[I]) = 0 do
        begin
          if ParseText.Strings[I] = '' then exit;
          ExtraItem := Trim(ParseText.Strings[I]);
          frmMain.TreeViewPolicies.Items.AddChild(ParamNode, ExtraItem);
          ParseText.Strings[I] := '';
          Inc(I);
        end;
    exit;
  end;

if Pos('  Destination Zones', Parameter) > 0 then
  begin
    Result := Copy(Parameter, Pos(':', Parameter) + 2, Length(Parameter) - Pos(':', Parameter));
    ParamNode := frmMain.TreeViewPolicies.Items.AddChild(RuleNode, 'DestinationZones');
    frmMain.TreeViewPolicies.Items.AddChild(ParamNode, Result);
      while Pos(':', ParseText.Strings[I]) = 0 do
        begin
          if ParseText.Strings[I] = '' then exit;
          ExtraItem := Trim(ParseText.Strings[I]);
          frmMain.TreeViewPolicies.Items.AddChild(ParamNode, ExtraItem);
          ParseText.Strings[I] := '';
          Inc(I);
        end;
    exit;
  end;

if Pos('  Source Networks', Parameter) > 0 then
  begin
    Result := Copy(Parameter, Pos(':', Parameter) + 2, Length(Parameter) - Pos(':', Parameter));
    ParamNode := frmMain.TreeViewPolicies.Items.AddChild(RuleNode, 'SourceNetworks');
    I := ActiveParseLine;
    ExtraItem := Result;
    while not done do
      begin
        if Pos('(group)', ParseText.Strings[I]) > 0 then
          begin
            ItemsNode := frmMain.TreeViewPolicies.Items.AddChild(ParamNode, ExtraItem);
            Inc(I);
              while FindFirstNonSpace(ParseText.Strings[I]) > (Pos(':', Parameter) + 2) do
                begin
                  if ParseText.Strings[I] = '' then exit;
                  ExtraItem := Trim(ParseText.Strings[I]);
                  frmMain.TreeViewPolicies.Items.AddChild(ItemsNode, ExtraItem);
                  ParseText.Strings[I] := '';
                  Inc(I);
                  ExtraItem := Trim(ParseText.Strings[I]);
                end;
          end else
          begin
            if Pos('  Source Networks', ParseText.Strings[I]) > 0 then ExtraItem := Result else ExtraItem := Trim(ParseText.Strings[I]);
            if ParseText.Strings[I] = '' then exit;
            frmMain.TreeViewPolicies.Items.AddChild(ParamNode, ExtraItem);
            ParseText.Strings[I] := '';
            Inc(I);
            ExtraItem := Trim(ParseText.Strings[I]);
            if ParseText.Strings[I] = '' then exit;
          end;
        Done := Pos('  :', ParseText.Strings[I]) > 0;
      end;
    exit;
  end;

if Pos('  Destination Networks', Parameter) > 0 then
  begin
    Result := Copy(Parameter, Pos(':', Parameter) + 2, Length(Parameter) - Pos(':', Parameter));
    ParamNode := frmMain.TreeViewPolicies.Items.AddChild(RuleNode, 'DestinationNetworks');
    I := ActiveParseLine;
    ExtraItem := Result;
    while not done do
      begin
        if Pos('(group)', ParseText.Strings[I]) > 0 then
          begin
            ItemsNode := frmMain.TreeViewPolicies.Items.AddChild(ParamNode, ExtraItem);
            Inc(I);
              while FindFirstNonSpace(ParseText.Strings[I]) > (Pos(':', Parameter) + 2) do
                begin
                  if ParseText.Strings[I] = '' then exit;
                  ExtraItem := Trim(ParseText.Strings[I]);
                  frmMain.TreeViewPolicies.Items.AddChild(ItemsNode, ExtraItem);
                  ParseText.Strings[I] := '';
                  Inc(I);
                  ExtraItem := Trim(ParseText.Strings[I]);
                end;
          end else
          begin
            if Pos('  Destination Networks', ParseText.Strings[I]) > 0 then ExtraItem := Result else ExtraItem := Trim(ParseText.Strings[I]);
            if ParseText.Strings[I] = '' then exit;
            frmMain.TreeViewPolicies.Items.AddChild(ParamNode, ExtraItem);
            ParseText.Strings[I] := '';
            Inc(I);
            ExtraItem := Trim(ParseText.Strings[I]);
            if ParseText.Strings[I] = '' then exit;
          end;
        Done := Pos('  :', ParseText.Strings[I]) > 0;
      end;
    exit;
  end;

if Pos('  Source Ports', Parameter) > 0 then
  begin
    Result := Copy(Parameter, Pos(':', Parameter) + 2, Length(Parameter) - Pos(':', Parameter));
    ParamNode := frmMain.TreeViewPolicies.Items.AddChild(RuleNode, 'SourcePorts');
    I := ActiveParseLine;
    ExtraItem := Result;
    while not done do
      begin
        if Pos('(group)', ParseText.Strings[I]) > 0 then
          begin
            ItemsNode := frmMain.TreeViewPolicies.Items.AddChild(ParamNode, ExtraItem);
            Inc(I);
              while FindFirstNonSpace(ParseText.Strings[I]) > (Pos(':', Parameter) + 2) do
                begin
                  if ParseText.Strings[I] = '' then exit;
                  ExtraItem := Trim(ParseText.Strings[I]);
                  frmMain.TreeViewPolicies.Items.AddChild(ItemsNode, ExtraItem);
                  ParseText.Strings[I] := '';
                  Inc(I);
                  ExtraItem := Trim(ParseText.Strings[I]);
                end;
          end else
          begin
            if Pos('  Source Ports', ParseText.Strings[I]) > 0 then ExtraItem := Result else ExtraItem := Trim(ParseText.Strings[I]);
            if ParseText.Strings[I] = '' then exit;
            frmMain.TreeViewPolicies.Items.AddChild(ParamNode, ExtraItem);
            ParseText.Strings[I] := '';
            Inc(I);
            ExtraItem := Trim(ParseText.Strings[I]);
            if ParseText.Strings[I] = '' then exit;
          end;
        Done := Pos('  :', ParseText.Strings[I]) > 0;
      end;
    exit;
  end;

if Pos('  Destination Ports', Parameter) > 0 then
  begin
    Result := Copy(Parameter, Pos(':', Parameter) + 2, Length(Parameter) - Pos(':', Parameter));
    ParamNode := frmMain.TreeViewPolicies.Items.AddChild(RuleNode, 'DestinationPorts');
    I := ActiveParseLine;
    ExtraItem := Result;
    while not done do
      begin
        if Pos('(group)', ParseText.Strings[I]) > 0 then
          begin
            ItemsNode := frmMain.TreeViewPolicies.Items.AddChild(ParamNode, ExtraItem);
            Inc(I);
              while FindFirstNonSpace(ParseText.Strings[I]) > (Pos(':', Parameter) + 2) do
                begin
                  if ParseText.Strings[I] = '' then exit;
                  ExtraItem := Trim(ParseText.Strings[I]);
                  frmMain.TreeViewPolicies.Items.AddChild(ItemsNode, ExtraItem);
                  ParseText.Strings[I] := '';
                  Inc(I);
                  ExtraItem := Trim(ParseText.Strings[I]);
                end;
          end else
          begin
            if Pos('  Destination Ports', ParseText.Strings[I]) > 0 then ExtraItem := Result else ExtraItem := Trim(ParseText.Strings[I]);
            if ParseText.Strings[I] = '' then exit;
            frmMain.TreeViewPolicies.Items.AddChild(ParamNode, ExtraItem);
            ParseText.Strings[I] := '';
            Inc(I);
            ExtraItem := Trim(ParseText.Strings[I]);
            if ParseText.Strings[I] = '' then exit;
          end;
        Done := Pos('  :', ParseText.Strings[I]) > 0;
      end;
    exit;
  end;

if Pos('  URL Entry', Parameter) > 0 then
  begin
    Result := Copy(Parameter, Pos(':', Parameter) + 2, Length(Parameter) - Pos(':', Parameter));
    ParamNode := frmMain.TreeViewPolicies.Items.AddChild(RuleNode, 'URLEntries');
    I := ActiveParseLine;
    ExtraItem := Result;
    while not done do
      begin
        if Pos('(group)', ParseText.Strings[I]) > 0 then
          begin
            ItemsNode := frmMain.TreeViewPolicies.Items.AddChild(ParamNode, ExtraItem);
            Inc(I);
              while FindFirstNonSpace(ParseText.Strings[I]) > (Pos(':', Parameter) + 2) do
                begin
                  if ParseText.Strings[I] = '' then exit;
                  ExtraItem := Trim(ParseText.Strings[I]);
                  frmMain.TreeViewPolicies.Items.AddChild(ItemsNode, ExtraItem);
                  ParseText.Strings[I] := '';
                  Inc(I);
                  ExtraItem := Trim(ParseText.Strings[I]);
                end;
          end else
          begin
            while Pos('  URL Entry', ParseText.Strings[I]) > 0 do
              begin
                ExtraItem := ParseText.Strings[I];
                ExtraItem := Copy(ExtraItem, Pos(':', ExtraItem) + 2, Length(ExtraItem) - Pos(':', ExtraItem));
                ExtraItem := Trim(ExtraItem);
                frmMain.TreeViewPolicies.Items.AddChild(ParamNode, ExtraItem);
                ParseText.Strings[I] := '';
                Inc(I);
                if ParseText.Strings[I] = '' then exit;
              end;
          end;
        Done := Pos('  URL Entry', ParseText.Strings[I]) = 0;
      end;
    exit;
  end;

if Pos('  Application', Parameter) > 0 then
  begin
    Result := Copy(Parameter, Pos(':', Parameter) + 2, Length(Parameter) - Pos(':', Parameter));
    ParamNode := frmMain.TreeViewPolicies.Items.AddChild(RuleNode, 'Applications');
    I := ActiveParseLine;
    ExtraItem := Result;
    while Pos('Application', ParseText.Strings[I]) > 0 do
      begin
        ExtraItem := ParseText.Strings[I];
        ExtraItem := Copy(ExtraItem, Pos(':', ExtraItem) + 2, Length(ExtraItem) - Pos(':', ExtraItem));
        ExtraItem := Trim(ExtraItem);
        frmMain.TreeViewPolicies.Items.AddChild(ParamNode, ExtraItem);
        ParseText.Strings[I] := '';
        Inc(I);
        if ParseText.Strings[I] = '' then exit;
      end;
    exit;
  end;

if Pos('  Category', Parameter) > 0 then
  begin
    Result := Copy(Parameter, Pos(':', Parameter) + 2, Length(Parameter) - Pos(':', Parameter));
    ParamNode := frmMain.TreeViewPolicies.Items.AddChild(RuleNode, 'Category');
    I := ActiveParseLine;
    ExtraItem := ParseText.Strings[I];
    while Pos('Category', ParseText.Strings[I]) > 0 do
      begin
        ExtraItem := ParseText.Strings[I];
        ExtraItem := Copy(ExtraItem, Pos(':', ExtraItem) + 2, Length(ExtraItem) - Pos(':', ExtraItem));
        ExtraItem := Trim(ExtraItem);
        frmMain.TreeViewPolicies.Items.AddChild(ParamNode, ExtraItem);
        ParseText.Strings[I] := '';
        Inc(I);
        if Pos('Reputation', ParseText.Strings[I]) > 0 then Inc(I);
        if ParseText.Strings[I] = '' then exit;
      end;
    exit;
  end;

if Pos('  Beginning', Parameter) > 0 then
  begin
    Result := Copy(Parameter, Pos(':', Parameter) + 2, Length(Parameter) - Pos(':', Parameter));
    ParamNode := frmMain.TreeViewPolicies.Items.AddChild(RuleNode, 'LogBeginning');
    frmMain.TreeViewPolicies.Items.AddChild(ParamNode, Result);
    exit;
  end;

if Pos('  End', Parameter) > 0 then
  begin
    Result := Copy(Parameter, Pos(':', Parameter) + 2, Length(Parameter) - Pos(':', Parameter));
    ParamNode := frmMain.TreeViewPolicies.Items.AddChild(RuleNode, 'LogEnd');
    frmMain.TreeViewPolicies.Items.AddChild(ParamNode, Result);
    exit;
  end;

if Pos('  Group', Parameter) > 0 then
  begin
    Result := Copy(Parameter, Pos(':', Parameter) + 2, Length(Parameter) - Pos(':', Parameter));
    ParamNode := frmMain.TreeViewPolicies.Items.AddChild(RuleNode, 'UserGroup');
    frmMain.TreeViewPolicies.Items.AddChild(ParamNode, Result);
    exit;
  end;
end;

Function ProcessRule(Rule : string) : string;
begin
  Result := Copy(Rule, Pos('[ Rule: ', Rule) + 8, Pos(']', Rule) - (Pos('[ Rule: ', Rule) + 8));
  RuleNode := frmMain.TreeViewPolicies.Items.Add(nil, Result)
end;

procedure SplitCommaObjects;
var
  i, j: Integer;
  Node: TTreeNode;
  NamePart, RemainingPart: string;
  SplitList: TStringList;
begin
  frmMain.TreeViewPolicies.Items.BeginUpdate;
  for i := 0 to frmMain.TreeViewPolicies.Items.Count - 1 do
  begin
    Node := frmMain.TreeViewPolicies.Items[i];
    if (Node.Text = 'SourceNetworks') or (Node.Text = 'DestinationNetworks') then
    begin
      for j := 0 to Node.Count - 1 do
      begin
        if Pos(',', Node.Item[j].Text) > 0 then
        begin
          NamePart := Copy(Node.Item[j].Text, 1, Pos('(', Node.Item[j].Text) - 1);
          RemainingPart := Copy(Node.Item[j].Text, Pos('(', Node.Item[j].Text) + 1, MaxInt);
          RemainingPart := StringReplace(RemainingPart, ')', '', [rfReplaceAll]);
          Trim(NamePart);
          NamePart := NamePart + '(group)';
          Trim(RemainingPart);
          Node.Item[j].Text := NamePart;
          SplitList := TStringList.Create;
          try
            SplitList.CommaText := RemainingPart;
            for NamePart in SplitList do
            begin
              Trim(NamePart);
              // NewNode := frmMain.TreeViewPolicies.Items.AddChild(Node.Item[j], NamePart);
              frmMain.TreeViewPolicies.Items.AddChild(Node.Item[j], NamePart);
            end;
          finally
            SplitList.Free;
          end;
        end;
      end;
      Application.ProcessMessages;
    end;
  end;
  frmMain.TreeViewPolicies.Items.EndUpdate;
end;

procedure PalofySourceTree;
var
  rootNode, datasetNode, valueNode, groupNode: TTreeNode;
  uniqueIDCounter: Integer;
  ruleUniqueID, valueUniqueID, groupValueUniqueID: String;
  datasetName, valueName, groupName: String;
  valueListEditor, groupValueListEditor: TValueListEditor;
  i: Integer;
begin
  uniqueIDCounter := 0;
  rootNode := frmMain.TreeViewPolicies.Items.GetFirstNode;
  frmMain.TreeViewPolicies.Items.BeginUpdate;
  frmMain.RuleNames.BeginUpdate;
  frmMain.Zones.BeginUpdate;
  frmMain.NetworkObjects.BeginUpdate;
  frmMain.Services.BeginUpdate;
  frmMain.URLEntries.BeginUpdate;
  frmMain.URLEntriesGroup.BeginUpdate;
  frmMain.Action.BeginUpdate;
  frmMain.Applications.BeginUpdate;
  frmMain.UserGroup.BeginUpdate;
  frmMain.Category.BeginUpdate;
  frmMain.LogEnd.BeginUpdate;
  frmMain.LogBeginning.BeginUpdate;
  while rootNode <> nil do
    begin
      Inc(uniqueIDCounter);
      ruleUniqueID := 'ID' + IntToStr(uniqueIDCounter);
      frmMain.RuleNames.Values[ruleUniqueID] := trim(rootNode.Text);
      rootNode.Text := ruleUniqueID;
      datasetNode := rootNode.GetFirstChild;
    while datasetNode <> nil do
    begin
      datasetName := datasetNode.Text;
      if datasetName = 'SourceZones' then datasetName := 'Zones';
      if datasetName = 'DestinationZones' then datasetName := 'Zones';
      if datasetName = 'SourceNetworks' then datasetName := 'NetworkObjects';
      if datasetName = 'DestinationNetworks' then datasetName := 'NetworkObjects';
      if datasetName = 'SourcePorts' then datasetName := 'Services';
      if datasetName = 'DestinationPorts' then datasetName := 'Services';
      valueListEditor := frmMain.FindComponent(datasetName) as TValueListEditor;
      if Assigned(valueListEditor) then
      begin
        valueNode := datasetNode.GetFirstChild;
        while valueNode <> nil do
        begin
          valueName := valueNode.Text;
          if (valueNode.HasChildren) and (Pos('(group)', valueName) > 0) then
          begin
            groupName := valueName;
            groupValueListEditor := frmMain.FindComponent(datasetName + 'Group') as TValueListEditor;
            groupValueUniqueID := '';
            for i := 1 to groupValueListEditor.RowCount - 1 do
              begin
                if groupValueListEditor.Values[groupValueListEditor.Keys[i]] = groupName then
                begin
                  groupValueUniqueID := groupValueListEditor.Keys[i];
                  Break;
                end;
              end;
            if groupValueUniqueID = '' then
              begin
                Inc(uniqueIDCounter);
                groupValueUniqueID := 'ID' + IntToStr(uniqueIDCounter);
                groupValueListEditor.Values[groupValueUniqueID] := groupName;
              end;
            valueNode.Text := groupValueUniqueID;
            groupNode := valueNode.GetFirstChild;
            while groupNode <> nil do
              begin
                valueName := groupNode.Text;
                valueUniqueID := '';
                for i := 1 to valueListEditor.RowCount - 1 do
                  begin
                    if valueListEditor.Values[valueListEditor.Keys[i]] = valueName then
                      begin
                        valueUniqueID := valueListEditor.Keys[i];
                        Break;
                      end;
                  end;
                  if valueUniqueID = '' then
                    begin
                      Inc(uniqueIDCounter);
                      valueUniqueID := 'ID' + IntToStr(uniqueIDCounter);
                      valueListEditor.Values[valueUniqueID] := valueName;
                    end;
                groupNode.Text := valueUniqueID;
                groupNode := groupNode.GetNextSibling;
              end;
          end else
            begin
              valueUniqueID := '';
              for i := 1 to valueListEditor.RowCount - 1 do
                begin
                  if valueListEditor.Values[valueListEditor.Keys[i]] = valueName then
                    begin
                      valueUniqueID := valueListEditor.Keys[i];
                      Break;
                    end;
                end;
                if valueUniqueID = '' then
                  begin
                    Inc(uniqueIDCounter);
                    valueUniqueID := 'ID' + IntToStr(uniqueIDCounter);
                    valueListEditor.Values[valueUniqueID] := valueName;
                  end;
              valueNode.Text := valueUniqueID;
            end;
          valueNode := valueNode.GetNextSibling;
        end;
      end;
      datasetNode := datasetNode.GetNextSibling;
      Application.ProcessMessages;
    end;
    rootNode := rootNode.GetNextSibling;
    frmMain.ProgressBarProsessen.Position := frmMain.ProgressBarProsessen.Position + 1;
    Application.ProcessMessages;
  end;
  frmMain.TreeViewPolicies.Items.EndUpdate;
  frmMain.RuleNames.EndUpdate;
  frmMain.Zones.EndUpdate;
  frmMain.NetworkObjects.EndUpdate;
  frmMain.Services.EndUpdate;
  frmMain.URLEntries.EndUpdate;
  frmMain.URLEntriesGroup.EndUpdate;
  frmMain.Action.EndUpdate;
  frmMain.Applications.EndUpdate;
  frmMain.UserGroup.EndUpdate;
  frmMain.Category.EndUpdate;
  frmMain.LogEnd.EndUpdate;
  frmMain.LogBeginning.EndUpdate;
end;

function FindChildNode(Node: TTreeNode; const Text: string): TTreeNode;
var
  ChildNode: TTreeNode;
begin
  Result := nil;
  if Assigned(Node) then
  begin
    ChildNode := Node.GetFirstChild;
    while Assigned(ChildNode) do
    begin
      if ChildNode.Text = Text then
      begin
        Result := ChildNode;
        Break;
      end;
      ChildNode := ChildNode.GetNextSibling;
    end;
  end;
end;

function FindNodeWithText(StartNode: TTreeNode; const Text: string): TTreeNode;
var
  Node: TTreeNode;
begin
  Result := nil;
  Node := StartNode;
  while Node <> nil do
  begin
    if Node.Text = Text then
    begin
      Result := Node;
      Break;
    end;
    Node := Node.GetNext;
  end;
end;

procedure FillAction;
var
  rootNode, clonedRootNode, actionNode, clonedActionNode, actionRefNode: TTreeNode;
  refID, actionValue: string;
begin
  frmMain.TreeViewCommit.Items.BeginUpdate;
  try
    for rootNode in frmMain.TreeViewPolicies.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        clonedRootNode := FindNodeWithText(frmMain.TreeViewCommit.Items.GetFirstNode, frmMain.RuleNames.Values[rootNode.Text]);
        actionNode := FindChildNode(rootNode, 'Action');
        if Assigned(actionNode) then
        begin
          actionRefNode := actionNode.getFirstChild;
          if Assigned(actionRefNode) then
          begin
            refID := actionRefNode.Text;
            actionValue := frmMain.Action.Values[refID];
            clonedActionNode := FindChildNode(clonedRootNode, 'Action');
            if Assigned(clonedActionNode) then
            begin
              frmMain.TreeViewCommit.Items.AddChild(clonedActionNode, actionValue);
            end;
          end;
        end;
      end;
    end;
  finally
    frmMain.TreeViewCommit.Items.EndUpdate;
  end;
end;

procedure FillSourceZones;
var
  rootNode, clonedRootNode, sourceZoneNode, clonedSourceZoneNode, zoneRefNode: TTreeNode;
  refID, zoneValue: string;
begin
  frmMain.TreeViewCommit.Items.BeginUpdate;
  try
    for rootNode in frmMain.TreeViewPolicies.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        clonedRootNode := FindNodeWithText(frmMain.TreeViewCommit.Items.GetFirstNode, frmMain.RuleNames.Values[rootNode.Text]);
        sourceZoneNode := FindChildNode(rootNode, 'SourceZones');
        if Assigned(sourceZoneNode) then
        begin
          clonedSourceZoneNode := FindChildNode(clonedRootNode, 'SourceZones');
          zoneRefNode := sourceZoneNode.getFirstChild;
          while Assigned(zoneRefNode) do
          begin
            refID := zoneRefNode.Text;
            zoneValue := frmMain.Zones.Values[refID];
            if Assigned(clonedSourceZoneNode) then
            begin
              frmMain.TreeViewCommit.Items.AddChild(clonedSourceZoneNode, zoneValue);
            end;
            zoneRefNode := zoneRefNode.getNextSibling;
          end;
        end;
      end;
    end;
  finally
    frmMain.TreeViewCommit.Items.EndUpdate;
  end;
end;

procedure FillDestinationZones;
var
  rootNode, clonedRootNode, sourceZoneNode, clonedSourceZoneNode, zoneRefNode: TTreeNode;
  refID, zoneValue: string;
begin
  frmMain.TreeViewCommit.Items.BeginUpdate;
  try
    for rootNode in frmMain.TreeViewPolicies.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        clonedRootNode := FindNodeWithText(frmMain.TreeViewCommit.Items.GetFirstNode, frmMain.RuleNames.Values[rootNode.Text]);
        sourceZoneNode := FindChildNode(rootNode, 'DestinationZones');
        if Assigned(sourceZoneNode) then
        begin
          clonedSourceZoneNode := FindChildNode(clonedRootNode, 'DestinationZones');
          zoneRefNode := sourceZoneNode.getFirstChild;
          while Assigned(zoneRefNode) do
          begin
            refID := zoneRefNode.Text;
            zoneValue := frmMain.Zones.Values[refID];
            if Assigned(clonedSourceZoneNode) then
            begin
              frmMain.TreeViewCommit.Items.AddChild(clonedSourceZoneNode, zoneValue);
            end;
            zoneRefNode := zoneRefNode.getNextSibling;
          end;
        end;
      end;
    end;
  finally
    frmMain.TreeViewCommit.Items.EndUpdate;
  end;
end;

procedure FillSourceNetworks;
var
  rootNode, clonedRootNode, sourceNetworkNode, clonedSourceNetworkNode, networkRefNode, clonedNetworkRefNode, childNode: TTreeNode;
  refID, networkValue, childID, childValue: string;
begin
  frmMain.TreeViewCommit.Items.BeginUpdate;
  try
    for rootNode in frmMain.TreeViewPolicies.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        clonedRootNode := FindNodeWithText(frmMain.TreeViewCommit.Items.GetFirstNode, frmMain.RuleNames.Values[rootNode.Text]);
        sourceNetworkNode := FindChildNode(rootNode, 'SourceNetworks');
        if Assigned(sourceNetworkNode) then
        begin
          clonedSourceNetworkNode := FindChildNode(clonedRootNode, 'SourceNetworks');
          networkRefNode := sourceNetworkNode.getFirstChild;
          while Assigned(networkRefNode) do
          begin
            refID := networkRefNode.Text;
            networkValue := frmMain.NetworkObjects.Values[refID];
            // Check if networkValue is in NetworkObjects or NetworkObjectsGroup
            if networkValue = '' then
            begin
              // If networkValue is in NetworkObjectsGroup, handle it and its children
              networkValue := frmMain.NetworkObjectsGroup.Values[refID];
              if networkValue <> '' then
              begin
                clonedNetworkRefNode := frmMain.TreeViewCommit.Items.AddChild(clonedSourceNetworkNode, networkValue);
                // Iterate through the children in the original NetworkObjectsGroup
                childNode := networkRefNode.getFirstChild;
                while Assigned(childNode) do
                begin
                  childID := childNode.Text;
                  childValue := frmMain.NetworkObjects.Values[childID];
                  if childValue = '' then
                    childValue := frmMain.NetworkObjectsGroup.Values[childID];
                  if childValue <> '' then
                  begin
                    frmMain.TreeViewCommit.Items.AddChild(clonedNetworkRefNode, childValue);
                  end;
                  childNode := childNode.getNextSibling;
                end;
              end;
            end
            else
            begin
              // If networkValue is directly in NetworkObjects, just add it
              frmMain.TreeViewCommit.Items.AddChild(clonedSourceNetworkNode, networkValue);
            end;
            networkRefNode := networkRefNode.getNextSibling;
          end;
        end;
      end;
    end;
  finally
    frmMain.TreeViewCommit.Items.EndUpdate;
  end;
end;

procedure FillDestinationNetworks;
var
  rootNode, clonedRootNode, sourceNetworkNode, clonedSourceNetworkNode, networkRefNode, clonedNetworkRefNode, childNode: TTreeNode;
  refID, networkValue, childID, childValue: string;
begin
  frmMain.TreeViewCommit.Items.BeginUpdate;
  try
    for rootNode in frmMain.TreeViewPolicies.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        clonedRootNode := FindNodeWithText(frmMain.TreeViewCommit.Items.GetFirstNode, frmMain.RuleNames.Values[rootNode.Text]);
        sourceNetworkNode := FindChildNode(rootNode, 'DestinationNetworks');
        if Assigned(sourceNetworkNode) then
        begin
          clonedSourceNetworkNode := FindChildNode(clonedRootNode, 'DestinationNetworks');
          networkRefNode := sourceNetworkNode.getFirstChild;
          while Assigned(networkRefNode) do
          begin
            refID := networkRefNode.Text;
            networkValue := frmMain.NetworkObjects.Values[refID];
            // Check if networkValue is in NetworkObjects or NetworkObjectsGroup
            if networkValue = '' then
            begin
              // If networkValue is in NetworkObjectsGroup, handle it and its children
              networkValue := frmMain.NetworkObjectsGroup.Values[refID];
              if networkValue <> '' then
              begin
                clonedNetworkRefNode := frmMain.TreeViewCommit.Items.AddChild(clonedSourceNetworkNode, networkValue);
                // Iterate through the children in the original NetworkObjectsGroup
                childNode := networkRefNode.getFirstChild;
                while Assigned(childNode) do
                begin
                  childID := childNode.Text;
                  childValue := frmMain.NetworkObjects.Values[childID];
                  if childValue = '' then
                    childValue := frmMain.NetworkObjectsGroup.Values[childID];
                  if childValue <> '' then
                  begin
                    frmMain.TreeViewCommit.Items.AddChild(clonedNetworkRefNode, childValue);
                  end;
                  childNode := childNode.getNextSibling;
                end;
              end;
            end
            else
            begin
              // If networkValue is directly in NetworkObjects, just add it
              frmMain.TreeViewCommit.Items.AddChild(clonedSourceNetworkNode, networkValue);
            end;
            networkRefNode := networkRefNode.getNextSibling;
          end;
        end;
      end;
    end;
  finally
    frmMain.TreeViewCommit.Items.EndUpdate;
  end;
end;

procedure LowercaseTreeView(Tree: TTreeView);
  procedure LowercaseNode(Node: TTreeNode);
  begin
    if Node = nil then Exit;
    Node.Text := LowerCase(Node.Text);
    LowercaseNode(Node.getFirstChild);
    LowercaseNode(Node.getNextSibling);
  end;
begin
  Tree.Items.BeginUpdate;
  try
    LowercaseNode(Tree.Items.GetFirstNode);
  finally
    Tree.Items.EndUpdate;
  end;
end;

procedure UpdateResultTree;
var
  rootNode, clonedNode: TTreeNode;
  i: Integer;
  children: array[0..14] of string;
begin
  children[0] := 'Action';
  children[1] := 'SourceZones';
  children[2] := 'DestinationZones';
  children[3] := 'SourceNetworks';
  children[4] := 'DestinationNetworks';
  children[5] := 'SourcePorts';
  children[6] := 'DestinationPorts';
  children[7] := 'Applications';
  children[8] := 'LogBeginning';
  children[9] := 'LogEnd';
  children[10] := 'URLEntries';
  children[11] := 'Category';
  children[12] := 'UserGroup';
  children[13] := 'Tag';
  children[14] := 'Description';
  frmMain.TreeViewCommit.Items.BeginUpdate;
  try
    frmMain.TreeViewCommit.Items.Clear;
    for rootNode in frmMain.TreeViewPolicies.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        clonedNode := frmMain.TreeViewCommit.Items.AddChild(nil, frmMain.RuleNames.Values[rootNode.Text]);
        for i := Low(children) to High(children) do
        begin
          frmMain.TreeViewCommit.Items.AddChild(clonedNode, children[i]);
        end;
      end;
    end;
  finally
    frmMain.TreeViewCommit.Items.EndUpdate;
  end;
end;

procedure FillSourcePorts;
var
  rootNode, clonedRootNode, sourceNetworkNode, clonedSourceNetworkNode, networkRefNode, clonedNetworkRefNode, childNode: TTreeNode;
  refID, networkValue, childID, childValue: string;
begin
  frmMain.TreeViewCommit.Items.BeginUpdate;
  try
    for rootNode in frmMain.TreeViewPolicies.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        clonedRootNode := FindNodeWithText(frmMain.TreeViewCommit.Items.GetFirstNode, frmMain.RuleNames.Values[rootNode.Text]);
        sourceNetworkNode := FindChildNode(rootNode, 'SourcePorts');
        if Assigned(sourceNetworkNode) then
        begin
          clonedSourceNetworkNode := FindChildNode(clonedRootNode, 'SourcePorts');
          networkRefNode := sourceNetworkNode.getFirstChild;
          while Assigned(networkRefNode) do
          begin
            refID := networkRefNode.Text;
            networkValue := frmMain.Services.Values[refID];
            if networkValue = '' then
            begin
              networkValue := frmMain.ServicesGroup.Values[refID];
              if networkValue <> '' then
              begin
                clonedNetworkRefNode := frmMain.TreeViewCommit.Items.AddChild(clonedSourceNetworkNode, networkValue);
                childNode := networkRefNode.getFirstChild;
                while Assigned(childNode) do
                begin
                  childID := childNode.Text;
                  childValue := frmMain.Services.Values[childID];
                  if childValue = '' then
                    childValue := frmMain.ServicesGroup.Values[childID];
                  if childValue <> '' then
                  begin
                    frmMain.TreeViewCommit.Items.AddChild(clonedNetworkRefNode, childValue);
                  end;
                  childNode := childNode.getNextSibling;
                end;
              end;
            end
            else
            begin
              frmMain.TreeViewCommit.Items.AddChild(clonedSourceNetworkNode, networkValue);
            end;
            networkRefNode := networkRefNode.getNextSibling;
          end;
        end;
      end;
    end;
  finally
    frmMain.TreeViewCommit.Items.EndUpdate;
  end;
end;

procedure FillDestinationPorts;
var
  rootNode, clonedRootNode, sourceNetworkNode, clonedSourceNetworkNode, networkRefNode, clonedNetworkRefNode, childNode: TTreeNode;
  refID, networkValue, childID, childValue: string;
begin
  frmMain.TreeViewCommit.Items.BeginUpdate;
  try
    for rootNode in frmMain.TreeViewPolicies.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        clonedRootNode := FindNodeWithText(frmMain.TreeViewCommit.Items.GetFirstNode, frmMain.RuleNames.Values[rootNode.Text]);
        sourceNetworkNode := FindChildNode(rootNode, 'DestinationPorts');
        if Assigned(sourceNetworkNode) then
        begin
          clonedSourceNetworkNode := FindChildNode(clonedRootNode, 'DestinationPorts');
          networkRefNode := sourceNetworkNode.getFirstChild;
          while Assigned(networkRefNode) do
          begin
            refID := networkRefNode.Text;
            networkValue := frmMain.Services.Values[refID];
            if networkValue = '' then
            begin
              networkValue := frmMain.ServicesGroup.Values[refID];
              if networkValue <> '' then
              begin
                clonedNetworkRefNode := frmMain.TreeViewCommit.Items.AddChild(clonedSourceNetworkNode, networkValue);
                childNode := networkRefNode.getFirstChild;
                while Assigned(childNode) do
                begin
                  childID := childNode.Text;
                  childValue := frmMain.Services.Values[childID];
                  if childValue = '' then
                    childValue := frmMain.ServicesGroup.Values[childID];
                  if childValue <> '' then
                  begin
                    frmMain.TreeViewCommit.Items.AddChild(clonedNetworkRefNode, childValue);
                  end;
                  childNode := childNode.getNextSibling;
                end;
              end;
            end
            else
            begin
              frmMain.TreeViewCommit.Items.AddChild(clonedSourceNetworkNode, networkValue);
            end;
            networkRefNode := networkRefNode.getNextSibling;
          end;
        end;
      end;
    end;
  finally
    frmMain.TreeViewCommit.Items.EndUpdate;
  end;
end;

procedure FillApplications;
var
  rootNode, clonedRootNode, sourceZoneNode, clonedSourceZoneNode, zoneRefNode: TTreeNode;
  refID, zoneValue: string;
begin
  frmMain.TreeViewCommit.Items.BeginUpdate;
  try
    for rootNode in frmMain.TreeViewPolicies.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        clonedRootNode := FindNodeWithText(frmMain.TreeViewCommit.Items.GetFirstNode, frmMain.RuleNames.Values[rootNode.Text]);
        sourceZoneNode := FindChildNode(rootNode, 'Applications');
        if Assigned(sourceZoneNode) then
        begin
          clonedSourceZoneNode := FindChildNode(clonedRootNode, 'Applications');
          zoneRefNode := sourceZoneNode.getFirstChild;
          while Assigned(zoneRefNode) do
          begin
            refID := zoneRefNode.Text;
            zoneValue := frmMain.Applications.Values[refID];
            if Assigned(clonedSourceZoneNode) then
            begin
              frmMain.TreeViewCommit.Items.AddChild(clonedSourceZoneNode, zoneValue);
            end;
            zoneRefNode := zoneRefNode.getNextSibling;
          end;
        end;
      end;
    end;
  finally
    frmMain.TreeViewCommit.Items.EndUpdate;
  end;
end;

procedure FillURLs;
var
  rootNode, clonedRootNode, sourceNetworkNode, clonedSourceNetworkNode, networkRefNode, clonedNetworkRefNode, childNode: TTreeNode;
  refID, networkValue, childID, childValue: string;
begin
  frmMain.TreeViewCommit.Items.BeginUpdate;
  try
    for rootNode in frmMain.TreeViewPolicies.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        clonedRootNode := FindNodeWithText(frmMain.TreeViewCommit.Items.GetFirstNode, frmMain.RuleNames.Values[rootNode.Text]);
        sourceNetworkNode := FindChildNode(rootNode, 'URLEntries');
        if Assigned(sourceNetworkNode) then
        begin
          clonedSourceNetworkNode := FindChildNode(clonedRootNode, 'URLEntries');
          networkRefNode := sourceNetworkNode.getFirstChild;
          while Assigned(networkRefNode) do
          begin
            refID := networkRefNode.Text;
            networkValue := frmMain.URLEntries.Values[refID];
            if networkValue = '' then
            begin
              networkValue := frmMain.URLEntriesGroup.Values[refID];
              if networkValue <> '' then
              begin
                clonedNetworkRefNode := frmMain.TreeViewCommit.Items.AddChild(clonedSourceNetworkNode, networkValue);
                childNode := networkRefNode.getFirstChild;
                while Assigned(childNode) do
                begin
                  childID := childNode.Text;
                  childValue := frmMain.URLEntries.Values[childID];
                  if childValue = '' then
                    childValue := frmMain.URLEntriesGroup.Values[childID];
                  if childValue <> '' then
                  begin
                    frmMain.TreeViewCommit.Items.AddChild(clonedNetworkRefNode, childValue);
                  end;
                  childNode := childNode.getNextSibling;
                end;
              end;
            end
            else
            begin
              frmMain.TreeViewCommit.Items.AddChild(clonedSourceNetworkNode, networkValue);
            end;
            networkRefNode := networkRefNode.getNextSibling;
          end;
        end;
      end;
    end;
  finally
    frmMain.TreeViewCommit.Items.EndUpdate;
  end;
end;

procedure AddAnyChildIfEmpty(Tree: TTreeView);
const
  NodeNames: array[1..7] of string = (
    'SourceZones', 'DestinationZones', 'SourceNetworks', 'DestinationNetworks',
    'DestinationPorts', 'Applications', 'URLEntries'
  );
var
  rootNode, childNode: TTreeNode;
  i: Integer;
  nodeName: string;
begin
  Tree.Items.BeginUpdate;
  try
    for rootNode in Tree.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        for i := Low(NodeNames) to High(NodeNames) do
        begin
          nodeName := NodeNames[i];
          childNode := FindChildNode(rootNode, nodeName);
          if Assigned(childNode) and (childNode.getFirstChild = nil) then
          begin
            Tree.Items.AddChild(childNode, 'any');
          end;
        end;
      end;
    end;
  finally
    Tree.Items.EndUpdate;
  end;
end;

procedure FillCategory;
var
  rootNode, clonedRootNode, sourceZoneNode, clonedSourceZoneNode, zoneRefNode: TTreeNode;
  refID, zoneValue: string;
begin
  frmMain.TreeViewCommit.Items.BeginUpdate;
  try
    for rootNode in frmMain.TreeViewPolicies.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        clonedRootNode := FindNodeWithText(frmMain.TreeViewCommit.Items.GetFirstNode, frmMain.RuleNames.Values[rootNode.Text]);
        sourceZoneNode := FindChildNode(rootNode, 'Category');
        if Assigned(sourceZoneNode) then
        begin
          clonedSourceZoneNode := FindChildNode(clonedRootNode, 'Category');
          zoneRefNode := sourceZoneNode.getFirstChild;
          while Assigned(zoneRefNode) do
          begin
            refID := zoneRefNode.Text;
            zoneValue := frmMain.Category.Values[refID];
            if Assigned(clonedSourceZoneNode) then
            begin
              frmMain.TreeViewCommit.Items.AddChild(clonedSourceZoneNode, zoneValue);
            end;
            zoneRefNode := zoneRefNode.getNextSibling;
          end;
        end;
      end;
    end;
  finally
    frmMain.TreeViewCommit.Items.EndUpdate;
  end;
end;

procedure FillLogBeginning;
var
  rootNode, clonedRootNode, sourceZoneNode, clonedSourceZoneNode, zoneRefNode: TTreeNode;
  refID, zoneValue: string;
begin
  frmMain.TreeViewCommit.Items.BeginUpdate;
  try
    for rootNode in frmMain.TreeViewPolicies.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        clonedRootNode := FindNodeWithText(frmMain.TreeViewCommit.Items.GetFirstNode, frmMain.RuleNames.Values[rootNode.Text]);
        sourceZoneNode := FindChildNode(rootNode, 'LogBeginning');
        if Assigned(sourceZoneNode) then
        begin
          clonedSourceZoneNode := FindChildNode(clonedRootNode, 'LogBeginning');
          zoneRefNode := sourceZoneNode.getFirstChild;
          while Assigned(zoneRefNode) do
          begin
            refID := zoneRefNode.Text;
            zoneValue := frmMain.LogBeginning.Values[refID];
            if Assigned(clonedSourceZoneNode) then
            begin
              frmMain.TreeViewCommit.Items.AddChild(clonedSourceZoneNode, zoneValue);
            end;
            zoneRefNode := zoneRefNode.getNextSibling;
          end;
        end;
      end;
    end;
  finally
    frmMain.TreeViewCommit.Items.EndUpdate;
  end;
end;

procedure FillLogEnd;
var
  rootNode, clonedRootNode, sourceZoneNode, clonedSourceZoneNode, zoneRefNode: TTreeNode;
  refID, zoneValue: string;
begin
  frmMain.TreeViewCommit.Items.BeginUpdate;
  try
    for rootNode in frmMain.TreeViewPolicies.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        clonedRootNode := FindNodeWithText(frmMain.TreeViewCommit.Items.GetFirstNode, frmMain.RuleNames.Values[rootNode.Text]);
        sourceZoneNode := FindChildNode(rootNode, 'LogEnd');
        if Assigned(sourceZoneNode) then
        begin
          clonedSourceZoneNode := FindChildNode(clonedRootNode, 'LogEnd');
          zoneRefNode := sourceZoneNode.getFirstChild;
          while Assigned(zoneRefNode) do
          begin
            refID := zoneRefNode.Text;
            zoneValue := frmMain.LogEnd.Values[refID];
            if Assigned(clonedSourceZoneNode) then
            begin
              frmMain.TreeViewCommit.Items.AddChild(clonedSourceZoneNode, zoneValue);
            end;
            zoneRefNode := zoneRefNode.getNextSibling;
          end;
        end;
      end;
    end;
  finally
    frmMain.TreeViewCommit.Items.EndUpdate;
  end;
end;

procedure FillUserGroup;
var
  rootNode, clonedRootNode, sourceZoneNode, clonedSourceZoneNode, zoneRefNode: TTreeNode;
  refID, zoneValue: string;
begin
  frmMain.TreeViewCommit.Items.BeginUpdate;
  try
    for rootNode in frmMain.TreeViewPolicies.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        clonedRootNode := FindNodeWithText(frmMain.TreeViewCommit.Items.GetFirstNode, frmMain.RuleNames.Values[rootNode.Text]);
        sourceZoneNode := FindChildNode(rootNode, 'UserGroup');
        if Assigned(sourceZoneNode) then
        begin
          clonedSourceZoneNode := FindChildNode(clonedRootNode, 'UserGroup');
          zoneRefNode := sourceZoneNode.getFirstChild;
          while Assigned(zoneRefNode) do
          begin
            refID := zoneRefNode.Text;
            zoneValue := frmMain.UserGroup.Values[refID];
            if Assigned(clonedSourceZoneNode) then
            begin
              frmMain.TreeViewCommit.Items.AddChild(clonedSourceZoneNode, zoneValue);
            end;
            zoneRefNode := zoneRefNode.getNextSibling;
          end;
        end;
      end;
    end;
  finally
    frmMain.TreeViewCommit.Items.EndUpdate;
  end;
end;

procedure AddUrlCategoryToTag(Tree: TTreeView);
var
  rootNode, categoryNode, tagNode: TTreeNode;
begin
  Tree.Items.BeginUpdate;
  try
    for rootNode in Tree.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        categoryNode := FindChildNode(rootNode, 'Category');
        tagNode := FindChildNode(rootNode, 'Tag');
        if Assigned(categoryNode) and (categoryNode.getFirstChild <> nil) and Assigned(tagNode) then
        begin
          Tree.Items.AddChild(tagNode, 'url-category');
        end;
      end;
    end;
  finally
    Tree.Items.EndUpdate;
  end;
  taglist.Add('url-category');
end;

procedure AddUserGroupToTag(Tree: TTreeView);
var
  rootNode, categoryNode, tagNode: TTreeNode;
begin
  Tree.Items.BeginUpdate;
  try
    for rootNode in Tree.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        categoryNode := FindChildNode(rootNode, 'UserGroup');
        tagNode := FindChildNode(rootNode, 'Tag');
        if Assigned(categoryNode) and (categoryNode.getFirstChild <> nil) and Assigned(tagNode) then
        begin
          Tree.Items.AddChild(tagNode, 'user-id');
        end;
      end;
    end;
  finally
    Tree.Items.EndUpdate;
  end;
  taglist.Add('user-id');
end;

procedure AddDescriptionToRootNodes(Tree: TTreeView);
var
  rootNode, descriptionNode: TTreeNode;
  rootNodeCounter: Integer;
begin
  Tree.Items.BeginUpdate;
  try
    rootNodeCounter := 0;
    for rootNode in Tree.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        Inc(rootNodeCounter);
        descriptionNode := FindChildNode(rootNode, 'Description');
        if Assigned(descriptionNode) then
        begin
          Tree.Items.AddChild(descriptionNode, 'Imported-rule-#' + IntToStr(rootNodeCounter));
        end;
      end;
    end;
  finally
    Tree.Items.EndUpdate;
  end;
end;

procedure GenerateCLICommands(Tree: TTreeView);
const
  ParentNodeNames: array[0..12] of string = (
    'Action', 'SourceZones', 'DestinationZones', 'SourceNetworks', 'DestinationNetworks',
    'SourcePorts', 'DestinationPorts', 'Applications', 'LogBeginning', 'LogEnd',
    'URLEntries', 'Tag', 'Description');
var
  rootNode, childNode, subChildNode: TTreeNode;
  i: Integer;
begin
  if frmMain.EditPanorama.Text <> '' then CommandString := 'set device-group ' + frmMain.EditPanorama.Text +  ' pre-rulebase security rules "' else CommandString := 'set rulebase security rules "';
  for rootNode in Tree.Items do
  begin
    Application.ProcessMessages;
    if rootNode.Parent = nil then
    begin
      for i := Low(ParentNodeNames) to High(ParentNodeNames) do
      begin
        childNode := FindChildNode(rootNode, ParentNodeNames[i]);
        if Assigned(childNode) then
        begin
          subChildNode := childNode.getFirstChild;
          while Assigned(subChildNode) do
          begin
            if ParentNodeNames[i] = 'Action' then
              GenerateCLIAction(rootNode.Text, subChildNode.Text)
            else if ParentNodeNames[i] = 'SourceZones' then
              GenerateCLISourceZones(rootNode.Text, subChildNode.Text)
            else if ParentNodeNames[i] = 'DestinationZones' then
              GenerateCLIDestinationZones(rootNode.Text, subChildNode.Text)
            else if ParentNodeNames[i] = 'SourceNetworks' then
              GenerateCLISourceNetworks(rootNode.Text, subChildNode.Text)
            else if ParentNodeNames[i] = 'DestinationNetworks' then
              GenerateCLIDestinationNetworks(rootNode.Text, subChildNode.Text)
            else if ParentNodeNames[i] = 'SourcePorts' then
              GenerateCLISourcePorts(rootNode.Text, subChildNode.Text)
            else if ParentNodeNames[i] = 'DestinationPorts' then
              GenerateCLIDestinationPorts(rootNode.Text, subChildNode.Text)
            else if ParentNodeNames[i] = 'Applications' then
              GenerateCLIApplications(rootNode.Text, subChildNode.Text)
            else if ParentNodeNames[i] = 'LogBeginning' then
              GenerateCLILogBeginning(rootNode.Text, subChildNode.Text)
            else if ParentNodeNames[i] = 'LogEnd' then
              GenerateCLILogEnd(rootNode.Text, subChildNode.Text)
            else if ParentNodeNames[i] = 'URLEntries' then
              GenerateCLIURLEntries(rootNode.Text, subChildNode.Text)
            else if ParentNodeNames[i] = 'Tag' then
              GenerateCLITag(rootNode.Text, subChildNode.Text)
            else if ParentNodeNames[i] = 'Description' then
              GenerateCLIDescription(rootNode.Text, subChildNode.Text);
            subChildNode := subChildNode.getNextSibling;
          end;
        end;
      end;
    end;
  end;
end;

procedure GenerateCLIAction(rootNodeName, foundChildName: string);
begin
  if lowercase(foundChildName) = 'disabled' then
    begin
      frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" action drop');
      frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" disabled yes');
    end else frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" action ' + lowercase(foundChildName));
  // *** extra options *** //
  if frmMain.LabeledEditProfileGroup.Text <> '' then frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" profile-setting group ' + frmMain.LabeledEditProfileGroup.Text);
  if frmMain.LabeledEditLogSetting.Text <> '' then frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" log-setting ' + frmMain.LabeledEditLogSetting.Text);
  if frmMain.CheckBoxLogBeginning.Checked then frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" log-start yes');
  if frmMain.CheckBoxLogEnd.Checked then frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" log-end yes');
  if frmMain.CheckBoxUseApplicationDefault.Checked then frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" service application-default');
end;

procedure GenerateCLISourceZones(rootNodeName, foundChildName: string);
begin
  if pos('[', foundChildName) > 0 then frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" from ' + foundChildName) else
   frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" from "' + foundChildName + '"');
end;

procedure GenerateCLIDestinationZones(rootNodeName, foundChildName: string);
begin
  if pos('[', foundChildName) > 0 then frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" to ' + foundChildName) else
   frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" to "' + foundChildName + '"');
end;

procedure GenerateCLISourceNetworks(rootNodeName, foundChildName: string);
begin
  if pos('[', foundChildName) > 0 then frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" source ' + foundChildName) else
   frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" source "' + foundChildName + '"');
end;

procedure GenerateCLIDestinationNetworks(rootNodeName, foundChildName: string);
begin
  if pos('[', foundChildName) > 0 then frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" destination ' + foundChildName) else
   frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" destination "' + foundChildName + '"');
end;

procedure GenerateCLISourcePorts(rootNodeName, foundChildName: string);
begin
  if pos('[', foundChildName) > 0 then frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" service ' + foundChildName) else
   frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" service "' + foundChildName + '"');
end;

procedure GenerateCLIDestinationPorts(rootNodeName, foundChildName: string);
begin
  if pos('[', foundChildName) > 0 then frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" service ' + foundChildName) else
   frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" service "' + foundChildName + '"');
end;

procedure GenerateCLIApplications(rootNodeName, foundChildName: string);
begin
  if pos('[', foundChildName) > 0 then frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" application ' + foundChildName) else
   frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" application "' + foundChildName + '"');
end;

procedure GenerateCLILogBeginning(rootNodeName, foundChildName: string);
begin
  if frmMain.CheckBoxLogBeginning.Checked then exit;
  if lowercase(foundChildName) = 'disabled' then frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" log-start no') else frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" log-start yes')
end;

procedure GenerateCLILogEnd(rootNodeName, foundChildName: string);
begin
  if frmMain.CheckBoxLogEnd.Checked then exit;
  if lowercase(foundChildName) = 'disabled' then frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" log-end no') else frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" log-end yes')
end;

procedure GenerateCLIURLEntries(rootNodeName, foundChildName: string);
begin
  if pos('[', foundChildName) > 0 then frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" category ' + foundChildName) else
   frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" category "' + foundChildName + '"');
end;

procedure GenerateCLITag(rootNodeName, foundChildName: string);
begin
  if pos('[', foundChildName) > 0 then frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" tag ' + foundChildName) else
   frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" tag "' + foundChildName + '"');
end;

procedure GenerateCLIDescription(rootNodeName, foundChildName: string);
begin
  frmMain.MemoPaloCLI.Lines.Add(CommandString + rootNodeName + '" description "' + foundChildName + '"');
end;

procedure ReplaceInValue(ValueListEditor: TValueListEditor; const SearchWord, ReplaceWord: string);
var
  i: Integer;
  Key, OriginalValue, ModifiedValue: string;
begin
  for i := 1 to ValueListEditor.RowCount - 1 do
  begin
    Key := ValueListEditor.Keys[i];
    OriginalValue := ValueListEditor.Values[Key];
    ModifiedValue := StringReplace(OriginalValue, SearchWord, ReplaceWord, [rfReplaceAll, rfIgnoreCase]);
    ValueListEditor.Values[Key] := ModifiedValue;
  end;
end;

Procedure ExtractNetworkObjectsValues;
var
  i: Integer;
  Key, Value, OriginalValue, Name: string;
begin
  for i := 1 to frmMain.NetworkObjects.RowCount - 1 do
  begin
    Key := frmMain.NetworkObjects.Keys[i];
    OriginalValue := frmMain.NetworkObjects.Values[Key];
    if Pos('(', OriginalValue) > 0 then
      begin
        Value := Copy(OriginalValue, Pos('(', OriginalValue), Length(OriginalValue) - Pos('(', OriginalValue) + 1);
        Value := StringReplace(Value, '(', '', [rfReplaceAll, rfIgnoreCase]);
        Value := StringReplace(Value, ')', '', [rfReplaceAll, rfIgnoreCase]);
        Value := Trim(Value);
        if (Value <> '') and (pos('.', Value) = 0) and (pos(':', Value) = 0) then Value := Value + '                   ******    CHANGE ME!    ******';
        Name := Copy(OriginalValue, 0, Pos('(', OriginalValue) - 1);
        Name := Trim(Name);
      end else
      begin
       Value := OriginalValue;
       if (Value <> '') and (pos('.', Value) = 0) and (pos(':', Value) = 0) then Value := Value + '                    ******    CHANGE ME!    ******';
       Name := 'literal-' + OriginalValue;
       Name := StringReplace(Name, '/', '-', [rfReplaceAll, rfIgnoreCase]);
       Name := StringReplace(Name, ':', '_', [rfReplaceAll, rfIgnoreCase]);
      end;
    frmMain.NetworkObjectsValues.Values[Key] := Value;
    frmMain.NetworkObjects.Values[Key] := Name;
  end;
end;

Procedure ExtractServiceValues;
var
  i: Integer;
  Key, Value, OriginalValue, Name, AddName: string;
begin
  for i := 1 to frmMain.Services.RowCount - 1 do
  begin
    Key := frmMain.Services.Keys[i];
    OriginalValue := frmMain.Services.Values[Key];
    OriginalValue := StringReplace(OriginalValue, '(', '', [rfReplaceAll, rfIgnoreCase]);
    OriginalValue := StringReplace(OriginalValue, ')', '', [rfReplaceAll, rfIgnoreCase]);
    OriginalValue := StringReplace(OriginalValue, 'protocol 6, port', ';tcp port', [rfReplaceAll, rfIgnoreCase]);
    OriginalValue := StringReplace(OriginalValue, 'protocol 17, port', ';udp port', [rfReplaceAll, rfIgnoreCase]);
    OriginalValue := StringReplace(OriginalValue, 'protocol', ' ;convert_to_application protocol', [rfReplaceAll, rfIgnoreCase]);
    OriginalValue := StringReplace(OriginalValue, ',', '', [rfReplaceAll, rfIgnoreCase]);
    if Pos(';', OriginalValue) = 1 then
      begin
        Value := StringReplace(OriginalValue, ';', '', [rfReplaceAll, rfIgnoreCase]);
        Name := Copy(Value, 1, 3) + '-' + Copy(Value, (Pos('port ', Value) + 5), (Length(Value) - (Pos('port ', Value) + 4)));
       end else
      begin
       Name := Copy(OriginalValue, 0, (Pos(';', OriginalValue) - 1));
       Value := Copy(OriginalValue, (Pos(';', OriginalValue) + 1), (Length(OriginalValue) - (Pos(';', OriginalValue))));
       Value := StringReplace(Value, ';', '', [rfReplaceAll, rfIgnoreCase]);
      end;
    Value := Trim(Value);
    Name := Trim(Name);
    if Pos('convert_to_application', Value) = 1 then AddName := '_Convert-To-Application' else AddName := '';
    if Name = '' then Name := 'No-Name';
    Name := Name + AddName;
    frmMain.ServicesValues.Values[Key] := Value;
    frmMain.Services.Values[Key] := Name;
  end;
end;

Procedure ExtractURLValues;
var
  i: Integer;
  Key, Value, OriginalValue, Name, Addname: string;
begin
  for i := 1 to frmMain.URLEntries.RowCount - 1 do
  begin
    Key := frmMain.URLEntries.Keys[i];
    OriginalValue := frmMain.URLEntries.Values[Key];
    if Pos(' - ', OriginalValue) > 0 then
      begin
        Value := Copy(OriginalValue, Pos(' - ', OriginalValue) + 3, Length(OriginalValue) - Pos(' - ', OriginalValue) + 2);
        Value := Trim(Value);
        Name := Copy(OriginalValue, 0, Pos(' - ', OriginalValue) - 1);
        Name := Trim(Name);
        Name := 'url-' + Name;
      end else
      begin
       Value := OriginalValue;
       Name := 'url-' + OriginalValue;
      end;
    if Pos(' ', Value) > 0 then AddName := '                    ******    CHANGE ME!    ******' else AddName := '';
    Value := Value + AddName;
    Name := Name + ' ' + Trim(AddName);
    frmMain.URLEntriesValues.Values[Key] := Value;
    frmMain.URLEntries.Values[Key] := Name;
  end;
end;

procedure UpdateValuesFromComboBox(ValueList: TValueListEditor; ComboBox: TComboBox);
var
  i, j: Integer;
  tempStr, valueStr, modifiedStr: string;
  found: Boolean;
begin
  for i := 1 to ValueList.RowCount - 1 do
  begin
    valueStr := ValueList.Cells[1, i];
    tempStr := Trim(Copy(valueStr, 1, Pos('(', valueStr) - 1));
    tempStr := Trim(tempStr);
    found := False;
    for j := 0 to ComboBox.Items.Count - 1 do
    begin
      if AnsiCompareText(tempStr, ComboBox.Items[j]) = 0 then
      begin
        ValueList.Cells[1, i] := ComboBox.Items[j];
        found := True;
        Break;
      end;
    end;
    if not found and (Pos(' ', tempStr) > 0) then
    begin
      modifiedStr := StringReplace(tempStr, ' ', '-', [rfReplaceAll]);
      for j := 0 to ComboBox.Items.Count - 1 do
      begin
        if AnsiCompareText(modifiedStr, ComboBox.Items[j]) = 0 then
        begin
          ValueList.Cells[1, i] := ComboBox.Items[j];
//          found := True;
          Break;
        end;
      end;
    end;
      if TempStr = 'HTTP' then ValueList.Cells[1, i] := 'web-browsing';
      if TempStr = 'HTTPS' then ValueList.Cells[1, i] := 'ssl';
      if TempStr = 'ICMP for IPv6' then ValueList.Cells[1, i] := 'ipv6-icmp';
      if TempStr = 'LDAPS' then ValueList.Cells[1, i] := '[ ldap ssl ]';
      if TempStr = 'NetBIOS-ssn' then ValueList.Cells[1, i] := 'netbios-ss';
      if TempStr = 'RDP' then ValueList.Cells[1, i] := 'ms-rdp';
      if TempStr = 'SMTPS' then ValueList.Cells[1, i] := '[ smtp ssl ]';
      if TempStr = 'DCE/RPC' then ValueList.Cells[1, i] := 'rpc';
      if TempStr = 'MS SQL' then ValueList.Cells[1, i] := 'mssql-db';
      if TempStr = 'AirPlay' then ValueList.Cells[1, i] := 'apple-airplay';
      if TempStr = 'AirTunes' then ValueList.Cells[1, i] := 'apple-airplay';
      if TempStr = 'Microsoft Update' then ValueList.Cells[1, i] := 'ms-update';
      if TempStr = 'Microsoft Teams' then ValueList.Cells[1, i] := 'ms-teams';
      if TempStr = 'OpenDNS' then ValueList.Cells[1, i] := 'opendns';
      if TempStr = 'Symantec System Center' then ValueList.Cells[1, i] := 'symantec-syst-center';
      if TempStr = 'Office 365' then ValueList.Cells[1, i] := 'ms-office365';
      if TempStr = 'FTP Data' then ValueList.Cells[1, i] := 'ftp';
      if TempStr = 'SSL client' then ValueList.Cells[1, i] := 'ssl';
      if TempStr = 'SFTP' then ValueList.Cells[1, i] := '[ ftp ssl ]';
      if TempStr = 'FTP Active' then ValueList.Cells[1, i] := 'ftp';
      if TempStr = 'FTP Passive' then ValueList.Cells[1, i] := 'ftp';
      if TempStr = 'Google APIs' then ValueList.Cells[1, i] := 'google-base';
      if TempStr = 'FTPS' then ValueList.Cells[1, i] := '[ ftp ssl ]';
      if TempStr = 'OpenVPN' then ValueList.Cells[1, i] := 'open-vpn';
      if TempStr = 'RADIUS-acct' then ValueList.Cells[1, i] := 'radius';
      if TempStr = 'Windows Update' then ValueList.Cells[1, i] := 'ms-update';
      if TempStr = 'Microsoft Azure' then ValueList.Cells[1, i] := 'windows-azure';
      if TempStr = 'Exchange Online' then ValueList.Cells[1, i] := 'ms-exchange';
      if TempStr = 'Microsoft Visual Studio' then ValueList.Cells[1, i] := 'ms-visual-studio-tfs';
      if TempStr = 'Flash Video' then ValueList.Cells[1, i] := 'adobe-media-player';
      if TempStr = 'OCSPD' then ValueList.Cells[1, i] := 'ocsp';
      if TempStr = 'Azure cloud portal' then ValueList.Cells[1, i] := 'windows-azure';
      if TempStr = 'AD DRS' then ValueList.Cells[1, i] := 'active-directory';
      if TempStr = 'Microsoft Global Catalog' then ValueList.Cells[1, i] := 'active-directory';
      if TempStr = 'Netlogon' then ValueList.Cells[1, i] := 'ms-netlogon';
      if TempStr = 'ID83=SMBv1' then ValueList.Cells[1, i] := 'ms-ds-smbv1';
      if TempStr = 'ID84=SMBv2' then ValueList.Cells[1, i] := 'ms-ds-smbv2';
      if TempStr = 'SMBv3-encrypted' then ValueList.Cells[1, i] := 'ms-ds-smbv3';
      if TempStr = 'Microsoft System Center Operations Manager' then ValueList.Cells[1, i] := 'ms-scom';
      if TempStr = 'Operations Manager - Health Service' then ValueList.Cells[1, i] := 'ms-scom';
      if TempStr = 'symantec-syst-center' then ValueList.Cells[1, i] := 'symantec-syst-center';
      if TempStr = 'SymantecUpdates' then ValueList.Cells[1, i] := 'symantec-av-update';
      if TempStr = 'NetBIOS-dgm' then ValueList.Cells[1, i] := 'netbios-dg';
      if TempStr = 'TACACS+' then ValueList.Cells[1, i] := 'tacacs-plus';
      if TempStr = 'IMAPS' then ValueList.Cells[1, i] := '[ imap ssl ]';
      if TempStr = 'TNS/Oracle' then ValueList.Cells[1, i] := 'oracle';
      if Pos(':', ValueStr) > 0 then ValueList.Cells[1, i] := ValueStr + ' - Convert to application filter';
  end;
end;

procedure DeleteSelectedItemAndRelatedNodes;
var
  selectedKey: string;
  selectedValue: string;
  i: Integer;
  associatedEditor: TValueListEditor;
  node, nextNode: TTreeNode;
begin
  if Assigned(ClickedEditor) and (ClickedEditor.Row > 0) and (ClickedEditor.Row <= ClickedEditor.RowCount) then
  begin
    selectedKey := ClickedEditor.Cells[0, ClickedEditor.Row];
    selectedValue := ClickedEditor.Cells[1, ClickedEditor.Row];
    ClickedEditor.DeleteRow(ClickedEditor.Row);
    associatedEditor := TValueListEditor(frmMain.FindComponent(ClickedEditor.Name + 'Values'));
    if Assigned(associatedEditor) then
    begin
      for i := 1 to associatedEditor.RowCount - 1 do
      begin
        if AnsiCompareText(associatedEditor.Cells[0, i], selectedKey) = 0 then
        begin
          associatedEditor.DeleteRow(i);
          Break;
        end;
      end;
    end;
    i := frmMain.TreeViewPolicies.Items.Count - 1;
    frmMain.TreeViewPolicies.Items.BeginUpdate;
    while i >= 0 do
    begin
      node := frmMain.TreeViewPolicies.Items.Item[i];
      if Assigned(node) and (node.Text = selectedKey) then
      begin
        nextNode := node.GetNextSibling;
        node.Delete;
        node := nextNode;
      end
      else
      begin
        Dec(i);
      end;
    end;
    frmMain.TreeViewPolicies.Items.EndUpdate;
  end
  else
  begin
    ShowMessage('No item is selected!');
  end;
end;

procedure AddTagToRootNode(const MatchString, TagAddString: string);
var
  i, j: Integer;
  node, parentNode, tagnode: TTreeNode;
  tagNodeExists: Boolean;
begin
  tagnode := nil;
  i := frmMain.TreeViewPolicies.Items.Count - 1;
  while i >= 0 do
  begin
    node := frmMain.TreeViewPolicies.Items.Item[i];
    if node.Text = MatchString then
    begin
      parentNode := node;
      while Assigned(parentNode.Parent) do
        parentNode := parentNode.Parent;
      tagNodeExists := False;
      for j := 0 to parentNode.Count - 1 do
      begin
        if parentNode.Item[j].Text = 'Tag' then
        begin
          tagnode := parentNode.Item[j];
          tagNodeExists := True;
          Break;
        end;
      end;
      if not tagNodeExists then
        tagnode := frmMain.TreeViewPolicies.Items.AddChild(parentNode, 'Tag');
      frmMain.TreeViewPolicies.Items.AddChild(tagnode, TagAddString);
    end;
    Dec(i);
  end;
end;

procedure CopyTagChildrenToCommit;
var
  rootNode, clonedRootNode, sourceZoneNode, clonedSourceZoneNode, zoneRefNode: TTreeNode;
begin
  frmMain.TreeViewCommit.Items.BeginUpdate;
  try
    for rootNode in frmMain.TreeViewPolicies.Items do
    begin
      Application.ProcessMessages;
      if rootNode.Parent = nil then
      begin
        clonedRootNode := FindNodeWithText(frmMain.TreeViewCommit.Items.GetFirstNode, frmMain.RuleNames.Values[rootNode.Text]);
        sourceZoneNode := FindChildNode(rootNode, 'Tag');
        if Assigned(sourceZoneNode) then
        begin
          clonedSourceZoneNode := FindChildNode(clonedRootNode, 'Tag');
          zoneRefNode := sourceZoneNode.getFirstChild;
          while Assigned(zoneRefNode) do
          begin
            if Assigned(clonedSourceZoneNode) then
            begin
              frmMain.TreeViewCommit.Items.AddChild(clonedSourceZoneNode, zoneRefNode.Text);
            end;
            zoneRefNode := zoneRefNode.getNextSibling;
          end;
        end;
      end;
    end;
  finally
    frmMain.TreeViewCommit.Items.EndUpdate;
  end;
end;

procedure ModifyValueListEditorValues(ValueListEditor: TValueListEditor; Operator: string);
var
  i: Integer;
  key, value: string;
begin
  if (Operator = 'lower') or (Operator = 'upper') then
  begin
    for i := 1 to ValueListEditor.RowCount - 1 do
    begin
      key := ValueListEditor.Keys[i];
      value := ValueListEditor.Values[key];
      if (key <> '') and (value <> '') then
      begin
        if Operator = 'lower' then
          ValueListEditor.Values[key] := LowerCase(value)
        else if Operator = 'upper' then
          ValueListEditor.Values[key] := UpperCase(value);
      end;
    end;
  end
end;

procedure CreateAddresses;
var
  i: Integer;
  typix, key, name, value: string;
begin
  if frmMain.EditPanorama.Text <> '' then CommandString := 'set device-group ' + frmMain.EditPanorama.Text +  ' address "' else CommandString := 'set address "';
    for i := 1 to frmMain.NetworkObjects.RowCount - 1 do
    begin
      key := frmMain.NetworkObjects.Keys[i];
      name := frmMain.NetworkObjects.Values[key];
      value := frmMain.NetworkObjectsValues.Values[Key];
      if pos('-', Value) > 0 then Typix := ' ip-range ' else typix := ' ip-netmask ';
      if (key <> '') and (value <> '') then
        begin
          frmMain.MemoPaloCLI.Lines.Add(CommandString + Name + '"' + Typix + Value);
        end;
    end
end;

procedure CreateServices;
var
  i: Integer;
  typix, key, name, value: string;
begin
  if frmMain.EditPanorama.Text <> '' then CommandString := 'set device-group ' + frmMain.EditPanorama.Text +  ' service "' else CommandString := 'set service "';
    for i := 1 to frmMain.Services.RowCount - 1 do
    begin
      key := frmMain.Services.Keys[i];
      name := frmMain.Services.Values[key];
      value := frmMain.ServicesValues.Values[Key];
      Typix := ' protocol ';
      if (key <> '') and (value <> '') then
        begin
          frmMain.MemoPaloCLI.Lines.Add(CommandString + Name + '"' + Typix + Value);
        end;
    end
end;

procedure CreateUrls;
var
  i: Integer;
  typix, key, name, value: string;
begin
  if frmMain.EditPanorama.Text <> '' then CommandString := 'set device-group ' + frmMain.EditPanorama.Text +  ' profiles custom-url-category "' else CommandString := 'set profiles custom-url-category "';
    for i := 1 to frmMain.URLEntries.RowCount - 1 do
    begin
      key := frmMain.URLEntries.Keys[i];
      name := frmMain.URLEntries.Values[key];
      name := trim(name);
      value := frmMain.URLEntriesValues.Values[Key];
      Typix := ' type "URL List" list ';
      if (key <> '') and (value <> '') then
        begin
          frmMain.MemoPaloCLI.Lines.Add(CommandString + Name + '"' + Typix + Value);
        end;
    end
end;

procedure CreateTags;
var
  i: Integer;
  value: string;
begin
  if frmMain.EditPanorama.Text <> '' then CommandString := 'set device-group ' + frmMain.EditPanorama.Text +  ' tag "' else CommandString := 'set tag "';
    for i := 0 to taglist.Count - 1 do
    begin
      value := taglist[i];
      frmMain.MemoPaloCLI.Lines.Add(CommandString + value + '"');
    end
end;

procedure TrimToMax(ValueListEditor: TValueListEditor);
var
  i: Integer;
  key, value: string;
  MaxNumber : String;
  Max : Integer;
begin
  MaxNumber := ValueListEditor.TitleCaptions[1];
  MaxNumber := Copy(MaxNumber, Pos('(', MaxNumber) + 1, 2);
  Max := StrToInt(MaxNumber);
    for i := 1 to ValueListEditor.RowCount - 1 do
      begin
        key := ValueListEditor.Keys[i];
        value := ValueListEditor.Values[key];
        if (key <> '') and (value <> '') then
          begin
            if Length(value) > Max then
              begin
                SetLength(value, Max);
                ValueListEditor.Values[key] := value;
              end;
          end;
      end;
end;

procedure UpdateGroupsList(SearchString: string; GroupList: TStringList; Addresses: TValueListEditor);
var
  Node: TTreeNode;
  function SearchNode(Node: TTreeNode): Boolean;
  var
    ChildNode: TTreeNode;
  begin
    Result := False;
    while Assigned(Node) and not Result do
    begin
      if Node.Text = SearchString then
      begin
        ChildNode := Node.GetFirstChild;
        while Assigned(ChildNode) do
        begin
          GroupList.Add(ChildNode.Text);
          ChildNode := ChildNode.GetNextSibling;
        end;
        Result := True;
      end
      else
      begin
        Result := SearchNode(Node.GetFirstChild);
        Node := Node.GetNextSibling;
      end;
    end;
  end;
var
  i: Integer;
  Value: string;
begin
  GroupList.Clear;
  SearchNode(frmMain.TreeViewPolicies.Items.GetFirstNode);
 for i := 0 to GroupList.Count - 1 do
  begin
    Value := Addresses.Values[GroupList[i]];
    if Value <> '' then
    begin
      GroupList.Strings[i] := Value;
    end;
  end;
end;

procedure CreateAddressGroups;
var
  i, j: Integer;
  key, value: string;
  members : String;
begin
  if frmMain.EditPanorama.Text <> '' then CommandString := 'set device-group ' + frmMain.EditPanorama.Text +  ' address-group "' else CommandString := 'set address-group "';
    for i := 1 to frmMain.NetworkObjectsGroup.RowCount - 1 do
      begin
        members := '';
        key := '';
        value := '';
        key := frmMain.NetworkObjectsGroup.Keys[i];
        value := frmMain.NetworkObjectsGroup.Values[key];
        value := value + '"';
        UpdateGroupsList(key, GroupList, frmMain.NetworkObjects);
        for j := 0 to GroupList.Count - 1 do
          begin
            members := members + ' ' + GroupList.Strings[j];
          end;
        members := ' [' + members + ' ]';
        frmMain.MemoPaloCLI.Lines.Add(CommandString + value + members);
      end;
end;

procedure CreateServiceGroups;
var
  i, j: Integer;
  key, value: string;
  members : String;
begin
  if frmMain.EditPanorama.Text <> '' then CommandString := 'set device-group ' + frmMain.EditPanorama.Text +  ' service-group "' else CommandString := 'set service-group "';
    for i := 1 to frmMain.ServicesGroup.RowCount - 1 do
      begin
        members := '';
        key := '';
        value := '';
        key := frmMain.ServicesGroup.Keys[i];
        value := frmMain.ServicesGroup.Values[key];
        value := value + '"';
        UpdateGroupsList(key, GroupList, frmMain.Services);
        for j := 0 to GroupList.Count - 1 do
          begin
            members := members + ' ' + GroupList.Strings[j];
          end;
        members := ' [' + members + ' ]';
        frmMain.MemoPaloCLI.Lines.Add(CommandString + value + members);
      end;
end;

procedure CreateURLGroups;
var
  i, j: Integer;
  key, value: string;
  members : String;
begin
  if frmMain.EditPanorama.Text <> '' then CommandString := 'set device-group ' + frmMain.EditPanorama.Text +  ' profiles custom-url-category "' else CommandString := 'set profiles custom-url-category "';
    for i := 1 to frmMain.URLEntriesGroup.RowCount - 1 do
      begin
        members := '';
        key := '';
        value := '';
        key := frmMain.URLEntriesGroup.Keys[i];
        value := frmMain.URLEntriesGroup.Values[key];
        value := value + '"';
        UpdateGroupsList(key, GroupList, frmMain.URLEntries);
        for j := 0 to GroupList.Count - 1 do
          begin
            members := members + ' ' + GroupList.Strings[j];
          end;
        members := ' [' + members + ' ]';
        frmMain.MemoPaloCLI.Lines.Add(CommandString + value + members);
      end;
end;

procedure MoveMatchingItemsToApplications(TreeView: TTreeView; const SearchWord: string);
var
  RootNode, TypeNode, ItemNode, ApplicationsNode, ServicesNode: TTreeNode;
  i, j: Integer;
  FoundApplications: Boolean;
begin
  // Iterate through all root nodes
  for i := 0 to TreeView.Items.Count - 1 do
  begin
    RootNode := TreeView.Items[i];
    // Check if RootNode is valid and has children
    if (RootNode <> nil) and (RootNode.HasChildren) then
    begin
      ServicesNode := nil;
      ApplicationsNode := nil;
      FoundApplications := False;
      // Iterate through type nodes under the root to find 'Services' and 'Applications'
      for j := 0 to RootNode.Count - 1 do
      begin
        TypeNode := RootNode.Item[j];
        if TypeNode.Text = 'DestinationPorts' then
          ServicesNode := TypeNode;

        if TypeNode.Text = 'Applications' then
        begin
          ApplicationsNode := TypeNode;
          FoundApplications := True;
        end;
      end;
      // If 'Services' node is found but 'Applications' is not, create 'Applications'
      if (ServicesNode <> nil) and not FoundApplications then
        ApplicationsNode := TreeView.Items.AddChild(RootNode, 'Applications');
      // If both 'Services' and 'Applications' nodes are found or created
      if (ServicesNode <> nil) and (ApplicationsNode <> nil) then
      begin
        // Iterate through item nodes under 'Services'
        for j := ServicesNode.Count - 1 downto 0 do
        begin
          ItemNode := ServicesNode.Item[j];
          // Check if the item node matches the search word
          if (ItemNode.Text = SearchWord) then
          begin
            // Move the item node to 'Applications'
            ItemNode.MoveTo(ApplicationsNode, naAddChild);
          end;
        end;
      end;
    end;
  end;
end;

end.
