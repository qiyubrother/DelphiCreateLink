unit UnitMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,
  {这三个单元是必须的}
  ComObj, ActiveX, ShlObj;

type
  TForm1 = class(TForm)
    Button1: TButton;
    procedure CreateLink(ProgramPath, ProgramArg, LinkPath, Descr: String);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;
const
  maxPath = 200; // 定义最大字符串数组长度

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var

  tmp: array [0..maxPath] of Char;
  WinDir: string;
  pitem:PITEMIDLIST;
  usrDeskTopPath: string;
begin

  //获取当前用户桌面的位置
  SHGetSpecialFolderLocation(self.Handle, CSIDL_DESKTOP, pitem);
  setlength(usrDeskTopPath, maxPath);
  shGetPathFromIDList(pitem, PWideChar(usrDeskTopPath));
  usrDeskTopPath := String(PWideChar(usrDeskTopPath));

  // 创建快捷方式
  CreateLink(
    ParamStr(0),                                       // 应用程序完整路径
    '-22 -dd xx="aa"',                                 // 传给应用程序的参数
    usrDeskTopPath + '\' + Application.Title + '.lnk', // 快捷方式完整路径
    'Application.Title'                                // 备注
  );
end;
procedure TForm1.CreateLink(ProgramPath, ProgramArg, LinkPath, Descr: String);
var
  AnObj: IUnknown;
  ShellLink: IShellLink;
  AFile: IPersistFile;
  FileName: WideString;
begin
  if UpperCase(ExtractFileExt(LinkPath)) <> '.LNK' then //检查扩展名是否正确
  begin
    raise Exception.Create('快捷方式的扩展名必须是 ′′LNK′′!');
    //若不是则产生异常
  end;
try
  OleInitialize(nil);//初始化OLE库，在使用OLE函数前必须调用初始化
  AnObj := CreateComObject(CLSID_ShellLink); //根据给定的ClassID生成
  //一个COM对象，此处是快捷方式
  ShellLink := AnObj as IShellLink;//强制转换为快捷方式接口
  AFile := AnObj as IPersistFile;//强制转换为文件接口
  //设置快捷方式属性，此处只设置了几个常用的属性
  ShellLink.SetPath(PChar(ProgramPath)); // 快捷方式的目标文件，一般为可执行文件
  ShellLink.SetArguments(PChar(ProgramArg));// 目标文件参数
  ShellLink.SetWorkingDirectory(PChar(ExtractFilePath(ProgramPath)));//目标文件的工作目录
  ShellLink.SetDescription(PChar(Descr));// 对目标文件的描述
  FileName := LinkPath;//把文件名转换为WideString类型
  AFile.Save(PWChar(FileName), False);//保存快捷方式
finally
　OleUninitialize;//关闭OLE库，此函数必须与OleInitialize成对调用
end;

end;

end.
