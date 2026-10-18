using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;
using System.Xml.Linq;

[assembly: System.Reflection.AssemblyTitle("Word文本比对器")]
[assembly: System.Reflection.AssemblyProduct("Word文本比对器")]
namespace WordCompare {
 static class Program {
  [STAThread] static void Main(string[] args) {
   Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
   if(args.Length==2 && args[0]=="--self-test") { try { Engine.Test(args[1]); } catch(Exception ex) { File.WriteAllText(Path.Combine(args[1],"test-error.txt"),ex.ToString()); Environment.ExitCode=1; } return; }
   Application.Run(new MainForm());
  }
 }
 static class Engine {
  public static void Release(object o) { if(o!=null && Marshal.IsComObject(o)) Marshal.FinalReleaseComObject(o); }
  public static string Compare(string original,string revised,string folder,string author,bool formatting) {
   dynamic app=null,docs=null,a=null,b=null,result=null;
   try {
    Type t=Type.GetTypeFromProgID("Word.Application"); if(t==null) throw new Exception("未检测到 Microsoft Word，请先安装桌面版 Word。");
    app=Activator.CreateInstance(t); app.Visible=false; app.DisplayAlerts=0; app.AutomationSecurity=3;
    docs=app.Documents;
    a=docs.Open(FileName:original,ConfirmConversions:false,ReadOnly:true,AddToRecentFiles:false,Visible:false);
    b=docs.Open(FileName:revised,ConfirmConversions:false,ReadOnly:true,AddToRecentFiles:false,Visible:false);
    result=app.CompareDocuments(OriginalDocument:a,RevisedDocument:b,Destination:2,Granularity:1,CompareFormatting:formatting,CompareCaseChanges:true,CompareWhitespace:true,CompareTables:true,CompareHeaders:true,CompareFootnotes:true,CompareTextboxes:true,CompareFields:true,CompareComments:true,CompareMoves:true,RevisedAuthor:author,IgnoreAllComparisonWarnings:true);
    string stem=Path.GetFileNameWithoutExtension(revised)+"_对比稿_"+DateTime.Now.ToString("yyyyMMdd_HHmmss");
    string path=Path.Combine(folder,stem+".docx"); int n=2;
    // Reserve the output name atomically so an existing report is never replaced.
    while(true) { try { using(File.Open(path,FileMode.CreateNew,FileAccess.Write,FileShare.None)) {} break; } catch(IOException) { if(!File.Exists(path)) throw; path=Path.Combine(folder,stem+"_"+(n++)+".docx"); } }
    try { result.SaveAs2(FileName:path,FileFormat:16,AddToRecentFiles:false); } catch { if(File.Exists(path) && new FileInfo(path).Length==0) File.Delete(path); throw; }
    return path;
   } finally {
    try { if((object)result!=null) result.Close(0); } catch {} Release((object)result);
    try { if((object)b!=null) b.Close(0); } catch {} Release((object)b);
    try { if((object)a!=null) a.Close(0); } catch {} Release((object)a); Release((object)docs);
    try { if((object)app!=null) app.Quit(0); } catch {} Release((object)app);
   }
  }
  public static void Test(string folder) {
   Directory.CreateDirectory(folder); dynamic app=null,docs=null,d=null;
   try { app=Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application")); app.Visible=false; app.DisplayAlerts=0; docs=app.Documents;
    foreach(string name in new[]{"original","revised"}) { d=docs.Add(); dynamic range=d.Content; range.Text=name=="original"?"采购合同\r付款期限为30日。\r合同金额100元。":"采购合同\r付款期限为60日。\r合同金额100元。\r新增交付条款。"; Release((object)range); d.SaveAs2(Path.Combine(folder,name+".docx"),16); d.Close(0); Release((object)d); d=null; }
   } finally { if((object)d!=null) {try{d.Close(0);}catch{} Release((object)d);} Release((object)docs); if((object)app!=null){try{app.Quit(0);}catch{} Release((object)app);} }
   string output=Compare(Path.Combine(folder,"original.docx"),Path.Combine(folder,"revised.docx"),folder,"测试修订人",false);
   File.WriteAllText(Path.Combine(folder,"test-result.txt"),output);
  }
 }
 class MainForm:Form {
  TextBox original=new TextBox(),revised=new TextBox(),author=new TextBox(),folder=new TextBox();
  CheckBox ignore=new CheckBox(),autoOpen=new CheckBox(); Button run=new Button(),open=new Button(),reveal=new Button(); Label status=new Label();
  TableLayoutPanel body=new TableLayoutPanel(); string output; bool busy;
  string config=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"FrankWordCompare","settings.xml");
  public MainForm() {
   Text="Word文本比对器"; Size=new Size(830,570); MinimumSize=new Size(760,550); StartPosition=FormStartPosition.CenterScreen;
   Font=new Font("Microsoft YaHei UI",10); BackColor=Color.FromArgb(245,247,251); AutoScaleMode=AutoScaleMode.Dpi;
   body.Dock=DockStyle.Fill; body.Padding=new Padding(28); body.ColumnCount=1; body.RowCount=9;
   body.RowStyles.Add(new RowStyle(SizeType.Absolute,44)); body.RowStyles.Add(new RowStyle(SizeType.Absolute,32)); body.RowStyles.Add(new RowStyle(SizeType.Absolute,145)); body.RowStyles.Add(new RowStyle(SizeType.Absolute,44)); body.RowStyles.Add(new RowStyle(SizeType.Absolute,44)); body.RowStyles.Add(new RowStyle(SizeType.Absolute,42)); body.RowStyles.Add(new RowStyle(SizeType.Absolute,54)); body.RowStyles.Add(new RowStyle(SizeType.Absolute,42)); body.RowStyles.Add(new RowStyle(SizeType.Percent,100)); Controls.Add(body);
   body.Controls.Add(new Label{Text="两份文档，一键看清修改",Font=new Font(Font.FontFamily,20,FontStyle.Bold),AutoSize=true});
   body.Controls.Add(new Label{Text="调用本机 Microsoft Word · 文件在本地处理 · 输出原生修订稿",AutoSize=true,ForeColor=Color.DimGray});
   var files=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=3}; files.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,50)); files.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,65)); files.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,50));
   files.Controls.Add(FileBox("01  原稿",original),0,0); var swap=new Button{Text="⇄",Dock=DockStyle.Fill,Margin=new Padding(8,45,8,45)}; swap.Click+=(s,e)=>{string v=original.Text;original.Text=revised.Text;revised.Text=v;}; files.Controls.Add(swap,1,0); files.Controls.Add(FileBox("02  修改稿",revised),2,0); body.Controls.Add(files);
   var settings=new FlowLayoutPanel{Dock=DockStyle.Fill}; ignore.Text="忽略格式变化"; ignore.Checked=true; ignore.AutoSize=true; author.Width=180; author.Text="修订人"; settings.Controls.Add(ignore); settings.Controls.Add(new Label{Text="    修订人",AutoSize=true,Margin=new Padding(3,6,3,3)}); settings.Controls.Add(author); body.Controls.Add(settings);
   var dest=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=3}; dest.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,85));dest.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));dest.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,90));dest.Controls.Add(new Label{Text="保存位置",AutoSize=true,Margin=new Padding(0,7,0,0)});folder.Dock=DockStyle.Fill;dest.Controls.Add(folder);var browse=new Button{Text="选择…",Dock=DockStyle.Fill};browse.Click+=(s,e)=>{using(var d=new FolderBrowserDialog()){if(d.ShowDialog()==DialogResult.OK)folder.Text=d.SelectedPath;}};dest.Controls.Add(browse);body.Controls.Add(dest);
   autoOpen.Text="完成后自动打开对比稿";autoOpen.AutoSize=true;autoOpen.Checked=true;body.Controls.Add(autoOpen);
   run.Text="生成对比稿";run.Dock=DockStyle.Fill;run.BackColor=Color.FromArgb(37,88,210);run.ForeColor=Color.White;run.FlatStyle=FlatStyle.Flat;run.Click+=Start;body.Controls.Add(run);
   status.Text="将文件拖入上方区域，或点击选择文件。保存位置留空时使用修改稿文件夹。";status.Dock=DockStyle.Fill;body.Controls.Add(status);
   var actions=new FlowLayoutPanel{Dock=DockStyle.Fill};open.Text="打开对比稿";reveal.Text="打开所在文件夹";open.AutoSize=reveal.AutoSize=true;open.Visible=reveal.Visible=false;open.Click+=(s,e)=>Launch(output);reveal.Click+=(s,e)=>Launch(Path.GetDirectoryName(output));actions.Controls.Add(open);actions.Controls.Add(reveal);body.Controls.Add(actions);
   try {var x=XElement.Load(config);author.Text=(string)x.Element("author")??"修订人";ignore.Checked=(bool?)x.Element("ignore")??true;autoOpen.Checked=(bool?)x.Element("autoOpen")??true;folder.Text=(string)x.Element("folder")??"";}catch{}
   FormClosing+=(s,e)=>{if(busy){e.Cancel=true;MessageBox.Show("Word 正在生成对比稿，请完成后再关闭。");}};
  }
  Control FileBox(string title,TextBox field) {
   var box=new GroupBox{Text=title,Dock=DockStyle.Fill,Padding=new Padding(12),AllowDrop=true};field.ReadOnly=true;field.Dock=DockStyle.Top;var select=new Button{Text="点击选择，或拖入 Word 文件",Dock=DockStyle.Bottom,Height=43}; select.Click+=(s,e)=>{using(var d=new OpenFileDialog{Filter="Word 文档|*.docx;*.doc",CheckFileExists=true}){if(d.ShowDialog()==DialogResult.OK)field.Text=d.FileName;}};
   DragEventHandler enter=(s,e)=>{if(e.Data.GetDataPresent(DataFormats.FileDrop))e.Effect=DragDropEffects.Copy;};DragEventHandler drop=(s,e)=>{var f=e.Data.GetData(DataFormats.FileDrop) as string[];if(f!=null && f.Length==1 && Valid(f[0]))field.Text=f[0];else MessageBox.Show("请拖入一个 .docx 或 .doc 文件。");};
   foreach(Control c in new Control[]{box,field,select}){c.AllowDrop=true;c.DragEnter+=enter;c.DragDrop+=drop;}box.Controls.Add(field);box.Controls.Add(select);return box;
  }
  static bool Valid(string p){return File.Exists(p) && (Path.GetExtension(p).Equals(".docx",StringComparison.OrdinalIgnoreCase)||Path.GetExtension(p).Equals(".doc",StringComparison.OrdinalIgnoreCase));}
  void Launch(string path){try{System.Diagnostics.Process.Start(path);}catch(Exception ex){MessageBox.Show("文件已保存，但打开失败："+ex.Message);}}
  void Start(object sender,EventArgs args) {
   if(!Valid(original.Text)||!Valid(revised.Text)){MessageBox.Show("请先选择原稿和修改稿。");return;}if(string.Equals(Path.GetFullPath(original.Text),Path.GetFullPath(revised.Text),StringComparison.OrdinalIgnoreCase)){MessageBox.Show("请选择两份不同的文档。");return;}
   string dir=string.IsNullOrWhiteSpace(folder.Text)?Path.GetDirectoryName(revised.Text):folder.Text.Trim();if(!Directory.Exists(dir)){MessageBox.Show("保存文件夹不存在，请重新选择。");return;}if(string.IsNullOrWhiteSpace(author.Text)){MessageBox.Show("请填写修订人。");return;}
   try{Directory.CreateDirectory(Path.GetDirectoryName(config));new XElement("settings",new XElement("author",author.Text.Trim()),new XElement("ignore",ignore.Checked),new XElement("autoOpen",autoOpen.Checked),new XElement("folder",folder.Text)).Save(config);}catch{}
   string a=original.Text,b=revised.Text,who=author.Text.Trim();bool formatting=!ignore.Checked,shouldOpen=autoOpen.Checked;
   busy=true;body.Enabled=false;status.Text="正在调用 Word 对比，请稍候…";open.Visible=reveal.Visible=false;
   var worker=new Thread(()=>{string result=null;Exception error=null;try{result=Engine.Compare(a,b,dir,who,formatting);}catch(Exception ex){error=ex;}BeginInvoke(new Action(()=>{busy=false;body.Enabled=true;if(error!=null){status.Text="生成失败，请检查文件是否可正常打开。";MessageBox.Show("Word 对比失败："+error.Message+"\n请确认已安装并激活 Microsoft Word，文件未加密且可以正常打开。","生成失败",MessageBoxButtons.OK,MessageBoxIcon.Error);}else{output=result;status.Text="已保存："+Path.GetFileName(output);open.Visible=reveal.Visible=true;if(shouldOpen)Launch(output);}}));});worker.SetApartmentState(ApartmentState.STA);worker.IsBackground=true;worker.Start();
  }
 }
}

