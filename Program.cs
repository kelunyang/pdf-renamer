using iText.Kernel.Pdf;
using Microsoft.VisualBasic.FileIO;
using System.Text.RegularExpressions;
using PdfDocument = iText.Kernel.Pdf.PdfDocument;
using ReadDocument = UglyToad.PdfPig.PdfDocument;
using UglyToad.PdfPig.Content;

namespace pdfRenamer
  {
  class Program
  {
    static void Main(string[] args)
    {
      List<rule> ruleItems = new List<rule>();
      Console.WriteLine("歡迎使用PDF切割&重新命名小工具，本工具可以幫你把連續的PDF切割成小檔案，並根據你提供的搜尋原則重新命名／加密這些檔案");
      string pdfName = inputHelper("請提供要切割的檔案名稱，如果你的檔案已經切割好了，直接按Enter跳過這一步", true);
      string location = "";
      string oriMeta = "";
      if(pdfName != "") {
        location = inputHelper("你要把切割好的PDF檔案放在哪裡？", true);
        if(File.Exists(pdfName)) {
          Console.WriteLine("PDF檔案...已確認！");
          string? pageStr = inputHelper("你要幾頁切割為一個檔案？", false);
          int numPage = int.Parse(pageStr);
          if(!Directory.Exists(location)) {
            Directory.CreateDirectory(location);
          }
          Console.WriteLine("輸出資料夾...已確認！");
          using (ReadDocument document = ReadDocument.Open(pdfName))
          {
            var di = document.Information.DocumentInformationDictionary;
            if (di != null)
            {
              foreach (var item in di.Data)
              {
                oriMeta += item.Value.ToString();
              }
            }
          }
          using (var pdfDoc = new PdfDocument(new PdfReader(pdfName))) {
            FileInfo file = new FileInfo(pdfName);
            int numberOfPages = pdfDoc.GetNumberOfPages();
            Console.WriteLine("PDF檔案...一共有" + numberOfPages + "頁，你最後會得到" + (numberOfPages/numPage) + "個檔案");
            for (int i = 0; i < numberOfPages; i+=numPage) {
              int pageLimit = i + numPage <= numberOfPages ? i + numPage : numberOfPages;
              string exportName = file.Name.Substring(0, file.Name.LastIndexOf(".")) + "-";
              string newPdfFileName = string.Format(exportName + "{0}.pdf", pageLimit);
              string pageFilePath = Path.Combine(location, newPdfFileName);
              PdfDocument pdf = new PdfDocument(new PdfWriter(new FileStream(pageFilePath, FileMode.Create, FileAccess.Write)));
              /*using (PdfWriter writer = new PdfWriter(new FileStream(pageFilePath, FileMode.Create, FileAccess.Write))) {
                using (var pdf = new PdfDocument(writer)) {*/
                  Console.WriteLine("正在輸出...第" + i + "到" + pageLimit + "頁，檔名：" + newPdfFileName);
                  pdfDoc.CopyPagesTo(pageFrom: (i+1), pageTo: pageLimit, toDocument: pdf, insertBeforePage: 1);
                /*}
              }*/
              pdf.Close();
            }
          }
        }
      }
      string? csvName = inputHelper("你要PDF重新命名規則CSV檔案放在哪裡（請務必使用本工具附帶的規則檔CSV）？", false);
      bool encryptEnable = inputHelper("是否要啟動PDF更名加密？（要啟動請輸入yes，不要就按Enter跳過）", true) != "";
      string? defaultUserPass = "";
      string? defaultOwnerPass = "";
      if(encryptEnable) {
        defaultUserPass = inputHelper("預設的「開啟PDF」密碼？（你可以設定在CSV裡，這個值只有在CSV檔的「開啟密碼」欄位為空下才會啟動）", false);
        defaultOwnerPass = inputHelper("預設的「編輯PDF」密碼？（你可以設定在CSV裡，這個值只有在CSV檔的「編輯密碼」欄位為空下才會啟動）", false);
      }
      using (TextFieldParser textFieldParser = new TextFieldParser(csvName))
      {
        textFieldParser.TextFieldType = FieldType.Delimited;
        textFieldParser.SetDelimiters(",");
        int ruleCount = 0;
        while (!textFieldParser.EndOfData)
        {
          string[]? rows = textFieldParser.ReadFields();
          if(ruleCount > 0) {
            if(rows != null) {
              string userPass = rows[4] == "" ? defaultUserPass : rows[4];
              string ownerPass = rows[5] == "" ? defaultOwnerPass : rows[5];
              ruleItems.Add(new rule(rows[0], rows[1], rows[2], rows[3], userPass, ownerPass, rows[4] == "" , rows[5] == "", encryptEnable));
            }
          }
          ruleCount++;
        }
      }
      int searched = 0;
      int matched = 0;
      Console.WriteLine("比對更名規則已載入..." + ruleItems.Count() + "條");
      string? searchLocation = inputHelper("你要搜尋的PDF檔案放在哪個資料夾裡？" + (location != "" ? "（按下Enter會載入你剛剛分割的資料夾位置）" : ""), location != "");
      if(searchLocation == "") { searchLocation = location; }
      if(Directory.Exists(searchLocation)) {
        DirectoryInfo di = new DirectoryInfo(searchLocation);
        FileInfo[] pdfs = di.GetFiles("*.pdf");
        Console.WriteLine("PDF資料夾已載入...找到" + pdfs.Length + "個PDF");
        string metadata = "";
        if(oriMeta != "") {
          metadata = oriMeta;
        }
        for(int i=0; i<pdfs.Length; i++) {
          string content = "";//這邊應該要確認路徑，以防同名的檔案已經被換置了
          using (ReadDocument document = ReadDocument.Open(pdfs[i].FullName))
          {
            var dInfo = document.Information.DocumentInformationDictionary;
            if (dInfo != null)
            {
              foreach (var item in dInfo.Data)
              {
                metadata += item.Value.ToString();
              }
            }
            foreach (Page page in document.GetPages())
            {
              if(i== 0)
              {
                if(page.Number == 1)
                {
                  Console.WriteLine("以下是第一個檔案第一頁的預覽文字（有些PDF可能會有亂碼，建議按照預覽文字調整你的搜尋結果）：");
                  Console.Write(page.Text);
                  Console.WriteLine("確認更名的話請直接按下Enter，有誤請按Ctrl+C中斷程式（然後你可能要修改更名CSV）");
                  Console.ReadLine();
                }
              }
              content += page.Text;
            }
          }
          searched++;
          Console.WriteLine("檔案[" + pdfs[i].Name + "]讀取完成，開始匹配搜尋規則");
          for(int r=0; r<ruleItems.Count(); r++) {
            if(ruleItems[r].targetType == "內容") {
              MatchCollection matches = ruleItems[r].ruleFrom.Matches(content);
              if(matches.Count() == ruleItems[r].occurrenceMatch) {
                matched += fileRenamer(ruleItems[r], pdfs[i], encryptEnable, searchLocation) ? 1 : 0;
              }
            } else if(ruleItems[r].targetType == "中介") {
              MatchCollection matches = ruleItems[r].ruleFrom.Matches(metadata);
              if(matches.Count() == ruleItems[r].occurrenceMatch) {
                matched += fileRenamer(ruleItems[r], pdfs[i], encryptEnable, searchLocation) ? 1 : 0;
              }
            }
          }
          GC.Collect();
        }
      }
      Console.WriteLine("PDF作業完成，一共搜尋了" + searched + "個PDF檔案，更名了" + matched + "個檔案！按任意鍵結束");
      Console.ReadLine();
    }
    static bool fileRenamer(rule ruleItem, FileInfo pdf, bool encrypted, string searchLocation) {
      if(!File.Exists(searchLocation + "\\" + ruleItem.nameTo + ".pdf")) {
        if(encrypted) {
          PdfReader pdfReader = new PdfReader(pdf.FullName);
          WriterProperties writerProperties = new WriterProperties();
          writerProperties.SetStandardEncryption(ruleItem.userPass, ruleItem.ownerPass, EncryptionConstants.ALLOW_PRINTING, EncryptionConstants.ENCRYPTION_AES_256);
          PdfWriter pdfWriter = new PdfWriter(new FileStream( searchLocation + "\\" + ruleItem.nameTo + ".pdf", FileMode.Create), writerProperties);
          PdfDocument pdfDocument = new PdfDocument(pdfReader, pdfWriter);
          pdfDocument.Close();
          File.Delete(pdf.FullName);
          Console.WriteLine("找到檔案[" + pdf.Name + "]的" + ruleItem.targetType + "符合規則[" + ruleItem.ruleFrom.ToString() + "]，檔案加密更名完成");
        } else {
          File.Move(pdf.FullName, searchLocation + "\\" + ruleItem.nameTo + ".pdf");
          Console.WriteLine("找到檔案[" + pdf.Name + "]的" + ruleItem.targetType + "符合規則[" + ruleItem.ruleFrom.ToString() + "]，檔案更名完成");
        }
        return true;
      } else {
        Console.WriteLine("檔案[" + ruleItem.nameTo + ".pdf]已存在，跳過重新更名");
        return false;
      }
    }
    static string inputHelper(string msg, bool nullable) {
      string? value = "";
      int count = 0;
      string helper = "";
      if(nullable) {
        helper = "[按下Enter可直接跳過本題]";
      } else {
        if(count > 0) {
          helper = "[你剛剛沒有輸入任何值]";
        } else {
          helper = "[本題不可跳過]";
        }
      }
      while(value == "") {
        Console.WriteLine(msg + helper);
        value = Console.ReadLine();
        if(value == null) {
          value = nullable ? " " : "";
        }
        if(nullable) {
          value = value == "" ? " " : value;
        }
        count++;
      }
      return value == " " ? "" : value;
    }
  }
}

class rule {
  public Regex ruleFrom;
  public string nameTo;
  public string targetType = "內容";
  public int occurrenceMatch;
  public byte[] userPass;
  public byte[] ownerPass;
  public rule(string rule, string name, string type, string occurrence, string userPass, string ownerPass, bool userPassset, bool ownerPassset, bool encryptEnable) {
    this.ruleFrom = new Regex(rule);
    this.nameTo = name;
    this.targetType = type;
    this.occurrenceMatch = int.Parse(occurrence);
    this.userPass = System.Text.Encoding.Default.GetBytes(userPass);
    this.ownerPass = System.Text.Encoding.Default.GetBytes(ownerPass);
    string userPrompt = !userPassset ? "開啟密碼已設定" : "開啟密碼採用預設密碼";
    string ownerPrompt = !ownerPassset ? "編輯密碼已設定" : "編輯密碼採用預設密碼";
    string passPrompt = encryptEnable ? "（" + userPrompt + "／" + ownerPrompt + "）" : "";
    Console.WriteLine("找：" + rule + "的" + type +"，重複出現" + occurrence + "次，更名為：" + name + passPrompt);
  }
}