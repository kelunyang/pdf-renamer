using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using Microsoft.VisualBasic.FileIO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.IO;
namespace pdfRenamer
  {
  class Program
  {
    static void Main(string[] args)
    {
      List<rule> ruleItems = new List<rule>();
      Console.WriteLine("歡迎使用PDF切割&重新命名小工具，本工具可以幫你把連續的PDF切割成小檔案，並根據你提供的搜尋原則重新命名這些檔案");
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
          using (var pdfDoc = new PdfDocument(new PdfReader(pdfName))) {
            PdfDictionary trailer = pdfDoc.GetTrailer();
            PdfDictionary metadataInfo = trailer.GetAsDictionary(PdfName.Info);
            var keys = metadataInfo.KeySet();
            foreach (var key in keys)
            {
              var value = ((PdfString)metadataInfo.Get(key)).GetValue();
              oriMeta += " " + value;
            }
            FileInfo file = new FileInfo(pdfName);
            int numberOfPages = pdfDoc.GetNumberOfPages();
            Console.WriteLine("PDF檔案...一共有" + numberOfPages + "頁，你最後會得到" + (numberOfPages/numPage) + "個檔案");
            for (int i = 0; i < numberOfPages; i+=numPage) {
              int pageLimit = i + numPage <= numberOfPages ? i + numPage : numberOfPages;
              string exportName = file.Name.Substring(0, file.Name.LastIndexOf(".")) + "-";
              string newPdfFileName = string.Format(exportName + "{0}.pdf", pageLimit);
              string pageFilePath = Path.Combine(location, newPdfFileName);
              using (PdfWriter writer = new PdfWriter(pageFilePath)) {
                using (var pdf = new PdfDocument(writer)) {
                  Console.WriteLine("正在輸出...第" + i + "到" + pageLimit + "頁，檔名：" + newPdfFileName);
                  pdfDoc.CopyPagesTo(pageFrom: (i+1), pageTo: pageLimit, toDocument: pdf, insertBeforePage: 1);
                }
              }
            }
          }
        }
      }
      string? csvName = inputHelper("你要PDF重新命名規則CSV檔案放在哪裡（請務必使用本工具附帶的規則檔CSV）？", false);
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
              ruleItems.Add(new rule(rows[0], rows[1], rows[2], rows[3]));
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
          string content = "";
          PdfReader pdfReader = new PdfReader(pdfs[i].FullName);
          PdfDocument pdfDoc = new PdfDocument(pdfReader);
          int pages = pdfDoc.GetNumberOfPages();
          for(int k=1; k<=pages; k++) {
            ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
            content += " " + PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(k), strategy);
          }
          if(oriMeta == "") {
            metadata = "";
            PdfDictionary trailer = pdfDoc.GetTrailer();
            PdfDictionary metadataInfo = trailer.GetAsDictionary(PdfName.Info);
            var keys = metadataInfo.KeySet();
            foreach (var key in keys)
            {
              var value = ((PdfString)metadataInfo.Get(key)).GetValue();
              metadata += " " + value;
            }
          }
          pdfDoc.Close();
          pdfReader.Close();
          searched++;
          Console.WriteLine("檔案[" + pdfs[i].Name + "]讀取完成，開始匹配搜尋規則");
          for(int r=0; r<ruleItems.Count(); r++) {
            if(ruleItems[r].targetType == "內容") {
              MatchCollection matches = ruleItems[r].ruleFrom.Matches(content);
              if(matches.Count() == ruleItems[r].occurrenceMatch) {
                if(!File.Exists(searchLocation + "\\" + ruleItems[r].nameTo + ".pdf")) {
                  File.Move(pdfs[i].FullName, searchLocation + "\\" + ruleItems[r].nameTo + ".pdf");
                  Console.WriteLine("找到檔案[" + pdfs[i].Name + "]的" + ruleItems[r].targetType + "符合規則[" + ruleItems[r].ruleFrom.ToString() + "]，檔案更名完成");
                  matched++;
                } else {
                  Console.WriteLine("檔案[" + ruleItems[r].nameTo + ".pdf]已存在，跳過重新更名");
                }
              }
            } else if(ruleItems[r].targetType == "中介") {
              MatchCollection matches = ruleItems[r].ruleFrom.Matches(metadata);
              if(matches.Count() == ruleItems[r].occurrenceMatch) {
                if(!File.Exists(searchLocation + "\\" + ruleItems[r].nameTo + ".pdf")) {
                  File.Move(pdfs[i].FullName, searchLocation + "\\" + ruleItems[r].nameTo + ".pdf");
                  Console.WriteLine("找到檔案[" + pdfs[i].Name + "]的" + ruleItems[r].targetType + "符合規則[" + ruleItems[r].ruleFrom.ToString() + "]，檔案更名完成");
                  matched++;
                } else {
                  Console.WriteLine("檔案[" + ruleItems[r].nameTo + ".pdf]已存在，跳過重新更名");
                }
              }
            }
          }
          GC.Collect();
        }
      }
      Console.WriteLine("PDF作業完成，一共搜尋了" + searched + "個PDF檔案，更名了" + matched + "個檔案！按任意鍵結束");
      Console.ReadLine();
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
  public rule(string rule, string name, string type, string occurrence) {
    this.ruleFrom = new Regex(rule);
    this.nameTo = name;
    this.targetType = type;
    this.occurrenceMatch = int.Parse(occurrence);
    Console.WriteLine("找：" + rule + "的" + type +"，重複出現" + occurrence + "次，更名為：" + name);
  }
}