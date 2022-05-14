using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using dotnetMVC_itextsharp.Models;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Http;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using System.Text;

namespace dotnetMVC_itextsharp.Controllers
{
    public class HomeController : Controller
    {
        private IWebHostEnvironment _hostEnvironment;
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger,IWebHostEnvironment environment)
        {
            _logger = logger;
            _hostEnvironment = environment;
        }

        public IActionResult Index()
        {
            ViewData["path"] = _hostEnvironment.WebRootPath + "\\font\\KAIU.TTF";
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }
        
        public IActionResult Preview(){
            Document doc = new Document(PageSize.A4);
            MemoryStream ms = new MemoryStream();
            PdfWriter.GetInstance(doc, ms).CloseStream = false;
            
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            BaseFont bfChinese = BaseFont.CreateFont(_hostEnvironment.WebRootPath + "\\font\\KAIU.TTF", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            BaseFont chBaseFont = BaseFont.CreateFont(_hostEnvironment.WebRootPath + "\\font\\KAIU.TTF", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            
            Font ChFont = new Font(bfChinese, 12);
            Font CbFont = new Font(chBaseFont, 12,5);
            //new Font(chBaseFont, 12,2);

            // Font ChFont_green = new Font(bfChinese, 40, Font.NORMAL, BaseColor.GREEN);
            // Font ChFont_msg = new Font(bfChinese, 12, Font.ITALIC, BaseColor.RED);
                       
            doc.Open();
            Paragraph title = new Paragraph("This Title Area",ChFont);
            title.Alignment = Element.ALIGN_CENTER;
            title.SpacingBefore =50;
            title.SpacingAfter =50;
            title.IndentationLeft=50;
            title.IndentationRight=50;
            doc.Add(title);
            doc.Add(new Paragraph("This Sub Title Area",ChFont));

            
            Chunk chunk = new Chunk("測試底線文字", CbFont);
            chunk.SetUnderline(0.2f, -2f);


            

            doc.Add(chunk);
            doc.Add(new Paragraph("\n"));
            doc.Add(new Paragraph("中文測試",ChFont));
            doc.Add(new Paragraph("\n"));

            //
            PdfPTable pt = new PdfPTable(3);
            pt.AddCell(new PdfPCell(new Phrase($" 所有第三欄合併 ",ChFont)){Colspan=3});
            for(int i = 1 ; i <= 3;++i){
                for(int j = 1 ; j <= 3;++j){
                    Phrase text =  new Phrase($"line{i},cell{j}");
                    PdfPCell cell= new PdfPCell(text);
                    if(i != 1 && j == 3 ) { continue;}
                    if(j == 3 ){
                        pt.AddCell(new PdfPCell(new Phrase($"cell{j}")){Rowspan=3});
                        continue;
                    }
                    //font
                    pt.AddCell(cell);
                }
            }
            doc.Add(new Paragraph(){pt});
            doc.Add(new Paragraph("\n"){});
            doc.Add(new Paragraph("\n"){});

            //
            pt = new PdfPTable(3);
            pt.AddCell(new PdfPCell(new Phrase($" 第三欄 ",ChFont)){Colspan=3});
            for(int i = 1 ; i <= 3;++i){
                for(int j = 1 ; j <= 3;++j){
                    Phrase text =  new Phrase($"line{i},cell{j}");
                    PdfPCell cell= new PdfPCell(text);
                    pt.AddCell(cell);
                }
            }
            doc.Add(new Paragraph(){pt});



            doc.Close();
            
            ms.Seek(0, SeekOrigin.Begin);
            return new FileStreamResult(ms, "application/pdf");
        }



        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
