using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;

//Need Nuget
//iTextSharp
//iTextSharp XML

namespace play_itextsharp
{
    class Program
    {
        static void Example_Problem_with_OL_LI_FontFamily_Control_Of_List_Bullets()
        {
            //*****************************************************************************************************************************************************************************
            // ATTACHMENTS  612x792 pixels
            // For this PDF generation we are using
            // iTextSharp HTML to PDF capabilities,
            // not PDFStamper like the other ones
            //
            // ENCOUNTERED ISSUES WITH iTextSharp "HTML to PDF" functionality and NicEdit
            // (1) iTextSharp will not allow me to control font-family for <ol> <li> items no matter what HTML and inline CSS i changed even using registered fonts.
            //     First idea was to not feed the NicEdit widget with <ol> and <li> and just use old school &nbsp; as indentation and you control the ordering of item bullets
            //     whether a. b. c. or I. II. III. or 1. 2. 3.
            //     But then the problem was with sub-indentation within a "list item", so then i decided to go back to straight PDF generation and make my own "NicEdit HTML parser".
            //******************************************************************************************************************************************************************************
            Dictionary<string, string> Request = new Dictionary<string, string>(); //This was originally from form submit in Microsoft MVC web-app
            string attach_pdf = "FINAL_attachments.pdf";

            Byte[] bytes;
            using (MemoryStream ms = new MemoryStream())
            {
                using (Document lang_doc = new Document(PageSize.A4, 5, 5, 15, 5))
                {
                    using (PdfWriter writer = PdfWriter.GetInstance(lang_doc, ms))
                    {
                        lang_doc.SetMargins(50f, 50f, 50f, 50f);
                        lang_doc.Open();

                        //"<head style=\"font-size:12.0pt; font-family:Times; \"></head>" +

                        string myhtml = "<html>" +
                                        "<head><style>" +
                                        "li:before {" +
                                        "   font-family: Courier;" +
                                        "} " +
                                        "li {" +
                                        "   font-family: Courier;" +
                                        "} " +
                                        "ol {" +
                                        "   font-family: Courier;" +
                                        "} " +
                                        "body {" +
                                        "   font-family: Courier;" +
                                        "} " +
                                        "</style></head>" +
                                        "<body style=\"font-size:12.0pt; font-family:Courier; \">";
                        myhtml += "<p align=\"right\" style=\"font-size:12.0pt;font-weight:normal;\">ATTACHMENT</p><br>";
                        myhtml += "<p align=\"left\" style=\"font-size:12.0pt;font-weight:normal;\"><span style=\"text-decoration:underline;\">CONTRACT MODIFICATIONS</span>:</p><br>";
                        myhtml += "<p align=\"left\" style=\"font-size:12.0pt;font-weight:normal;\">&nbsp;&nbsp;&nbsp;&nbsp;Effective <span style=\"text-decoration:underline;\">" + Request["effective_date"] + "</span>, the parties mutually agree that the PROVIDER/CONTRACTOR shall continue to provide the required services with the following modifications:</p><br>";

                        // ENCOUNTERED ISSUES WITH iTextSharp "HTML to PDF" functionality and NicEdit
                        // (1) iTextSharp will not allow me to control font-family for <ol> <li> items no matter what HTML and inline CSS i changed even using registered fonts.
                        //     First idea was to not feed the NicEdit widget with <ol> and <li> and just use old school &nbsp; as indentation and you control the ordering of item bullets
                        //     whether a. b. c. or I. II. III. or 1. 2. 3.
                        //     But then the problem was with sub-indentation within a "list item", so then i decided to go back to straight PDF generation and make my own "NicEdit HTML parser".

                        char workaround_list_index = 'a';

                        if (Request["conmodtype_date_extension"] == "on")
                        {
                            myhtml += "<p align=\"left\" style=\"font-size:12.0pt;margin-left:100;\">&nbsp;&nbsp;&nbsp;&nbsp;" + workaround_list_index + ".&nbsp;&nbsp;" + Request["date_extension"] + "</p>";
                            workaround_list_index++;
                        }

                        if (Request["conmodtype_budget_change"] == "on")
                        {
                            myhtml += "&nbsp;&nbsp;&nbsp;&nbsp;" + workaround_list_index + ".&nbsp;&nbsp;" + Request["budget_change"] + "<br>";
                            workaround_list_index++;
                        }

                        if (Request["conmodtype_rate_schedule_change"] == "on")
                        {
                            myhtml += "&nbsp;&nbsp;&nbsp;&nbsp;" + workaround_list_index + ".&nbsp;&nbsp;" + Request["rate_schedule_change"] + "<br>";
                            workaround_list_index++;
                        }

                        myhtml += "</body></html>";


                        //ok for some reason iTextSharp HTML to PDF did not like <p> tags without any "stuff inside it" like align and syle
                        //So when i do this below it works
                        myhtml = myhtml.Replace("<p>", "<p align=\"left\" style=\"font-size:12.0pt;\">");

                        myhtml = myhtml.Replace("<br>", "<br></br>");

                        StreamWriter quickdump = new StreamWriter("C:\\Users\\mlee\\Desktop\\quickdump.txt");
                        quickdump.WriteLine(myhtml);
                        quickdump.Close();


                        iTextSharp.tool.xml.XMLWorkerFontProvider xmlWorkerFontProvider = new iTextSharp.tool.xml.XMLWorkerFontProvider();


                        //It was a real pain to get fonts working for HTML to PDF, but these 3 websites will help a lot:
                        //http://stackoverflow.com/questions/14959741/itextsharp-font-tag-not-working-with-xmlworker-class
                        //http://stackoverflow.com/questions/19940536/choosing-a-fontproviderimp-in-itextsharp
                        //http://demo.itextsupport.com/xmlworker/itextdoc/flatsite.html

                        /************************
                        bool courier_ok = xmlWorkerFontProvider.IsRegistered("Courier");
                        if (courier_ok)
                            myhtml += "<p>Courier OK</p>";
                        else
                            myhtml += "<p>Courier NOT REGISTERED</p>";

                        bool times_ok = xmlWorkerFontProvider.IsRegistered("Times-Roman");
                        if (times_ok)
                            myhtml += "<p>Times-Roman OK</p>";
                        else
                            myhtml += "<p>Times-Roman NOT REGISTERED</p>";

                        bool times_new_roman_ok = xmlWorkerFontProvider.IsRegistered("Times New Roman");
                        if (times_new_roman_ok)
                            myhtml += "<p>Times New Roman OK</p>";
                        else
                            myhtml += "<p>Times New Roman NOT REGISTERED</p>";

                        foreach (string fontname in xmlWorkerFontProvider.RegisteredFonts)
                            myhtml += "<p>" + fontname + "</p>";
                        *************************/

                        iTextSharp.tool.xml.html.CssAppliersImpl cssAppliers = new iTextSharp.tool.xml.html.CssAppliersImpl(xmlWorkerFontProvider);
                        iTextSharp.tool.xml.css.StyleAttrCSSResolver cssRevolver = new iTextSharp.tool.xml.css.StyleAttrCSSResolver();
                        iTextSharp.tool.xml.pipeline.html.HtmlPipelineContext htmlContext = new iTextSharp.tool.xml.pipeline.html.HtmlPipelineContext(cssAppliers);
                        htmlContext.SetTagFactory(iTextSharp.tool.xml.html.Tags.GetHtmlTagProcessorFactory());
                        iTextSharp.tool.xml.pipeline.end.PdfWriterPipeline pdfWriterPipeline = new iTextSharp.tool.xml.pipeline.end.PdfWriterPipeline(lang_doc, writer);
                        iTextSharp.tool.xml.IPipeline pipeline = new iTextSharp.tool.xml.pipeline.css.CssResolverPipeline(cssRevolver, new iTextSharp.tool.xml.pipeline.html.HtmlPipeline(htmlContext, pdfWriterPipeline));
                        iTextSharp.tool.xml.XMLWorker worker = new iTextSharp.tool.xml.XMLWorker(pipeline, true);
                        iTextSharp.tool.xml.parser.XMLParser xmlParser = new iTextSharp.tool.xml.parser.XMLParser(worker);

                        StringReader mysr = new StringReader(myhtml);
                        xmlParser.Parse(mysr);

                        lang_doc.Close();
                    }
                }
                bytes = ms.ToArray();
            }
            System.IO.File.WriteAllBytes(attach_pdf, bytes);
        }

        static void Main(string[] args)
        {
            var timesFontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "TIMES.TTF");
            BaseFont timesBaseFont = BaseFont.CreateFont(timesFontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font TimRom_Font = new Font(timesBaseFont, 12, Font.NORMAL, BaseColor.BLACK);
            Font TimRom_Underline_Font = new Font(timesBaseFont, 12, Font.UNDERLINE, BaseColor.BLACK);
            Font TimRom_Bold_Font = new Font(timesBaseFont, 12, Font.BOLD, BaseColor.BLACK);
            Font TimRom_Bold_16_Font = new Font(timesBaseFont, 16, Font.BOLD, BaseColor.BLACK);

            //***********************************************
            // CONTRACT MOD TEMPLATE LANGUAGE  612x792 pixels
            //***********************************************
            string lang_path = "FINAL_template_language.pdf";
            Rectangle lang_rect = new Rectangle(612, 792);
            Document lang_doc = new Document(lang_rect);
            PdfWriter.GetInstance(lang_doc, new FileStream(lang_path, FileMode.Create));
            lang_doc.Open();

            Paragraph head_paragraph = new Paragraph();
            head_paragraph.Alignment = Element.ALIGN_CENTER;
            head_paragraph.Add(new Chunk("CONTRACT MODIFICATIONS\n", TimRom_Bold_16_Font));
            head_paragraph.Add(new Chunk("TEMPLATE LANGUAGE\n\n\n", TimRom_Bold_16_Font));

            Paragraph myparagraph = new Paragraph();
            Phrase myphrase = new Phrase();
            myphrase.Add(new Chunk("DATE EXTENSION", TimRom_Underline_Font));
            myphrase.Add(new Chunk(":\n\nParagraph Heading: __________\n\n", TimRom_Font));

            myphrase.Add(new Chunk("Basic Extension: ", TimRom_Bold_Font));
            myphrase.Add(new Chunk("In accordance with paragraph ", TimRom_Font));
            Chunk chunk_yellow_1 = new Chunk("___", TimRom_Font);
            chunk_yellow_1.SetBackground(BaseColor.YELLOW);
            myphrase.Add(chunk_yellow_1);
            myphrase.Add(new Chunk(", the performance period of the Contract is extended by an additional ", TimRom_Font));
            Chunk chunk_yellow_2 = new Chunk("___ (____) years/months", TimRom_Font);
            chunk_yellow_2.SetBackground(BaseColor.YELLOW);
            myphrase.Add(chunk_yellow_2);
            myphrase.Add(new Chunk(" from ", TimRom_Font));
            Chunk chunk_yellow_3 = new Chunk("________", TimRom_Font);
            chunk_yellow_3.SetBackground(BaseColor.YELLOW);
            myphrase.Add(chunk_yellow_3);
            myphrase.Add(new Chunk(", to and including ", TimRom_Font));
            Chunk chunk_yellow_4 = new Chunk("________", TimRom_Font);
            chunk_yellow_4.SetBackground(BaseColor.YELLOW);
            myphrase.Add(chunk_yellow_4);
            myphrase.Add(new Chunk(".\n\n"));

            myphrase.Add(new Chunk("Extension with Option to Extend: ", TimRom_Bold_Font));
            myphrase.Add(new Chunk("Language varies depending on type of contract\n", TimRom_Font));
            myphrase.Add(new Chunk("After option to extend language, add Basic Extension language above\n\n", TimRom_Font));

            myphrase.Add(new Chunk("HAR Extension: ", TimRom_Bold_Font));
            myphrase.Add(new Chunk("Pursuant to Hawaii Administrative Rules for ", TimRom_Font));
            Chunk chunk_yellow_5 = new Chunk("Chapter 103F, Section 3-149-301(c)", TimRom_Font);
            chunk_yellow_5.SetBackground(BaseColor.YELLOW);
            myphrase.Add(chunk_yellow_5);
            myphrase.Add(new Chunk(", the performance period of the Contract is extended by an additional ", TimRom_Font));
            myphrase.Add(chunk_yellow_2);
            myphrase.Add(new Chunk(" from ", TimRom_Font));
            myphrase.Add(chunk_yellow_3);
            myphrase.Add(new Chunk(", to and including ", TimRom_Font));
            myphrase.Add(chunk_yellow_4);
            myphrase.Add(new Chunk(", or when a replacement contract has been executed, whichever occurs first.\n\n", TimRom_Font));

            myphrase.Add(new Chunk("SPO Extension: ", TimRom_Bold_Font));
            myphrase.Add(new Chunk("Pursuant to Hawaii Revised Statutes, Section ", TimRom_Font));
            myphrase.Add(chunk_yellow_4);
            myphrase.Add(new Chunk(" and Hawaii Administrative Rules, Chapter ", TimRom_Font));
            myphrase.Add(chunk_yellow_4);
            myphrase.Add(new Chunk(", the performance period of the Contract is extended by an additional ", TimRom_Font));
            myphrase.Add(chunk_yellow_2);
            myphrase.Add(new Chunk(" from ", TimRom_Font));
            myphrase.Add(chunk_yellow_3);
            myphrase.Add(new Chunk(", to and including ", TimRom_Font));
            myphrase.Add(chunk_yellow_4);
            myphrase.Add(new Chunk(".\n\n"));

            myphrase.Add(new Chunk("BUDGET CHANGE", TimRom_Underline_Font));
            myphrase.Add(new Chunk(":\n\nParagraph Heading: __________\n\n", TimRom_Font));

            myparagraph.Add(myphrase);

            Paragraph paragraph_a = new Paragraph();
            paragraph_a.IndentationLeft = 20.0f;
            paragraph_a.Add(new Chunk("a.  The total amount of compensation is ", TimRom_Font));
            Chunk chunk_yellow_6 = new Chunk("increased/decreased", TimRom_Font);
            chunk_yellow_6.SetBackground(BaseColor.YELLOW);
            paragraph_a.Add(chunk_yellow_6);
            paragraph_a.Add(new Chunk(" by ", TimRom_Font));
            paragraph_a.Add(chunk_yellow_4);
            paragraph_a.Add(new Chunk(" (", TimRom_Font));
            Chunk chunk_yellow_7 = new Chunk("$________", TimRom_Font);
            chunk_yellow_7.SetBackground(BaseColor.YELLOW);
            paragraph_a.Add(chunk_yellow_7);
            paragraph_a.Add(new Chunk(") from ", TimRom_Font));
            paragraph_a.Add(chunk_yellow_4);
            paragraph_a.Add(new Chunk(" (", TimRom_Font));
            paragraph_a.Add(chunk_yellow_7);
            paragraph_a.Add(new Chunk(") to ", TimRom_Font));
            paragraph_a.Add(chunk_yellow_4);
            paragraph_a.Add(new Chunk(" (", TimRom_Font));
            paragraph_a.Add(chunk_yellow_7);
            paragraph_a.Add(new Chunk(") of ", TimRom_Font));
            Chunk chunk_yellow_8 = new Chunk("(State/Federal/Special)", TimRom_Font);
            chunk_yellow_8.SetBackground(BaseColor.YELLOW);
            paragraph_a.Add(chunk_yellow_8);
            paragraph_a.Add(new Chunk(" funds for fiscal year ", TimRom_Font));
            paragraph_a.Add(chunk_yellow_4);
            paragraph_a.Add(new Chunk(".\n\n"));

            Paragraph paragraph_b = new Paragraph();
            paragraph_b.IndentationLeft = 20.0f;
            paragraph_b.Add(new Chunk("b.  The Budget, attached to the Contract as Exhibit \"", TimRom_Font));
            paragraph_b.Add(chunk_yellow_1);
            paragraph_b.Add(new Chunk(",\" is hereby deleted and replaced with a revised Budget, attached hereto as Exhibit \"", TimRom_Font));
            paragraph_b.Add(chunk_yellow_1);
            paragraph_b.Add(new Chunk("\" and made a part hereof.\n", TimRom_Font));
            paragraph_b.Add(new Chunk("OPTIONAL: In addition, the PROVIDER/CONTRACTOR shall submit a STATE approved detailed Budget no later than ", TimRom_Font));
            paragraph_b.Add(chunk_yellow_4);
            paragraph_b.Add(new Chunk(", and failure to comply may result in the withholding of payments to the PROVIDER/CONTRACTOR.  Upon submission of a STATE approved detailed Budget, the STATE approved detailed Budget shall become part of Exhibit\"", TimRom_Font));
            paragraph_b.Add(chunk_yellow_1);
            paragraph_b.Add(new Chunk("\" and made a part of this Contract.\n\n", TimRom_Font));

            Paragraph paragraph_c = new Paragraph();
            paragraph_c.IndentationLeft = 20.0f;
            paragraph_c.Add(new Chunk("c.  Source of Funds: ", TimRom_Font));
            Chunk chunk_yellow_9 = new Chunk("____________________", TimRom_Font);
            chunk_yellow_9.SetBackground(BaseColor.YELLOW);
            paragraph_c.Add(chunk_yellow_9);
            paragraph_c.Add(new Chunk("\n\n", TimRom_Font));

            Paragraph paragraph_d = new Paragraph();
            paragraph_d.IndentationLeft = 20.0f;
            paragraph_d.Add(new Chunk("d.  Optional: Budget Justification\n", TimRom_Font));

            iTextSharp.text.List solidlist = new iTextSharp.text.List(iTextSharp.text.List.UNORDERED);
            solidlist.SetListSymbol("\u2022");
            solidlist.IndentationLeft = 15.0f;
            solidlist.Add(new ListItem("   User Q: Does this modification require a budget justification?", TimRom_Font));

            iTextSharp.text.List hollowlist = new iTextSharp.text.List(iTextSharp.text.List.UNORDERED);
            hollowlist.SetListSymbol("\u25CB");
            hollowlist.IndentationLeft = 30.0f;
            hollowlist.Add(new ListItem("   Yes → Text box provided for user", TimRom_Font));
            hollowlist.Add(new ListItem("   No → No text box provided; paragraph d not added\n\n", TimRom_Font));
            solidlist.Add(hollowlist);

            paragraph_d.Add(solidlist);

            Paragraph paragraph_rate_schedule_change = new Paragraph();
            paragraph_rate_schedule_change.Add(new Chunk("RATE SCHEDULE CHANGE", TimRom_Underline_Font));
            paragraph_rate_schedule_change.Add(new Chunk(":\n\nParagraph Heading: __________\n\n", TimRom_Font));
            paragraph_rate_schedule_change.Add(new Chunk("The Rate Schedule, attached to the Contract as Exhibit \"", TimRom_Font));
            paragraph_rate_schedule_change.Add(chunk_yellow_1);
            paragraph_rate_schedule_change.Add(new Chunk(",\" is hereby deleted and replaced with a revised Rate Schedule, attached hereto as Exhibit \"", TimRom_Font));
            paragraph_rate_schedule_change.Add(chunk_yellow_1);
            paragraph_rate_schedule_change.Add(new Chunk("\" and made a part hereof.", TimRom_Font));

            lang_doc.Add(head_paragraph);
            lang_doc.Add(myparagraph);
            lang_doc.Add(paragraph_a);
            lang_doc.Add(paragraph_b);
            lang_doc.Add(paragraph_c);
            lang_doc.Add(paragraph_d);
            lang_doc.Add(paragraph_rate_schedule_change);
            lang_doc.Close();  
        }
    }
}
