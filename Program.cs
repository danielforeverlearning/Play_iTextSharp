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

    public class My_Custom_iTextSharp_ElementHandler : iTextSharp.tool.xml.IElementHandler
    {
        public List<IElement> elements = new List<IElement>();
        public void Add(iTextSharp.tool.xml.IWritable w)
        {
            if (w is iTextSharp.tool.xml.pipeline.WritableElement)
            {
                elements.AddRange(((iTextSharp.tool.xml.pipeline.WritableElement)w).Elements());
            }
        }
    }


    class Program
    {

        static void  ExamplePDFStamper()
        {
            //***************************************
            // SUMMARY_SHEET  612x792 pixels
            // using PDFStamper
            //***************************************
            Dictionary<string, string> Request = new Dictionary<string, string>(); 
            string timesFontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "TIMES.TTF");
            BaseFont TimesRomanFont = BaseFont.CreateFont(timesFontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font TimRom_Font = new Font(TimesRomanFont, 12, Font.NORMAL, BaseColor.BLACK);

            string summary_template = "\\template\\Contract_Mods\\SUMMARY_SHEET.pdf";

            //copy the source template file
            string summary_copy = "\\copy_summary_temp.pdf";
            FileInfo summary_fi = new FileInfo(summary_template);
            summary_fi.CreationTime = DateTime.Now;
            summary_fi.CopyTo(summary_copy);

            //check make sure o.s. copied the file
            FileInfo summary_dest_fi = new FileInfo(summary_copy);
            if (summary_dest_fi.Exists == false)
            {
                throw new Exception("failed to copy file");
            }
            PdfReader summary_rdr = new PdfReader(summary_copy);
            iTextSharp.text.Rectangle summary_rr = summary_rdr.GetPageSizeWithRotation(1); //summary_rr contains width and Height
            string summary_final_pdf = "\\FINAL_summary.pdf";
            FileStream summary_fs = new FileStream(summary_final_pdf, FileMode.Create);
            PdfStamper summary_stamper = new PdfStamper(summary_rdr, summary_fs);
            PdfContentByte summary_cb = summary_stamper.GetOverContent(1);

            summary_cb.BeginText();
            summary_cb.SetFontAndSize(TimesRomanFont, 10);
            summary_cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Request["asolog1"] + "-" + Request["asolog2"], 222.0f, 642.0f, 0);                                  //ASO Log Number
            summary_cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Request["mod_num"], 222.0f, 622.0f, 0);                                                             //MODIFICATION ORDER NO.
            summary_cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Request["contractor_provider_name"], 222.0f, 603.0f, 0);                                            //Contractor/Provider
            summary_cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Request["initial_term_start_date"] + " - " + Request["initial_term_end_date"], 222.0f, 582.0f, 0);  //Initial Term of Contract
            summary_cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Request["extension_terms"], 222.0f, 562.0f, 0);                                                     //Extension Term(s)
            summary_cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Request["funding_source"], 222.0f, 542.0f, 0);                                                      //Source of Funding
            summary_cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Request["procurement_method"], 222.0f, 522.0f, 0);                                                  //Procurement Method
            summary_cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Request["geographic_area"], 222.0f, 502.0f, 0);                                                     //Geographic Area
            summary_cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Request["services_provided"], 222.0f, 482.0f, 0);                                                   //Services Provided
            summary_cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Request["remarks"], 222.0f, 427.0f, 0);                                                             //Remarks
            summary_cb.EndText();
            summary_stamper.Close();
            summary_rdr.Close();
            summary_fs.Close();
        }

        static void  ExampleDirectWriting()
        {
            string timesFontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "TIMES.TTF");
            BaseFont TimesRomanFont = BaseFont.CreateFont(timesFontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font TimRom_Font = new Font(TimesRomanFont, 12, Font.NORMAL, BaseColor.BLACK);

            string dummy_path = "\\dummy.pdf";
            Rectangle dummy_rect = new Rectangle(612, 792);
            Document dummy_doc = new Document(dummy_rect);
            PdfWriter.GetInstance(dummy_doc, new FileStream(dummy_path, FileMode.Create));
            dummy_doc.Open();
            Paragraph dummy_paragraph = new Paragraph();
            dummy_paragraph.Add(new Chunk("MY DUMMY TEST\n", TimRom_Font));
            iTextSharp.text.List first_list = new List(true, true);
            first_list.Lowercase = true;
            first_list.Add(new ListItem("apple", TimRom_Font));
            first_list.Add(new ListItem("should be second list now", TimRom_Font));

            iTextSharp.text.List second_list = new List(true, true);
            second_list.Lowercase = true;
            second_list.Add(new ListItem("car", TimRom_Font));
            second_list.Add(new ListItem("truck", TimRom_Font));
            second_list.Add(new ListItem("airplane", TimRom_Font));

            first_list.Add(second_list);
            first_list.Add(new ListItem("pear", TimRom_Font));
            first_list.Add(new ListItem("strawberry", TimRom_Font));

            dummy_paragraph.Add(first_list);
            dummy_doc.Add(dummy_paragraph);
            dummy_doc.Close();
        }

        static void Example_NicEditWidget_HTML_Repair()
        {
            Dictionary<string, string> Request = new Dictionary<string, string>(); 
            string timesFontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "TIMES.TTF");
            BaseFont TimesRomanFont = BaseFont.CreateFont(timesFontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font TimRom_Font = new Font(TimesRomanFont, 12, Font.NORMAL, BaseColor.BLACK);
            Font TimRom_Underline_Font = new Font(TimesRomanFont, 12, Font.UNDERLINE, BaseColor.BLACK);
            Font TimRom_Bold_Font = new Font(TimesRomanFont, 12, Font.BOLD, BaseColor.BLACK);
            Font TimRom_Bold_16_Font = new Font(TimesRomanFont, 16, Font.BOLD, BaseColor.BLACK);


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
            // (2) So above idea was a disaster, but now i successfully can get to elements list of phrase, and populate empty phrase with elements from html string.
            // (3) so now that we can make our own phrases direct from html string, we can force font-family for <ol> <li> tags......
            //     yes .........all this to control font-family in sub-list under main-list in <ol> <li> tags  (sigh...)
            // (4) manipulate the phrase to try and fix iTextSharp bug of uncontrollable font-family in "sublist" <ol> <li> inside "main" <ol><li>
            //     Holy cow it works !!!!!  ok later got to make this code work for date extension, budget change and rate schedule change,
            //     for now this is just test code for budget change, since we know it has a 3 item sublist a. b. c.
            //     we could have used old exisiting functio html2pdf which uses interop office word, but i saw the incorrect yellow background artifact
            //     that was missing the underline character, so i knew it was unreliable, and it is better to learn iTextSharp,
            //     and it is better to learn how to do HTML sub-ordered lists inside main-ordered lists from NicEdit widgets into PDF using iTextSharp.
            //******************************************************************************************************************************************************************************
            string dummy_path = "\\dummy.pdf";
            Rectangle dummy_rect = new Rectangle(612, 792);
            Document dummy_doc = new Document(dummy_rect);
            PdfWriter.GetInstance(dummy_doc, new FileStream(dummy_path, FileMode.Create));
            dummy_doc.Open();
            Paragraph dummy_paragraph = new Paragraph();
            dummy_paragraph.Add(new Chunk("MY DUMMY TEST\n", TimRom_Font));
            iTextSharp.text.List first_list = new List(true, true);
            first_list.Lowercase = true;
            first_list.Add(new ListItem("apple", TimRom_Font));
            first_list.Add(new ListItem("should be second list now", TimRom_Font));

            iTextSharp.text.List second_list = new List(true, true);
            second_list.Lowercase = true;
            second_list.Add(new ListItem("car", TimRom_Font));
            second_list.Add(new ListItem("truck", TimRom_Font));
            second_list.Add(new ListItem("airplane", TimRom_Font));

            first_list.Add(second_list);
            first_list.Add(new ListItem("pear", TimRom_Font));
            first_list.Add(new ListItem("strawberry", TimRom_Font));

            dummy_paragraph.Add(first_list);
            dummy_doc.Add(dummy_paragraph);
            dummy_doc.Close();

            //***********************************************************

            string attach_pdf =  "\\FINAL_attachments.pdf";
            Rectangle attach_rect = new Rectangle(612, 792);
            Document attach_doc = new Document(dummy_rect);
            PdfWriter.GetInstance(attach_doc, new FileStream(attach_pdf, FileMode.Create));
            attach_doc.Open();

            Paragraph head1_paragraph = new Paragraph();
            head1_paragraph.Alignment = Element.ALIGN_RIGHT;
            head1_paragraph.Add(new Chunk("ATTACHMENT\n", TimRom_Font));
            attach_doc.Add(head1_paragraph);

            Paragraph head2_paragraph = new Paragraph();
            head2_paragraph.Alignment = Element.ALIGN_LEFT;
            Chunk head2_chunk = new Chunk("CONTRACT_MODIFICATIONS", TimRom_Font);
            head2_chunk.SetUnderline(1.0f, -2.0f);
            head2_paragraph.Add(head2_chunk);
            Chunk colon_chunk = new Chunk(":\n", TimRom_Font);
            head2_paragraph.Add(colon_chunk);
            attach_doc.Add(head2_paragraph);

            Paragraph effective_paragraph = new Paragraph();
            effective_paragraph.IndentationLeft = 20.0f;
            effective_paragraph.Add(new Chunk("Effective ", TimRom_Font));
            Chunk effective_date_chunk = new Chunk(Request["effective_date"], TimRom_Font);
            effective_date_chunk.SetUnderline(1.0f, -2.0f);
            effective_paragraph.Add(effective_date_chunk);
            effective_paragraph.Add(new Chunk(", the parties mutually agree that the PROVIDER/CONTRACTOR shall continue to provide the required services with the following modifications:", TimRom_Font));
            attach_doc.Add(effective_paragraph);


            iTextSharp.text.List main_list = new List(true, false, 50.0f);
            string top_html = "<html>" +
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
                                "<body style=\"font-size:12.0pt; font-family:Courier;\">";
            if (Request["conmodtype_date_extension"] == "on")
            {
                string myhtml = top_html + Request["date_extension"] + "</body></html>";
                //ok for some reason iTextSharp HTML to PDF did not like <p> tags without any "stuff inside it" like align and syle
                //So when i do this below it works
                myhtml = myhtml.Replace("<p>", "<p align=\"left\" style=\"font-size:12.0pt;font-family:Courier;\">");
                myhtml = myhtml.Replace("<br>", "<br></br>");
                StringReader mysr = new StringReader(myhtml); //nicEdit html string
                Phrase myphrase = new Phrase();
                var mh = new My_Custom_iTextSharp_ElementHandler();
                iTextSharp.tool.xml.XMLWorkerHelper.GetInstance().ParseXHtml(mh, mysr);
                foreach (var element in mh.elements)
                {
                    myphrase.Add(element);
                }

                //for now date_extension no sublist like budget_change
                ListItem force_font_mainlist_item = new ListItem("", TimRom_Font); //necessary to force main_list font-family style in <ol> <li> tag
                force_font_mainlist_item.Add(myphrase);
                main_list.Add(force_font_mainlist_item);
            }


            if (Request["conmodtype_budget_change"] == "on")
            {
                string myhtml = top_html + Request["budget_change"] + "</body></html>";
                //ok for some reason iTextSharp HTML to PDF did not like <p> tags without any "stuff inside it" like align and syle
                //So when i do this below it works
                myhtml = myhtml.Replace("<p>", "<p align=\"left\" style=\"font-size:12.0pt;font-family:Courier;\">");
                myhtml = myhtml.Replace("<br>", "<br></br>");

                StreamWriter bigdump = new StreamWriter("C:\\Users\\mlee\\Desktop\\bigdump.txt");
                bigdump.WriteLine(myhtml);
                bigdump.Close();

                StringReader mysr = new StringReader(myhtml); //nicEdit html string
                Phrase myphrase = new Phrase();
                var mh = new My_Custom_iTextSharp_ElementHandler();
                iTextSharp.tool.xml.XMLWorkerHelper.GetInstance().ParseXHtml(mh, mysr);
                foreach (var element in mh.elements)
                {
                    myphrase.Add(element);
                }

                //****************************************************************************************************************************************
                // (4) manipulate the phrase to try and fix iTextSharp bug of uncontrollable font-family in "sublist" <ol> <li> inside "main" <ol><li>
                //     Holy cow it works !!!!!  ok later got to make this code work for date extension, budget change and rate schedule change,
                //     for now this is just test code for budget change, since we know it has a 3 item sublist a. b. c.
                //     we could have used old exisiting functio html2pdf which uses interop office word, but i saw the incorrect yellow background artifact
                //     that was missing the underline character, so i knew it was unreliable, and it is better to learn iTextSharp,
                //     and it is better to learn how to do HTML sub-ordered lists inside main-ordered lists from NicEdit widgets into PDF using iTextSharp.
                //*****************************************************************************************************************************************
                Phrase manipulated_phrase = new Phrase();
                manipulated_phrase.Add(myphrase[0]);            //iTextSharp.text.Chunk  Attachment #, Attachment Name

                iTextSharp.text.List sub_list = new List(true, false, 60.0f);

                //((iTextSharp.text.List)myphrase[1]).Items[0]; //this is a sub_list item  (ListItem)
                //So need to get Chunks inside this which is
                //((iTextSharp.text.List)myphrase[1]).Items[0].Chunks.Count
                //((iTextSharp.text.List)myphrase[1]).Items[0].Chunks[ii]
                Phrase sublist_phrase0 = new Phrase();
                int count0 = ((iTextSharp.text.List)myphrase[1]).Items[0].Chunks.Count;
                for (int ii = 0; ii < count0; ii++)
                {
                    sublist_phrase0.Add(((iTextSharp.text.List)myphrase[1]).Items[0].Chunks[ii]);
                }
                ListItem force_font_sublist_item0 = new ListItem("", TimRom_Font); //necessary to force sub_list font-family style in <ol> <li> tag
                force_font_sublist_item0.Add(sublist_phrase0);
                sub_list.Add(force_font_sublist_item0);

                //((iTextSharp.text.List)myphrase[1]).Items[1]; //this is a sub_list item  (ListItem)
                //So need to get Chunks inside this which is
                //((iTextSharp.text.List)myphrase[1]).Items[1].Chunks.Count
                //((iTextSharp.text.List)myphrase[1]).Items[1].Chunks[ii]
                Phrase sublist_phrase1 = new Phrase();
                int count1 = ((iTextSharp.text.List)myphrase[1]).Items[1].Chunks.Count;
                for (int ii = 0; ii < count1; ii++)
                {
                    sublist_phrase1.Add(((iTextSharp.text.List)myphrase[1]).Items[1].Chunks[ii]);
                }
                ListItem force_font_sublist_item1 = new ListItem("", TimRom_Font); //necessary to force sub_list font-family style in <ol> <li> tag
                force_font_sublist_item1.Add(sublist_phrase1);
                sub_list.Add(force_font_sublist_item1);

                //((iTextSharp.text.List)myphrase[1]).Items[2]; //this is a sub_list item  (ListItem)
                //So need to get Chunks inside this which is
                //((iTextSharp.text.List)myphrase[1]).Items[2].Chunks.Count
                //((iTextSharp.text.List)myphrase[1]).Items[2].Chunks[ii]
                Phrase sublist_phrase2 = new Phrase();
                int count2 = ((iTextSharp.text.List)myphrase[1]).Items[2].Chunks.Count;
                for (int ii = 0; ii < count2; ii++)
                {
                    sublist_phrase2.Add(((iTextSharp.text.List)myphrase[1]).Items[2].Chunks[ii]);
                }
                ListItem force_font_sublist_item2 = new ListItem("", TimRom_Font); //necessary to force sub_list font-family style in <ol> <li> tag
                force_font_sublist_item2.Add(sublist_phrase2);
                sub_list.Add(force_font_sublist_item2);
                manipulated_phrase.Add(sub_list);

                ListItem force_font_mainlist_item = new ListItem("", TimRom_Font); //necessary to force main_list font-family style in <ol> <li> tag
                force_font_mainlist_item.Add(manipulated_phrase);
                main_list.Add(force_font_mainlist_item);
            }


            if (Request["conmodtype_rate_schedule_change"] == "on")
            {
                string myhtml = top_html + Request["rate_schedule_change"] + "</body></html>";
                //ok for some reason iTextSharp HTML to PDF did not like <p> tags without any "stuff inside it" like align and syle
                //So when i do this below it works
                myhtml = myhtml.Replace("<p>", "<p align=\"left\" style=\"font-size:12.0pt;font-family:Courier;\">");
                myhtml = myhtml.Replace("<br>", "<br></br>");
                StringReader mysr = new StringReader(myhtml); //nicEdit html string
                Phrase myphrase = new Phrase();
                var mh = new My_Custom_iTextSharp_ElementHandler();
                iTextSharp.tool.xml.XMLWorkerHelper.GetInstance().ParseXHtml(mh, mysr);
                foreach (var element in mh.elements)
                {
                    myphrase.Add(element);
                }

                //for now no sublist in rate_schedule_change like budget change
                ListItem force_font_mainlist_item = new ListItem("", TimRom_Font); //necessary to force main_list font-family style in <ol> <li> tag
                force_font_mainlist_item.Add(myphrase);
                main_list.Add(force_font_mainlist_item);
            }

            attach_doc.Add(main_list);
            attach_doc.Close();
        }//Example_NicEditWidget_HTML_Repair

        static void Example_Merge_PDFs()
        {
            //******** MERGE PDF FILES *************************************************
            //see pdfUtils.cs - StampPageNumber for example of stamping page numbers
            //For now let us save this for a little later, we gotta move on
            //**************************************************************************
            List<string> fileList = new List<string>();
            //fileList.Add(copy_pdf_template);
            //fileList.Add(summary_copy);
            //fileList.Add(conack_copy);
            //fileList.Add(provack_copy);
            //fileList.Add(attach_pdf);

            string mergedFile = String.Empty;
            int totalPages = 0;
            if (fileList.Count > 0)
            {
                mergedFile = "\\FINAL_merged.pdf";

                iTextSharp.text.Document document = new iTextSharp.text.Document();

                using (FileStream copystream = new FileStream(mergedFile, FileMode.Create))
                {
                    try
                    {
                        PdfCopy copy = new PdfCopy(document, copystream);

                        document.Open();
                        PdfReader reader;

                        int docPages;

                        // loop over the documents you want to concatenate
                        for (int curFile = 0; curFile < fileList.Count; curFile++)
                        {
                            if (File.Exists(fileList[curFile]))
                            {
                                reader = new PdfReader(fileList[curFile]);

                                // loop over the pages in that document
                                docPages = reader.NumberOfPages;
                                totalPages += docPages;

                                for (int pageNum = 1; pageNum <= docPages; pageNum++)
                                {
                                    copy.AddPage(copy.GetImportedPage(reader, pageNum));
                                }

                                reader.Close();
                                copy.FreeReader(reader);
                            }

                        }

                        document.Close();
                        copystream.Close();

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
        }//Example_Merge_PDFs


        static void Main(string[] args)
        {

        }
    }//class
}//namespace
