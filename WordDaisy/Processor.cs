using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordProcessingTest {
    using DocumentFormat.OpenXml.Vml;
    using DocumentFormat.OpenXml.Vml.Office;
    using System.IO;
    using System.Security.AccessControl;
    using Blip = DocumentFormat.OpenXml.Drawing.Blip;


    class Processor {

        string docTitle, docAuthor;
        string sourceFile, dest;
        string filename;
        int curPageNum;

        int maxCount = 6;

        WordprocessingDocument document;
        ConvertEquations.Program converter;
        ConvertEquations.ConvertEquation ce = new ConvertEquations.ConvertEquation();

        int curLevel = 1;
        Boolean isFirstLevel = true;

        public Processor(string sourceFile, string dest, string filename) {
            this.sourceFile = sourceFile;
            this.dest = dest;
            this.filename = filename.Split('.').First();
            converter = new ConvertEquations.Program();

            curPageNum = 0;
        }


        /*
         * Process a DOCX document, extracting data and outputting a DTBOOK XML.
         * 
         * 
         */
        public Boolean ProcessDocument() {
            // open document - must be existing and not empty
            bool opened = false;
            for (int i = 0; i < maxCount; i++) {
                try {
                    document = WordprocessingDocument.Open(sourceFile, false);
                    opened = true;
                    break;
                } catch (IOException e) {
                    System.Threading.Thread.Sleep(1000);
                }
            }

            if (!opened) {
                return false;
            }

            Body body = document.MainDocumentPart.Document.Body;    // grab document Body
            StylesPart stylesPart = document.MainDocumentPart.StyleDefinitionsPart; // hold document Styles Part

            docTitle = document.PackageProperties.Title;
            docAuthor = document.PackageProperties.Creator;
            dest += @"\" + filename;

            System.IO.Directory.CreateDirectory(dest);
            System.IO.Directory.CreateDirectory(dest + @"\math");

            StringBuilder xml = new StringBuilder();    // will hold XML content as we parse document

            // XML metadata
            xml.Append("<?xml version='1.0' encoding='UTF-8'?><?xml-stylesheet href=\"dtbookbasic.css\""
                + " type=\"text/css\"?><dtbook xmlns=\"http://www.daisy.org/z3986/2005/dtbook/\" version=\"2005-3\""
                + " xml:lang=\"en-US\" xmlns:mml=\"http://www.w3.org/1998/Math/MathML\"><head></head><book showin=\"blp\">");
            xml.Append("<frontmatter><doctitle>" + docTitle + "</doctitle><docauthor>" + docAuthor + "</docauthor></frontmatter><bodymatter><level1>");

            // iterate through document elements and process
            xml.Append(processElements(body.Elements()));

            // close <level#> tags if still necessary
            if (curLevel > 2) {
                xml.Append("</level3>");
            }

            if (curLevel > 1) {
                xml.Append("</level2>");
            }

            // close out metadata
            xml.Append("</level1></bodymatter></book></dtbook>");

            // cleanup
            string xmlOut = xml.ToString().Replace(" & ", " &amp; ").Replace(" &<", " &amp;<").Replace(">& ", ">&amp; ");
            System.IO.File.WriteAllText(dest + @"\" + filename + ".xml", xmlOut);

            return true;
        }


        /*
         * PROCESS ELEMENTS
         * 
         * 
         */

        private string processElements(IEnumerable<OpenXmlElement> elements) {
            StringBuilder xml = new StringBuilder();
            

            foreach (OpenXmlElement element in elements) {
                if (element.GetType() == typeof(Paragraph)) {
                    Paragraph p = (Paragraph)element;

                    // check for Page Number
                    var pageNums = from style in p.Descendants<RunStyle>()
                                   where style.Val.Value == "PageNumberDAISY"
                                   select style;
                    if (pageNums.Count() > 0) {
                        string pgNum = "";
                        foreach (Run r in p.Elements<Run>()) {
                            if (r.Elements<Text>().Count() > 0) {
                                pgNum += r.Elements<Text>().First().Text;
                            }
                        }
                        xml.Append("<pagenum page=\"special\" id=\"" + curPageNum + "\">" + pgNum + "</pagenum>");
                        curPageNum++;
                        continue;
                    }

                    // in case the header isn't found for some reason
                    Boolean headFound = false;
                    if (hasParaStyle(p)) {
                        // look for a header
                        for (int i = 1; i <= 3; i++) {
                            if (p.ParagraphProperties.ParagraphStyleId.Val == "Heading" + i) {
                                if (!isFirstLevel) {
                                    if (curLevel == i) {
                                        xml.Append("</level" + i + ">");
                                    } else if ((curLevel - i) == 1) {
                                        xml.Append("</level" + curLevel + ">");
                                        xml.Append("</level" + i + ">");
                                    } else if ((curLevel - i) > 1) {
                                        xml.Append("</level" + curLevel + ">");
                                        xml.Append("</level" + (curLevel - 1) + ">");
                                        xml.Append("</level" + (curLevel - 2) + ">");
                                    }

                                    curLevel = i;
                                    xml.Append("<level" + curLevel + "><h" + curLevel + ">");
                                } else {
                                    xml.Append("<h" + i + ">");
                                    curLevel = i;
                                    isFirstLevel = false;
                                }

                                headFound = true;
                                break;
                            }
                        }
                    }
   
                    if (!headFound) {
                        xml.Append("<p>");
                    }

                    /*
                     * Iterate through Runs in a paragraph.
                     * Runs hold Text, Images, etc. and can be individually styled
                     * 
                     */
                    foreach (Run run in p.Elements<Run>()) {
                        xml.Append(processRun(run, false));
                    }


                    // end of Paragraph loop
                    if (headFound) {
                        xml.Append("</h" + curLevel + ">");
                    } else {
                        xml.Append("</p>");
                    }
                } else if (element.GetType() == typeof(Table)) {
                    // table
                    xml.Append("<table>");
                    Table table = (Table)element;

                    xml.Append(processTable(table));

                    // close table
                    xml.Append("</table>");
                }
            }



            return xml.ToString();
        }


        /*
         * PROCESS RUN
         * 
         * 
         */

        private string processRun(Run run, Boolean ignoreStyles) {
            StringBuilder xml = new StringBuilder();
            StylesPart stylesPart = document.MainDocumentPart.StyleDefinitionsPart;

            Boolean bold = false, italic = false;

            if (hasRunProperty(run)) {
                var runProperty = run.RunProperties;

                // check the RunProperty for bold/italic
                if (runProperty.Bold != null) { bold = true; }
                if (runProperty.Italic != null) { italic = true; }

                // if not set in RunProperty, lookup the RunStyle and check for bold/italic
                if (!(bold && italic)) {
                    // get ID of Run Style
                    var rStyleId = GetRunStyleId(run);

                    if (rStyleId != null) {
                        // check StylesPart for Run Style ID
                        var rStyle = FindStyleById(stylesPart, rStyleId);

                        if (rStyle != null) {
                            // look for bold elements
                            if (rStyle.Descendants<Bold>().Count() > 0) {
                                bold = true;
                            }

                            // look for italic elements
                            if (rStyle.Descendants<Italic>().Count() > 0) {
                                italic = true;
                            }
                        }
                    }
                }

            }

            // check each child element of the Run
            foreach (OpenXmlElement e in run.Elements()) {
                Type t = e.GetType();

                if (t.Equals(typeof(Text))) {
                    // raw text: apply styles and add to XML
                    if (bold) { xml.Append("<strong>"); }
                    if (italic) { xml.Append("<em>"); }

                    xml.Append(((Text)e).Text);

                    if (italic) { xml.Append("</em>"); }
                    if (bold) { xml.Append("</strong>"); }
                } else if (t.Equals(typeof(Drawing))) {
                    /*
                        * Image
                        * 
                        * Get image data (alt text & URI) and add to XML.
                        * Inline holds DocProperties with alt text attribute.
                        * Inline holds a Blip with rID - use to lookup URI
                        */
                    if (e.Descendants<Inline>().Count() > 0) {
                        Inline inline = e.Descendants<Inline>().First();

                        var descr = inline.DocProperties.Description;
                        string altText = descr == null ? "" : descr.Value.Replace("\"", "'");

                        Blip blip = inline.Descendants<Blip>().First();
                        ImagePart imagePart = (ImagePart)document.MainDocumentPart.GetPartById(blip.Embed.Value);

                        var uri = imagePart.Uri;
                        var filename = uri.ToString().Split('/').Last();
                        var stream = document.Package.GetPart(uri).GetStream();
                        var streamDst = File.OpenWrite(dest + @"\" + filename);

                        CopyStream(stream, streamDst);
                        streamDst.Close();
                        xml.Append("<imggroup><img src=\"" + filename + "\" alt=\"" + altText + "\"></img></imggroup>");
                        //xml.Append("<html xmlns=\"http://www.w3.org/1999/xhtml\"><img src=\"" + filename + "\" alt=\"" + descr + "\"></img></html>");
                    }
                } else if (t.Equals(typeof(EmbeddedObject))) {
                    // OLEObject (MathType equation)
                    var eo = (EmbeddedObject)e;
                    var id = eo.Descendants<ImageData>().First();
                    var eoPart = document.MainDocumentPart.GetPartById(id.RelationshipId);

                    var uri = eoPart.Uri;
                    var filename = uri.ToString().Split('/').Last();
                    var stream = document.Package.GetPart(uri).GetStream();
                    var streamDst = File.OpenWrite(dest + @"\math\" + filename);

                    CopyStream(stream, streamDst);
                    streamDst.Close();

                    var outName = filename.Split('.').First();
                    converter.FileToLocal(dest + @"\math\" + filename, dest + @"\math\" + outName + ".txt", ce);

                    var mml = dest + @"\math\" + outName + ".txt";
                    if (File.Exists(mml)) {
                        var mathml = File.ReadAllText(mml);
                        xml.Append(mathml.Replace("m:", "mml:"));
                    } else {
                        xml.Append("<mml:math><mml:mtext>PROCESSING ERROR: " + outName + "</mml:mtext></mml:math>");
                    }
                }
            }

            return xml.ToString();
        }


        /*
         * PROCESS TABLE
         * 
         * 
         */

        private string processTable(Table table) {
            StringBuilder xml = new StringBuilder();

            // iterate through Rows
            foreach (TableRow row in table.Elements<TableRow>()) {
                xml.Append("<tr>");

                // iterate through Cells (Columns) - also get run content here
                int colIndex = 0;   // keep track of current cell for counting row spans
                foreach (TableCell cell in row.Elements<TableCell>()) {
                    // check for header
                    string cellTag, cellEndTag;
                    if (checkForHeader(cell)) {
                        cellTag = "<th";
                        cellEndTag = "</th>";
                    } else {
                        cellTag = "<td";
                        cellEndTag = "</td>";
                    }

                    // check for horizontal merge (col span)
                    string spanString = "";
                    int colSpan = checkColSpan(cell);

                    if (colSpan > 0) {
                        spanString += " colspan=\"" + colSpan + "\"";
                    }

                    // check for vertical merge (row span)
                    string rowSpan = checkRowSpan(cell);
                    if (rowSpan != null) {
                        if (rowSpan == "restart") {
                            // get Row Span count
                            int spanCount = getRowSpanCount(row, colIndex);

                            if (spanCount > 0) {
                                spanString += " rowspan=\"" + spanCount + "\"";
                            }
                        }
                    }

                    if (rowSpan != "continue") {
                        xml.Append(cellTag + spanString + ">");
                        xml.Append(processElements(cell.Elements()));
                        xml.Append(cellEndTag);
                    }
                    colIndex++;
                }

                xml.Append("</tr>");
            }

            return xml.ToString();
        }

        /*
         * HELPER FUNCTIONS
         * 
         * 
         */

        private string GetRunStyleId(Run run) {
            return run.Descendants<RunStyle>().Count() > 0 ? run.Descendants<RunStyle>().First().Val : null;
        }

        private Style FindStyleById(StylesPart sp, string id) {
            var s = from style in sp.Styles.Elements<Style>()
                    where style.StyleId == id
                    select style;
            return s.Count() > 0 ? s.First() : null;
        }

        private Boolean hasParaStyle(Paragraph p) {
            return p.Descendants<ParagraphStyleId>().Count() > 0;
        }

        private Boolean hasParaProperty(Paragraph p) {
            return p.Elements<ParagraphProperties>().Count() > 0;
        }

        private Boolean hasRunProperty(Run run) {
            return run.Elements<RunProperties>().Count() > 0;
        }

        private Boolean hasRunStyle(Run run) {
            return run.Descendants<RunStyle>().Count() > 0;
        }

        private int checkColSpan(TableCell cell) {
            if (cell.Descendants<GridSpan>().Count() > 0) {
                return cell.Descendants<GridSpan>().First().Val;
            }
            return 0;
        }
        

        private string checkRowSpan(TableCell cell) {
            var span = cell.Descendants<VerticalMerge>().Count() > 0
                     ? cell.Descendants<VerticalMerge>().First() : null;

            if (span != null) {
                if (span.Val != null) {
                    return span.Val;
                }
                return "continue";
            }

            return null;
        }

        private int getRowSpanCount(TableRow row, int cellIndex) {
            int spanCount = 1;

            foreach (TableRow nextRow in row.ElementsAfter()) {
                string span = checkRowSpan(nextRow.Elements<TableCell>().ElementAt(cellIndex));
                if (span == "continue") {
                    spanCount++;
                } else if (span == "restart" || span == null) {
                    break;
                }
            }
            return spanCount;
        }

        private Boolean checkForHeader(TableCell cell) {
            if (cell.Descendants<ParagraphProperties>().Count() > 0
                && cell.Descendants<ParagraphStyleId>().Count() > 0) {
                return cell.Descendants<ParagraphStyleId>().First().Val == "DivDAISY";
            }

            return false;
        }

        private void CopyStream(Stream source, Stream destination) {
            byte[] buffer = new byte[0x8000];
            int read;
            while ((read = source.Read(buffer, 0, buffer.Length)) > 0) {
                destination.Write(buffer, 0, read);
            }
        }
    }
}
