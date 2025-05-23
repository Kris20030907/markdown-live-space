package org.ktpro.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import com.itextpdf.text.Document;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.XMLWorkerHelper;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import com.vladsch.flexmark.html.HtmlRenderer;
import com.vladsch.flexmark.parser.Parser;
import com.vladsch.flexmark.ext.tables.TablesExtension;
import com.vladsch.flexmark.util.ast.Node;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;

import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;

@Controller
public class ExportController {

    // TODO: PDF 导出暂时有些问题
    @PostMapping("/export/pdf")
    @ResponseBody
    public void exportToPdf(@RequestParam String markdown, @RequestParam(required = false, defaultValue = "github") String theme, HttpServletResponse response) throws Exception {
        response.setContentType("application/pdf");
        response.setHeader("Content-Disposition", "attachment; filename=export.pdf");

        // 将Markdown转换为HTML
        Parser parser = Parser.builder()
            .extensions(List.of(TablesExtension.create()))
            .build();
        Node document = parser.parse(markdown);
        HtmlRenderer renderer = HtmlRenderer.builder()
            .extensions(List.of(TablesExtension.create()))
            .escapeHtml(true)
            .build();
        String html = renderer.render(document);
        
        // 添加CSS样式到HTML，根据当前主题设置样式
        String cssStyles = "<style>";
        
        // 基础样式
        cssStyles += "body { font-family: Arial, sans-serif; margin: 0; padding: 20px; }" +
            "h1 { font-size: 24pt; margin-bottom: 10pt; }" +
            "h2 { font-size: 18pt; margin-bottom: 8pt; }" +
            "h3 { font-size: 14pt; margin-bottom: 6pt; }" +
            "p { margin-bottom: 10pt; }" +
            "ul, ol { margin-bottom: 10pt; padding-left: 20pt; }" +
            "li { margin-bottom: 5pt; }";
            
        // 根据主题设置不同的样式
        if ("dark".equals(theme)) {
            // 暗黑主题样式
            cssStyles += "body { background-color: #1e1e1e !important; color: #ffffff !important; }" +
                "table { border-collapse: collapse; width: 100%; margin-bottom: 10pt; background-color: #2d2d2d !important; }" +
                "table, th, td { border: 1px solid #444 !important; padding: 5pt; color: #ffffff !important; }" +
                "a { color: #58a6ff !important; }" +
                "code { background-color: #2d2d2d !important; color: #e0e0e0 !important; }" +
                "pre { background-color: #2d2d2d !important; }" +
                ".markdown-body { color: #ffffff !important; background-color: #2d2d2d !important; }" +
                ".markdown-body h1, .markdown-body h2, .markdown-body h3, .markdown-body h4, .markdown-body h5, .markdown-body h6 { color: #ffffff !important; }" +
                ".markdown-body h1, .markdown-body h2 { border-bottom-color: #444 !important; }" +
                ".markdown-body blockquote { color: #b0b0b0 !important; border-left-color: #444 !important; }" +
                ".markdown-body code { background-color: #3a3a3a !important; color: #e0e0e0 !important; }" +
                ".markdown-body pre { background-color: #3a3a3a !important; }" +
                ".markdown-body pre code { color: #e0e0e0 !important; }" +
                ".markdown-body a { color: #58a6ff !important; }" +
                ".markdown-body table { border-color: #444 !important; }" +
                ".markdown-body table th, .markdown-body table td { border-color: #444 !important; color: #ffffff !important; }" +
                ".markdown-body table tr { background-color: #2d2d2d !important; border-top-color: #444 !important; }" +
                ".markdown-body table tr:nth-child(2n) { background-color: #3a3a3a !important; }";
        } else {
            // GitHub主题样式（默认）
            cssStyles += "body { background-color: #ffffff !important; color: #24292e !important; }" +
                "table { border-collapse: collapse; width: 100%; margin-bottom: 10pt; }" +
                "table, th, td { border: 1px solid #dfe2e5 !important; padding: 5pt; }" +
                "a { color: #0366d6 !important; }" +
                "code { background-color: rgba(27, 31, 35, 0.05) !important; color: #24292e !important; }" +
                "pre { background-color: #f6f8fa !important; }" +
                ".markdown-body { font-family: -apple-system, BlinkMacSystemFont, \"Segoe UI\", Helvetica, Arial, sans-serif !important; font-size: 16px !important; line-height: 1.5 !important; word-wrap: break-word !important; color: #24292e !important; background-color: #fff !important; }" +
                ".markdown-body h1, .markdown-body h2 { border-bottom: 1px solid #eaecef !important; }" +
                ".markdown-body blockquote { padding: 0 1em !important; color: #6a737d !important; border-left: 0.25em solid #dfe2e5 !important; }" +
                ".markdown-body code { padding: 0.2em 0.4em !important; background-color: rgba(27, 31, 35, 0.05) !important; border-radius: 3px !important; }" +
                ".markdown-body pre { background-color: #f6f8fa !important; }" +
                ".markdown-body table { border-spacing: 0 !important; border-collapse: collapse !important; }" +
                ".markdown-body table th, .markdown-body table td { padding: 6px 13px !important; border: 1px solid #dfe2e5 !important; }" +
                ".markdown-body table tr { background-color: #fff !important; border-top: 1px solid #c6cbd1 !important; }" +
                ".markdown-body table tr:nth-child(2n) { background-color: #f6f8fa !important; }";
        }
        
        cssStyles += "</style>";
        
        // 只保留Markdown内容的HTML，不添加额外的HTML结构
        // 直接使用最简化的HTML内容，只包含必要的内容
        String purifiedHtml = "<div>" + html + "</div>";
        
        Document pdfDocument = new Document();
        PdfWriter writer = null;
        OutputStream out = null;
        try {
            out = response.getOutputStream();
            writer = PdfWriter.getInstance(pdfDocument, out);
            pdfDocument.open();
            
            try {
                // 将CSS样式直接应用到内容元素上
                ByteArrayInputStream htmlStream = new ByteArrayInputStream(purifiedHtml.getBytes(StandardCharsets.UTF_8));
                XMLWorkerHelper xmlWorkerHelper = XMLWorkerHelper.getInstance();
                
                // 创建CSS样式输入流 - 简化CSS以避免冲突
                String simplifiedCss = cssStyles
                    .replace("<style>", "")
                    .replace("</style>", "")
                    .replaceAll("!important", ""); // 移除!important标记，避免样式冲突
                
                ByteArrayInputStream cssStream = new ByteArrayInputStream(simplifiedCss.getBytes(StandardCharsets.UTF_8));
                
                // 使用XMLWorkerHelper直接解析HTML并应用CSS
                xmlWorkerHelper.parseXHtml(writer, pdfDocument, htmlStream, cssStream, StandardCharsets.UTF_8);
                
                // 检查是否成功添加内容
                if (pdfDocument.getPageNumber() == 0) {
                    fallbackPdfContent(pdfDocument, markdown);
                }
            } catch (Exception e) {
                // XMLWorker解析失败时的备选方案
                e.printStackTrace();
                fallbackPdfContent(pdfDocument, markdown);
            }
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        } finally {
            // 确保资源正确关闭
            try {
                if (pdfDocument != null && pdfDocument.isOpen()) {
                    pdfDocument.close();
                }
                if (writer != null) {
                    writer.close();
                }
                if (out != null) {
                    out.flush();
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
    
    /**
     * PDF导出失败时的备选内容生成方法
     * @param pdfDocument PDF文档对象
     * @param markdown 原始Markdown内容
     */
    private void fallbackPdfContent(Document pdfDocument, String markdown) throws Exception {
        // 不创建新页面，避免多余页面
        // 处理标题
        java.util.regex.Pattern headingPattern = java.util.regex.Pattern.compile("^(#{1,6})\\s+(.+)$", java.util.regex.Pattern.MULTILINE);
        java.util.regex.Matcher headingMatcher = headingPattern.matcher(markdown);
        
        StringBuffer processedContent = new StringBuffer();
        while (headingMatcher.find()) {
            String level = headingMatcher.group(1);
            String title = headingMatcher.group(2).trim();
            // 根据标题级别添加不同的前缀
            String prefix = "";
            if (level.length() == 1) prefix = "> ";
            else if (level.length() <= 3) prefix = "• ";
            String replacement = "\n" + prefix + title + "\n";
            headingMatcher.appendReplacement(processedContent, replacement);
        }
        headingMatcher.appendTail(processedContent);
        
        // 简化Markdown语法
        String plainText = processedContent.toString()
            .replaceAll("\\*\\*(.*?)\\*\\*", "$1") // 移除粗体
            .replaceAll("\\*(.*?)\\*", "$1")       // 移除斜体
            .replaceAll("`(.*?)`", "$1")           // 移除代码
            .replaceAll("\\[([^\\]]+)\\]\\([^\\)]+\\)", "$1") // 简化链接
            .replaceAll("(?m)^\\s*[\\*\\-\\+]\\s", "• ") // 简化无序列表
            .replaceAll("(?m)^\\s*\\d+\\.\\s", "• ");    // 简化有序列表
        
        // 按段落分割并添加到PDF
        String[] paragraphs = plainText.split("\n\n+");
        for (String para : paragraphs) {
            if (!para.trim().isEmpty()) {
                pdfDocument.add(new Paragraph(para.trim()));
            }
        }
    }

    @PostMapping("/export/word")
    @ResponseBody
    public void exportToWord(@RequestParam String markdown, HttpServletResponse response) throws Exception {
        response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        response.setHeader("Content-Disposition", "attachment; filename=export.docx");

        // 将Markdown转换为HTML
        Parser parser = Parser.builder()
            .extensions(List.of(TablesExtension.create()))
            .build();
        Node document = parser.parse(markdown);
        HtmlRenderer renderer = HtmlRenderer.builder()
            .extensions(List.of(TablesExtension.create()))
            .escapeHtml(true)
            .build();
        String html = renderer.render(document);
        
        try (OutputStream out = response.getOutputStream()) {
            XWPFDocument wordDocument = new XWPFDocument();
            
            // 处理HTML内容
            processHtmlForWord(html, wordDocument);
            
            wordDocument.write(out);
            wordDocument.close();
        }
    }
    
    /**
     * 处理HTML内容并转换为Word文档格式
     * @param html HTML内容
     * @param wordDocument Word文档对象
     */
    private void processHtmlForWord(String html, XWPFDocument wordDocument) {
        try {
            // 先处理表格
            processTablesForWord(html, wordDocument);
            
            // 移除表格内容以处理其他元素
            String htmlWithoutTables = html.replaceAll("<table[^>]*>.*?</table>", "");
            
            // 处理标题和段落
            processHeadingsAndParagraphs(htmlWithoutTables, wordDocument);
            
            // 处理列表
            processLists(htmlWithoutTables, wordDocument);
            
            // 添加文档结尾
            XWPFParagraph endParagraph = wordDocument.createParagraph();
            XWPFRun endRun = endParagraph.createRun();
            endRun.addBreak();
        } catch (Exception e) {
            // 添加错误信息
            XWPFParagraph errorParagraph = wordDocument.createParagraph();
            XWPFRun errorRun = errorParagraph.createRun();
            errorRun.setText("文档处理过程中出现错误，部分内容可能无法正确显示。");
            errorRun.setColor("FF0000");
            e.printStackTrace();
        }
    }
    
    /**
     * 处理HTML中的标题和段落
     * @param html HTML内容
     * @param wordDocument Word文档对象
     */
    private void processHeadingsAndParagraphs(String html, XWPFDocument wordDocument) {
        // 提取标题和段落
        java.util.regex.Pattern headingPattern = java.util.regex.Pattern.compile("<h([1-6])[^>]*>(.*?)</h\\1>", java.util.regex.Pattern.DOTALL);
        java.util.regex.Matcher headingMatcher = headingPattern.matcher(html);
        
        // 处理所有标题
        while (headingMatcher.find()) {
            String level = headingMatcher.group(1);
            String content = headingMatcher.group(2).replaceAll("<[^>]*>", "").trim();
            
            if (!content.isEmpty()) {
                XWPFParagraph paragraph = wordDocument.createParagraph();
                
                // 设置标题样式
                switch (level) {
                    case "1":
                        paragraph.setStyle("Heading1");
                        break;
                    case "2":
                        paragraph.setStyle("Heading2");
                        break;
                    case "3":
                        paragraph.setStyle("Heading3");
                        break;
                    default:
                        paragraph.setStyle("Heading4");
                }
                
                XWPFRun run = paragraph.createRun();
                run.setText(content);
                run.setBold(true);
                run.setFontSize(16 - Integer.parseInt(level));
            }
        }
        
        // 提取段落
        java.util.regex.Pattern paragraphPattern = java.util.regex.Pattern.compile("<p[^>]*>(.*?)</p>", java.util.regex.Pattern.DOTALL);
        java.util.regex.Matcher paragraphMatcher = paragraphPattern.matcher(html);
        
        // 处理所有段落
        while (paragraphMatcher.find()) {
            String paragraphHtml = paragraphMatcher.group(1);
            
            // 跳过空段落
            if (paragraphHtml.trim().isEmpty()) {
                continue;
            }
            
            // 创建段落
            XWPFParagraph paragraph = wordDocument.createParagraph();
            
            // 处理段落内的格式
            processInlineFormatting(paragraphHtml, paragraph);
        }
    }
    
    /**
     * 处理HTML中的列表
     * @param html HTML内容
     * @param wordDocument Word文档对象
     */
    private void processLists(String html, XWPFDocument wordDocument) {
        // 提取有序列表
        processOrderedLists(html, wordDocument);
        
        // 提取无序列表
        processUnorderedLists(html, wordDocument);
    }
    
    /**
     * 处理有序列表
     * @param html HTML内容
     * @param wordDocument Word文档对象
     */
    private void processOrderedLists(String html, XWPFDocument wordDocument) {
        // 提取有序列表
        java.util.regex.Pattern olPattern = java.util.regex.Pattern.compile("<ol[^>]*>(.*?)</ol>", java.util.regex.Pattern.DOTALL);
        java.util.regex.Matcher olMatcher = olPattern.matcher(html);
        
        while (olMatcher.find()) {
            String listHtml = olMatcher.group(1);
            
            // 提取列表项
            java.util.regex.Pattern liPattern = java.util.regex.Pattern.compile("<li[^>]*>(.*?)</li>", java.util.regex.Pattern.DOTALL);
            java.util.regex.Matcher liMatcher = liPattern.matcher(listHtml);
            
            int counter = 1;
            while (liMatcher.find()) {
                String itemContent = liMatcher.group(1).trim();
                
                // 创建列表项段落
                XWPFParagraph paragraph = wordDocument.createParagraph();
                paragraph.setIndentationLeft(720); // 缩进
                paragraph.setNumID(createNumbering(wordDocument, 1)); // 设置编号
                
                // 处理列表项内容
                processInlineFormatting(itemContent, paragraph);
                
                counter++;
            }
        }
    }
    
    /**
     * 处理无序列表
     * @param html HTML内容
     * @param wordDocument Word文档对象
     */
    private void processUnorderedLists(String html, XWPFDocument wordDocument) {
        // 提取无序列表
        java.util.regex.Pattern ulPattern = java.util.regex.Pattern.compile("<ul[^>]*>(.*?)</ul>", java.util.regex.Pattern.DOTALL);
        java.util.regex.Matcher ulMatcher = ulPattern.matcher(html);
        
        while (ulMatcher.find()) {
            String listHtml = ulMatcher.group(1);
            
            // 提取列表项
            java.util.regex.Pattern liPattern = java.util.regex.Pattern.compile("<li[^>]*>(.*?)</li>", java.util.regex.Pattern.DOTALL);
            java.util.regex.Matcher liMatcher = liPattern.matcher(listHtml);
            
            while (liMatcher.find()) {
                String itemContent = liMatcher.group(1).trim();
                
                // 创建列表项段落
                XWPFParagraph paragraph = wordDocument.createParagraph();
                paragraph.setIndentationLeft(720); // 缩进
                
                // 使用项目符号
                XWPFRun run = paragraph.createRun();
                run.setText("• ");
                
                // 处理列表项内容
                processInlineFormatting(itemContent, paragraph);
            }
        }
    }
    
    /**
     * 处理内联格式（粗体、斜体等）
     * @param html HTML内容
     * @param paragraph Word段落对象
     */
    private void processInlineFormatting(String html, XWPFParagraph paragraph) {
        // 移除HTML标签，保留文本
        String plainText = html.replaceAll("<[^>]*>", "").trim();
        
        if (!plainText.isEmpty()) {
            XWPFRun run = paragraph.createRun();
            run.setText(plainText);
            
            // 设置格式
            if (html.contains("<strong>") || html.contains("<b>")) {
                run.setBold(true);
            }
            if (html.contains("<em>") || html.contains("<i>")) {
                run.setItalic(true);
            }
        }
    }
    
    /**
     * 创建Word文档中的编号
     * @param document Word文档对象
     * @param level 编号级别
     * @return 编号ID
     */
    private java.math.BigInteger createNumbering(XWPFDocument document, int level) {
        try {
            // 获取编号实例
            org.apache.poi.xwpf.usermodel.XWPFNumbering numbering = document.createNumbering();
            
            // 创建抽象编号定义
            org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum ctAbstractNum = org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAbstractNum.Factory.newInstance();
            ctAbstractNum.setAbstractNumId(java.math.BigInteger.valueOf(0));
            
            // 设置编号格式
            org.openxmlformats.schemas.wordprocessingml.x2006.main.CTLvl lvl = ctAbstractNum.addNewLvl();
            lvl.setIlvl(java.math.BigInteger.valueOf(0));
            lvl.addNewNumFmt().setVal(org.openxmlformats.schemas.wordprocessingml.x2006.main.STNumberFormat.DECIMAL);
            lvl.addNewLvlText().setVal("%1.");
            lvl.addNewStart().setVal(java.math.BigInteger.valueOf(1));
            
            // 创建XWPFAbstractNum对象并设置CTAbstractNum
            org.apache.poi.xwpf.usermodel.XWPFAbstractNum abstractNum = new org.apache.poi.xwpf.usermodel.XWPFAbstractNum(ctAbstractNum);
            
            // 注册抽象编号 - 使用XWPFAbstractNum对象
            java.math.BigInteger abstractNumId = numbering.addAbstractNum(abstractNum);
            
            // 创建具体编号实例
            return numbering.addNum(abstractNumId);
        } catch (Exception e) {
            e.printStackTrace();
            return java.math.BigInteger.valueOf(1);
        }
    }
    
    /**
     * 处理HTML表格并转换为Word表格
     * @param html HTML内容
     * @param wordDocument Word文档对象
     */
    private void processTablesForWord(String html, XWPFDocument wordDocument) {
        try {
            // 使用正则表达式提取表格
            java.util.regex.Pattern pattern = java.util.regex.Pattern.compile("<table[^>]*>(.*?)</table>", java.util.regex.Pattern.DOTALL);
            java.util.regex.Matcher matcher = pattern.matcher(html);
            
            while (matcher.find()) {
                String tableHtml = matcher.group(1);
                
                // 提取表格行
                java.util.regex.Pattern rowPattern = java.util.regex.Pattern.compile("<tr[^>]*>(.*?)</tr>", java.util.regex.Pattern.DOTALL);
                java.util.regex.Matcher rowMatcher = rowPattern.matcher(tableHtml);
                
                // 计算行数和列数
                int rowCount = 0;
                int maxColCount = 0;
                
                while (rowMatcher.find()) {
                    rowCount++;
                    String rowHtml = rowMatcher.group(1);
                    
                    // 计算列数 - 同时支持th和td标签
                    java.util.regex.Pattern cellPattern = java.util.regex.Pattern.compile("<t[hd][^>]*>(.*?)</t[hd]>", java.util.regex.Pattern.DOTALL);
                    java.util.regex.Matcher cellMatcher = cellPattern.matcher(rowHtml);
                    
                    int colCount = 0;
                    while (cellMatcher.find()) {
                        colCount++;
                    }
                    
                    maxColCount = Math.max(maxColCount, colCount);
                }
                
                // 如果表格有效，创建Word表格
                if (rowCount > 0 && maxColCount > 0) {
                    // 创建表格
                    org.apache.poi.xwpf.usermodel.XWPFTable table = wordDocument.createTable(rowCount, maxColCount);
                    
                    // 设置表格基本样式
                    table.setWidth("100%");
                    
                    // 重新匹配行
                    rowMatcher = rowPattern.matcher(tableHtml);
                    int rowIndex = 0;
                    
                    while (rowMatcher.find() && rowIndex < rowCount) {
                        String rowHtml = rowMatcher.group(1);
                        
                        // 检查是否是表头行
                        boolean isHeaderRow = rowHtml.contains("<th");
                        
                        // 提取单元格
                        java.util.regex.Pattern cellPattern = java.util.regex.Pattern.compile("<t[hd][^>]*>(.*?)</t[hd]>", java.util.regex.Pattern.DOTALL);
                        java.util.regex.Matcher cellMatcher = cellPattern.matcher(rowHtml);
                        
                        int colIndex = 0;
                        while (cellMatcher.find() && colIndex < maxColCount) {
                            String cellHtml = cellMatcher.group(1);
                            // 保留基本格式，如粗体和斜体
                            String cellContent = cellHtml.replaceAll("<(?!/?b|/?strong|/?i|/?em)[^>]*>", "").trim();
                            // 替换HTML标签为纯文本
                            cellContent = cellContent.replaceAll("</?b>", "");
                            cellContent = cellContent.replaceAll("</?strong>", "");
                            cellContent = cellContent.replaceAll("</?i>", "");
                            cellContent = cellContent.replaceAll("</?em>", "");
                            
                            // 设置单元格内容
                            org.apache.poi.xwpf.usermodel.XWPFTableCell cell = table.getRow(rowIndex).getCell(colIndex);
                            cell.setText(""); // 清除默认文本
                            
                            // 创建段落和运行
                            XWPFParagraph paragraph = cell.addParagraph();
                            XWPFRun run = paragraph.createRun();
                            run.setText(cellContent);
                            
                            // 设置表头样式
                            if (isHeaderRow) {
                                run.setBold(true);
                            }
                            
                            // 应用格式
                            if (cellHtml.contains("<strong>") || cellHtml.contains("<b>")) {
                                run.setBold(true);
                            }
                            if (cellHtml.contains("<em>") || cellHtml.contains("<i>")) {
                                run.setItalic(true);
                            }
                            
                            colIndex++;
                        }
                        
                        rowIndex++;
                    }
                    
                    // 添加空行
                    wordDocument.createParagraph();
                }
            }
        } catch (Exception e) {
            // 添加错误信息段落
            XWPFParagraph errorParagraph = wordDocument.createParagraph();
            XWPFRun errorRun = errorParagraph.createRun();
            errorRun.setText("表格处理过程中出现错误，部分内容可能无法正确显示。");
            errorRun.setColor("FF0000");
            e.printStackTrace();
        }
    }
}