package org.ktpro.controller;

import com.vladsch.flexmark.html.HtmlRenderer;
import com.vladsch.flexmark.parser.Parser;
import com.vladsch.flexmark.util.ast.Node;
import com.vladsch.flexmark.ext.tables.TablesExtension;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

import java.util.List;

@Controller
public class MarkdownController {
    
    @GetMapping("/")
    public String index() {
        return "index";
    }
    
    @PostMapping("/preview")
    @ResponseBody
    public String preview(@RequestParam String markdown) {
        Parser parser = Parser.builder()
            .extensions(List.of(TablesExtension.create()))
            .build();
        Node document = parser.parse(markdown);
        HtmlRenderer renderer = HtmlRenderer.builder()
            .extensions(List.of(TablesExtension.create()))
            .escapeHtml(true)
            .build();
        return renderer.render(document);
    }
}