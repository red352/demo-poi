package com.demo.service.demopoi;

import lombok.SneakyThrows;
import org.apache.poi.xwpf.usermodel.XWPFComment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.json.JSONArray;
import org.json.JSONObject;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMarkup;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

import java.io.FileInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

/**
 * @author LuoYunXiao
 * @since 2024/1/19 20:50
 */
public class WordTest {
    private final String file = "C:/Users/18294/Desktop/workSpace/65a0de10d5a102ccc20d46a1.docx";

    @Test
    @SneakyThrows
    void test1() {
        XWPFDocument document = new XWPFDocument(Files.newInputStream(Path.of(file)));

        Map<String, String> hashMap = new HashMap<>(document.getComments().length);
        for (XWPFComment comment : document.getComments()) {
            hashMap.put(comment.getId(), comment.getText());
        }
        StringBuilder stringBuilder = new StringBuilder();
        var swap = 0;
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            System.out.println(paragraph.getDocument().toString());
            for (XWPFRun run : paragraph.getRuns()) {
                swap += run.text().length();
                for (CTMarkup ctMarkup : run.getCTR().getCommentReferenceArray()) {
                }
            }
            stringBuilder.append("\n");
        }
//        System.out.println(stringBuilder);


    }

    @Test
    @SneakyThrows
    void test2() {
        JSONArray jsonArray = processDocument(file);
        System.out.println(jsonArray.toString());
        StringBuilder builder = new StringBuilder();
        for (int i = 0; i < jsonArray.length(); i++) {
            JSONObject jsonObject = jsonArray.getJSONObject(i);
            String text = jsonObject.getString("text");
            builder.append(text);
            builder.append("\n");
        }
        System.out.println(builder);
    }


    public static JSONArray processDocument(String filePath) throws Exception {
        FileInputStream fis = new FileInputStream(filePath);
        XWPFDocument document = new XWPFDocument(fis);
        JSONArray paragraphsArray = new JSONArray();

        // 构建批注ID到批注内容的映射
        Map<String, XWPFComment> commentMap = new HashMap<>();
        for (XWPFComment comment : document.getComments()) {
            commentMap.put(comment.getId(), comment);
        }

        // 遍历文档的每个段落
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            JSONObject paragraphObject = new JSONObject();
            paragraphObject.put("text", paragraph.getText());

            JSONArray commentsArray = new JSONArray();

            // 遍历段落中的所有运行
            for (XWPFRun run : paragraph.getRuns()) {
                CTR ctr = run.getCTR();
                // 检查是否有批注引用
                if (ctr.getCommentReferenceArray().length > 0) {
                    CTMarkup commentReference = ctr.getCommentReferenceArray(0);
                    String commentId = String.valueOf(commentReference.getId());
                    if (commentMap.containsKey(commentId)) {
                        XWPFComment comment = commentMap.get(commentId);
                        JSONObject commentObject = new JSONObject();
                        commentObject.put("author", comment.getAuthor());
                        commentObject.put("content", comment.getText());
                        commentsArray.put(commentObject);
                    }
                }
            }

            if (commentsArray.length() > 0) {
                paragraphObject.put("comments", commentsArray);
            }
            paragraphsArray.put(paragraphObject);
        }

        fis.close();
        return paragraphsArray;
    }
}
