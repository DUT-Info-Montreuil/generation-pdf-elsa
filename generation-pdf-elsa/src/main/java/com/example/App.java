package com.example;

import org.apache.poi.xwpf.usermodel.*;
import java.io.*;

public class App {
    public static void main(String[] args) throws Exception {
        // Charger le fichier modèle .docx
        FileInputStream fis = new FileInputStream("src/main/resources/modele.docx");
        XWPFDocument document = new XWPFDocument(fis);

        // Remplacer les placeholders
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                String text = run.getText(0);
                if (text != null && text.contains("${NOM}")) {
                    text = text.replace("${NOM}", "Jean Dupont");
                    run.setText(text, 0);
                }
            }
        }

        // Sauvegarder le fichier modifié
        FileOutputStream fos = new FileOutputStream("output.docx");
        document.write(fos);
        fos.close();
        document.close();

        System.out.println("Document généré avec succès !");
    }
}
