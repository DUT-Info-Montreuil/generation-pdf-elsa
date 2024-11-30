package com.example;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.pdfbox.pdmodel.*;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

public class App {
    public static void main(String[] args) throws Exception {
        // Étape 1 : Remplacement des valeurs dynamiques dans le fichier Word
        String inputDocx = "modeleConventionExemple-2.docx";
        String outputDocx = "output_filled.docx";
        replacePlaceholdersInDocx(inputDocx, outputDocx);

        // Étape 2 : Conversion du fichier Word modifié en PDF
        String outputPdf = "output_filled.pdf";
        convertDocxToPdf(outputDocx, outputPdf);

        System.out.println("Fichier PDF généré avec succès.");
    }

    public static void replacePlaceholdersInDocx(String inputFileName, String outputPath) throws Exception {
        // Chargement du fichier Word
        InputStream fis = App.class.getClassLoader().getResourceAsStream(inputFileName);
        if (fis == null) {
            throw new FileNotFoundException("Fichier introuvable : " + inputFileName);
        }

        XWPFDocument document = new XWPFDocument(fis);

        // Liste des clés à remplacer et leurs valeurs correspondantes
        Map<String, String> replacements = new HashMap<>();
        replacements.put("${annee}", "2023 - 2024");
        replacements.put("${stagiaire}", "Lucas Martin");
        replacements.put("${enseignant référent}", "Dr. Émilie Dupont");
        replacements.put("${tuteur de stage}", "Sophie Durand");
        replacements.put("${représentant légal}", "John Doe");
        replacements.put("${étudiant}", "Lucas Martin");
        replacements.put("${NOM_ORGANISME}", "ALTEN");
        replacements.put("${ADR_ORGANISME}", "123 AI Street, San Francisco");
        replacements.put("${NOM_REPRESENTANT_ORG}", "John Doe");
        replacements.put("${QUAL_REPRESENTANT_ORG}", "Directeur");
        replacements.put("${TEL_ORGANISME}", "01 23 45 67 89");
        replacements.put("${MEL_ORGANISME}", "contact@openai.com");
        replacements.put("${LIEU_DU_STAGE}", "San Francisco HQ");
        replacements.put("${NOM_DU_SERVICE}", "Développement Logiciel");
        replacements.put("${NOM_ETUDIANT1}", "Martin");
        replacements.put("${PRENOM_ETUDIANT}", "Lucas");
        replacements.put("${SEXE_ETUDIANT}", "M");
        replacements.put("${DATE_NAIS_ETUDIANT}", "01/01/2000");
        replacements.put("${ADR_ETUDIANT}", "45 Rue des Lilas, Lyon");
        replacements.put("${TEL_ETUDIANT}", "06 78 90 12 34");
        replacements.put("${MEL_ETUDIANT}", "martin.lucas@example.com");
        replacements.put("${SUJET_DU_STAGE}", "Développement d'une application mobile");
        replacements.put("${DATE_DÉBUT_STAGE}", "01/06/2024");
        replacements.put("${DATE_FIN_STAGE}", "31/08/2024");
        replacements.put("${STA_DUREE}", "3 mois");
        replacements.put("${_STA_JOURS_TOT}", "66");
        replacements.put("${_STA_HEURES_TOT}", "924");
        replacements.put("${STA_REMU_HOR}", "600€/mois");
        replacements.put("${TUT_IUT}", "Dr. Émilie Dupont");
        replacements.put("${TUT_IUT_MEL}", "emilie.dupont@example.com");
        replacements.put("${PRENOM_ENCADRANT}", "Sophie");
        replacements.put("${NOM_ENCADRANT}", "Durand");
        replacements.put("${FONCTION_ENCADRANT}", "Manager");
        replacements.put("${TEL_ENCADRANT}", "07 89 45 12 36");
        replacements.put("${MEL_ENCADRANT}", "sophie.durand@example.com");
        replacements.put("${NOM_CPAM}", "CPAM Paris");
        replacements.put("${Stage_professionnel}", "BUT2");

        // Remplacement dans les paragraphes simples
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            replacePlaceholdersInParagraph(paragraph, replacements);
        }

        // Remplacement dans les cellules de tableau
        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        replacePlaceholdersInParagraph(paragraph, replacements);
                    }
                }
            }
        }

        // Sauvegarde du fichier modifié
        FileOutputStream fos = new FileOutputStream(outputPath);
        document.write(fos);
        fos.close();
        document.close();

        System.out.println("Fichier Word modifié enregistré avec succès.");
    }

    private static void replacePlaceholdersInParagraph(XWPFParagraph paragraph, Map<String, String> replacements) {
        StringBuilder paragraphText = new StringBuilder();

        // Combiner le texte des différents segments du paragraphe
        for (XWPFRun run : paragraph.getRuns()) {
            String text = run.getText(0);
            if (text != null) {
                paragraphText.append(text);
            }
            run.setText("", 0); // Nettoyage du texte existant
        }

        // Remplacer les placeholders par les valeurs associées
        String combinedText = paragraphText.toString();
        for (Map.Entry<String, String> entry : replacements.entrySet()) {
            combinedText = combinedText.replace(entry.getKey(), entry.getValue());
        }

        // Réécriture du texte mis à jour dans le paragraphe
        if (!paragraph.getRuns().isEmpty()) {
            paragraph.getRuns().get(0).setText(combinedText, 0);
        } else {
            paragraph.createRun().setText(combinedText);
        }
    }

    public static void convertDocxToPdf(String docxPath, String pdfPath) throws IOException {
        // Chargement du fichier Word
        FileInputStream fis = new FileInputStream(docxPath);
        XWPFDocument document = new XWPFDocument(fis);

        // Initialisation du document PDF
        PDDocument pdfDocument = new PDDocument();
        PDPage page = new PDPage();
        pdfDocument.addPage(page);

        try (PDPageContentStream contentStream = new PDPageContentStream(pdfDocument, page)) {
            contentStream.setFont(PDType1Font.HELVETICA, 12);
            contentStream.beginText();
            contentStream.setLeading(14.5f);
            contentStream.newLineAtOffset(50, 750);

            float yPosition = 750;
            float margin = 50;
            float lineHeight = 14.5f;

            // Extraction et écriture des paragraphes
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                String[] lines = paragraph.getText().split("\n");
                for (String line : lines) {
                    if (yPosition <= margin) {
                        break;
                    }
                    contentStream.showText(line);
                    contentStream.newLine();
                    yPosition -= lineHeight;
                }
            }

            // Extraction et écriture du texte des tableaux
            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            String[] lines = paragraph.getText().split("\n");
                            for (String line : lines) {
                                if (yPosition <= margin) {
                                    break;
                                }
                                contentStream.showText(line);
                                contentStream.newLine();
                                yPosition -= lineHeight;
                            }
                        }
                    }
                }
            }

            contentStream.endText();
        }

        // Sauvegarde du PDF
        pdfDocument.save(pdfPath);
        pdfDocument.close();
        fis.close();

        System.out.println("Fichier PDF enregistré avec succès.");
    }
}
