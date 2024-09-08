package com.example.dochouse.service;

import com.itextpdf.io.image.ImageData;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.*;
import com.itextpdf.layout.properties.UnitValue;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.sl.usermodel.Slide;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.apache.poi.sl.usermodel.TextShape;
import org.apache.poi.sl.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.sl.usermodel.PictureData;
import org.springframework.web.multipart.MultipartFile;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.Cell;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.Color;
import java.awt.image.BufferedImage;
import java.awt.geom.Rectangle2D;
import java.io.*;
import java.util.List;

import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Image;
import com.itextpdf.io.image.ImageDataFactory;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.Cell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Image;
import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;

@Service
public class FileConversionService {

    // Word to PDF
    public byte[] convertWordToPdf(MultipartFile file) throws Exception {
        try (InputStream in = file.getInputStream();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            PdfWriter writer = new PdfWriter(out);
            PdfDocument pdfDoc = new PdfDocument(writer);
            Document document = new Document(pdfDoc);

            String fileName = file.getOriginalFilename();

            if (fileName != null && fileName.endsWith(".docx")) {
                // Process .docx format
                XWPFDocument docx = new XWPFDocument(in);

                // Handle paragraphs and tables in .docx
                for (XWPFParagraph paragraph : docx.getParagraphs()) {
                    // Add text to PDF
                    for (XWPFRun run : paragraph.getRuns()) {
                        document.add(new Paragraph(run.text()));

                        // Render embedded pictures
                        for (XWPFPicture picture : run.getEmbeddedPictures()) {
                            byte[] imageBytes = picture.getPictureData().getData();
                            addImageToPdf(document, imageBytes);
                        }
                    }
                }

                for (XWPFTable table : docx.getTables()) {
                    addTableToPdf(document, table);
                }
            } else if (fileName != null && fileName.endsWith(".doc")) {
                // Process .doc format
                HWPFDocument doc = new HWPFDocument(in);
                Range range = doc.getRange();

                for (int i = 0; i < range.numParagraphs(); i++) {
                    org.apache.poi.hwpf.usermodel.Paragraph paragraph = range.getParagraph(i);
                    document.add(new Paragraph(paragraph.text()));
                }

                // Handle images in .doc files
                List<Picture> pictures = doc.getPicturesTable().getAllPictures();
                for (Picture picture : pictures) {
                    byte[] imageBytes = picture.getContent();
                    addImageToPdf(document, imageBytes);
                }
            }

            document.close();
            return out.toByteArray();
        }
    }

    private void addImageToPdf(Document document, byte[] imageBytes) throws IOException {
        try (ByteArrayInputStream imageIn = new ByteArrayInputStream(imageBytes)) {
            BufferedImage bufferedImage = ImageIO.read(imageIn);
            if (bufferedImage != null) {
                ByteArrayOutputStream imageOut = new ByteArrayOutputStream();
                ImageIO.write(bufferedImage, "png", imageOut);
                byte[] pdfImageBytes = imageOut.toByteArray();
                Image pdfImage = new Image(ImageDataFactory.create(pdfImageBytes));
                document.add(pdfImage);
            }
        }
    }

    private void addTableToPdf(Document document, XWPFTable table) {
        // Get the number of columns from the first row
        int numCols = table.getRow(0).getTableCells().size();

        // Create a PDF table with the number of columns
        Table pdfTable = new Table(UnitValue.createPercentArray(numCols)).useAllAvailableWidth();

        // Iterate through rows
        for (XWPFTableRow row : table.getRows()) {
            // Create a row in the PDF table
            for (XWPFTableCell cell : row.getTableCells()) {
                String cellText = cell.getText();

                // Create a cell with text
                Cell pdfCell = new Cell().add(new Paragraph(cellText));

                // Optional: Set cell width if needed
                pdfCell.setWidth(UnitValue.createPercentValue(100f / numCols));

                pdfTable.addCell(pdfCell);
            }
        }

        // Add the PDF table to the document
        document.add(pdfTable);
    }



    // PowerPoint to PDF
    public byte[] convertPptToPdf(MultipartFile file) throws Exception {
        InputStream in = file.getInputStream();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        PdfWriter writer = new PdfWriter(out);
        PdfDocument pdfDoc = new PdfDocument(writer);
        Document document = new Document(pdfDoc);

        String fileName = file.getOriginalFilename();
        SlideShow<?, ?> slideshow;

        if (fileName.endsWith(".pptx")) {
            slideshow = new XMLSlideShow(in);
        } else if (fileName.endsWith(".ppt")) {
            slideshow = new HSLFSlideShow(in);
        } else {
            throw new IllegalArgumentException("Unsupported file format");
        }

        // Set slide dimensions (based on original slide size)
        Dimension pageSize = slideshow.getPageSize();

        // Loop through each slide and convert it to an image
        for (Slide<?, ?> slide : slideshow.getSlides()) {
            BufferedImage slideImage = new BufferedImage(pageSize.width, pageSize.height, BufferedImage.TYPE_INT_RGB);
            Graphics2D graphics = slideImage.createGraphics();

            // Set background color and render the slide into the image
            graphics.setPaint(Color.WHITE);
            graphics.fill(new Rectangle2D.Float(0, 0, pageSize.width, pageSize.height));

            // Render the slide content into the graphics object
            slide.draw(graphics);

            // Write the BufferedImage to a ByteArrayOutputStream
            ByteArrayOutputStream imageOut = new ByteArrayOutputStream();
            ImageIO.write(slideImage, "png", imageOut);
            byte[] imageBytes = imageOut.toByteArray();

            // Create an iText Image from the BufferedImage
            Image slideImagePdf = new Image(ImageDataFactory.create(imageBytes));

            // Scale the image to fit the PDF page
            slideImagePdf.scaleToFit(pageSize.width, pageSize.height);

            // Add the image to the PDF document
            document.add(slideImagePdf);
        }

        document.close(); // Close the PDF document
        return out.toByteArray();
    }



    // Text to PDF
    public byte[] convertTextToPdf(MultipartFile file) throws IOException {
        String content = new String(file.getBytes());
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        PdfWriter writer = new PdfWriter(out);
        PdfDocument pdfDoc = new PdfDocument(writer);
        Document pdfDocument = new Document(pdfDoc);

        pdfDocument.add(new Paragraph(content));
        pdfDocument.close();
        return out.toByteArray();
    }



    // Image to PDF
    public byte[] convertImageToPdf(MultipartFile file) throws IOException {
        BufferedImage image = ImageIO.read(file.getInputStream());
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        PdfWriter writer = new PdfWriter(out);
        PdfDocument pdfDoc = new PdfDocument(writer);
        Document pdfDocument = new Document(pdfDoc);

        // Convert BufferedImage to byte array
        ByteArrayOutputStream imageOutputStream = new ByteArrayOutputStream();
        ImageIO.write(image, "png", imageOutputStream);
        byte[] imageBytes = imageOutputStream.toByteArray();

        // Use ImageDataFactory to create ImageData from byte array
        ImageData imageData = ImageDataFactory.create(imageBytes);
        Image pdfImage = new Image(imageData);

        // Add the image to the PDF
        pdfDocument.add(pdfImage);
        pdfDocument.close();

        return out.toByteArray();
    }




    // Excel to PDF
    public byte[] convertExcelToPdf(MultipartFile file) throws Exception {
        InputStream in = file.getInputStream();
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        // Create PdfWriter and PdfDocument
        PdfWriter writer = new PdfWriter(out);
        PdfDocument pdfDoc = new PdfDocument(writer);
        pdfDoc.setDefaultPageSize(PageSize.A4.rotate()); // Set page size to landscape if needed
        Document document = new Document(pdfDoc);

        Workbook workbook = new XSSFWorkbook(in);

        // Iterate through each sheet in the Excel workbook
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);

            // Determine the number of columns from the first row
            int numberOfColumns = sheet.getRow(0).getLastCellNum();
            float[] columnWidths = new float[numberOfColumns];
            for (int j = 0; j < numberOfColumns; j++) {
                columnWidths[j] = 1; // Adjust this value to scale column width
            }
            Table pdfTable = new Table(UnitValue.createPercentArray(columnWidths)).useAllAvailableWidth();

            // Add header row
            Row headerRow = sheet.getRow(0);
            for (int j = 0; j < numberOfColumns; j++) {
                org.apache.poi.ss.usermodel.Cell cell = headerRow.getCell(j);
                pdfTable.addHeaderCell(new com.itextpdf.layout.element.Cell().add(new Paragraph(getCellValue(cell))));
            }

            // Iterate through each row in the sheet
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) continue; // Skip empty rows

                // Add cells to the PDF table
                for (int colIndex = 0; colIndex < numberOfColumns; colIndex++) {
                    org.apache.poi.ss.usermodel.Cell cell = row.getCell(colIndex);
                    pdfTable.addCell(new com.itextpdf.layout.element.Cell().add(new Paragraph(getCellValue(cell))));
                }
            }

            // Add the table to the PDF document
            document.add(pdfTable);
        }

        // Close the document and return the byte array
        document.close();
        return out.toByteArray();
    }

    private String getCellValue(org.apache.poi.ss.usermodel.Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula(); // Handle formulas if necessary
            default:
                return "";
        }
    }



    private byte[] imageToPdfImageData(BufferedImage image) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(image, "png", baos);
        return baos.toByteArray();
    }

    // PNG to JPG and other image conversions
    public byte[] convertImageFormat(MultipartFile file, String format) throws IOException {
        BufferedImage image = ImageIO.read(file.getInputStream());
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(image, format, baos);
        return baos.toByteArray();
    }
}
