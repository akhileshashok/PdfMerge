package com.pdf.merge.pdfmerge.service;

import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.atomic.AtomicReference;

@Service
public class PdfMergeServiceImpl implements PdfMergeService {

    @Autowired
    private ExcelService excelService;

    @Override
    public void merge() {

            excelService.convertExcelToPdf();

            File dir = new File("C:/MySpace/Data");
            File[] pdfs = dir.listFiles();

            List<String> firstPage = new ArrayList<>();
            List<String> remainingPages = new ArrayList<>();

            Arrays.stream(pdfs).forEach(pdf -> {
                if (pdf.getName().substring(0, pdf.getName().length() - 4).matches("^[a-zA-Z]*$")) {
                    firstPage.add(pdf.getName());
                } else {
                    remainingPages.add(pdf.getName());
                }
            });

            firstPage.stream().forEach(page -> {
                File upper = new File("C:/MySpace/Data/" + page);
                File lower = findFile(upper.getName().substring(0, upper.getName().length() - 4), remainingPages);

                PDFMergerUtility obj = new PDFMergerUtility();
                obj.setDestinationFileName("C:/MySpace/Output/" + lower.getName());

                try {
                    obj.addSource(upper);
                    obj.addSource(lower);

                    obj.mergeDocuments();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            });
    }

    private File findFile(String upper, List<String> remainingPages) {
        AtomicReference<File> file = new AtomicReference<>();
        remainingPages.stream().forEach(page -> {
            if (page.substring(0, page.length() - 4).contains(upper)) {
                file.set(new File("C:/MySpace/Data/" + page));
            }
        });
        return file.get();
    }
}
