package com.pdf.merge.pdfmerge.service;

import com.spire.xls.*;
import com.spire.xls.core.IXLSRange;
import org.springframework.stereotype.Service;

@Service
public class ExcelServiceImpl implements ExcelService {

    @Override
    public void convertExcelToPdf() {
        try {
            Workbook workbook = new Workbook();
            workbook.loadFromFile("C:/MySpace/Dataset/Cover/Cover.xlsx");
            int sheetCount = workbook.getWorksheets().getCount();

            for (int i = 1; i < sheetCount; i++) {
                workbook.getWorksheets().get(i).setVisibility(WorksheetVisibility.Hidden);
            }

            for (int j = 0; j < sheetCount; j++) {
                Worksheet ws = workbook.getWorksheets().get(j);
                IXLSRange iXLSRange = ws.get("G6");
                String fileName = iXLSRange.getText().substring(0, iXLSRange.getText().indexOf(" "));
                workbook.saveToFile("C:/MySpace/Dataset/Data/" + fileName + ".pdf", FileFormat.PDF);

                if (j < workbook.getWorksheets().getCount() - 1) {
                    workbook.getWorksheets().get(j + 1).setVisibility(WorksheetVisibility.Visible);
                    workbook.getWorksheets().get(j).setVisibility(WorksheetVisibility.Hidden);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
