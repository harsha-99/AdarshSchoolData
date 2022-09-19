package org.example;

import org.apache.commons.math3.util.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

public class Main {

    private static Map<Integer, Pair<Integer, Integer>> config;

    private static void getConfig() {
        config = new HashMap<>();
        //0 indexed
        //CandidateName
        config.put(5, new Pair<>(5,2));
        //FatherName
        config.put(6, new Pair<>(5,8));
        //MotherName
        config.put(7, new Pair<>(8,2));
        //CandidateNameKruti
        config.put(8, new Pair<>(6,2));
        //FatherNameKruti
        config.put(9, new Pair<>(6,8));
        //MotherNameKruti
        config.put(10, new Pair<>(9,2));

        //Gender
        config.put(11, new Pair<>(8,8));
        //Medium
        config.put(13, new Pair<>(11,8));
        //CasteCategory
        config.put(14, new Pair<>(13,2));
        //CandidateType1
        config.put(15, new Pair<>(11,2));
        //IsMinority
        config.put(19, new Pair<>(13,8));

        //Subject01
        config.put(20, new Pair<>(19,2));
        //Subject02
        config.put(21, new Pair<>(20,2));
        //Subject03
        config.put(22, new Pair<>(21,2));
        //Subject04
        config.put(23, new Pair<>(19,4));
        //Subject05
        config.put(24, new Pair<>(20,4));
        //Subject06
        config.put(25, new Pair<>(21,4));

        //DOB - ddmmyyyy
        config.put(12, new Pair<>(15,2));
        //MobileNumber
        config.put(29, new Pair<>(17,2));
        //AadharNumber
        config.put(30, new Pair<>(15,8));
        //EmailID
        config.put(31, new Pair<>(17,8));

    }

    private static void updateDetails(XSSFWorkbook workbook, XSSFSheet worksheet, XSSFSheet resultSheet, int sourceRowNum, int rowId) throws IOException {
        // Get the source / new row
        for (Map.Entry<Integer, Pair<Integer, Integer>> entry : config.entrySet()) {
            int i = entry.getKey();
            int r = entry.getValue().getKey();
            int c = entry.getValue().getValue();
                XSSFCell oldCell = worksheet.getRow(sourceRowNum).getCell(i);
                XSSFCell newCell;
                newCell = resultSheet.getRow(r+rowId*26).getCell(c);

                // Set the cell data value
                switch (oldCell.getCellType()) {
                    case BLANK:
                        newCell.setCellValue(oldCell.getStringCellValue());
                        break;
                    case BOOLEAN:
                        newCell.setCellValue(oldCell.getBooleanCellValue());
                        break;
                    case ERROR:
                        newCell.setCellErrorValue(oldCell.getErrorCellValue());
                        break;
                    case FORMULA:
                        newCell.setCellFormula(oldCell.getCellFormula());
                        break;
                    case NUMERIC:
                        newCell.setCellValue(oldCell.getNumericCellValue());
                        break;
                    case STRING:
                        newCell.setCellValue(oldCell.getRichStringCellValue());
                        break;
                }
            }
        final FileInputStream stream =
                new FileInputStream( "/Users/harsha/IdeaProjects/test/src/main/resources/9th/" + String.format("%03d", sourceRowNum) + ".jpg" );
        final CreationHelper helper = workbook.getCreationHelper();
        final Drawing drawing = resultSheet.createDrawingPatriarch();

        final ClientAnchor anchor = helper.createClientAnchor();
        anchor.setAnchorType( ClientAnchor.AnchorType.MOVE_AND_RESIZE );


        final int pictureIndex =
                workbook.addPicture(IOUtils.toByteArray(stream), Workbook.PICTURE_TYPE_JPEG);

        final FileInputStream stream2 =
                new FileInputStream( "/Users/harsha/IdeaProjects/test/src/main/resources/logo.png" );
        final CreationHelper helper2 = workbook.getCreationHelper();
        final Drawing drawing2 = resultSheet.createDrawingPatriarch();

        final ClientAnchor anchor2 = helper.createClientAnchor();
        anchor.setAnchorType( ClientAnchor.AnchorType.MOVE_AND_RESIZE );


        final int pictureIndex2 =
                workbook.addPicture(IOUtils.toByteArray(stream), Workbook.PICTURE_TYPE_PNG);


        anchor2.setRow1(26*rowId);
        anchor2.setCol1(0);
        anchor2.setRow2(26*rowId+1);
        anchor2.setCol2(1);
        anchor.setCol1( 11 );
        anchor.setRow1( 5 + rowId*26);
        anchor.setRow2( 11 + rowId*26);
        anchor.setCol2( 12 );
        final Picture pict = drawing.createPicture( anchor, pictureIndex );
        pict.resize(0.8);
    }
    public static void main(String[] args) {
        getConfig();
        try
        {
            FileInputStream file = new FileInputStream(new File("/Users/harsha/IdeaProjects/test/src/main/resources/Candidate09RegistrationLatest.xlsx"));
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFSheet newSheet = workbook.getSheetAt(1);

            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            int i = 1;
            int rowid = 26;
            XSSFSheet resultSheet = workbook.createSheet();
            while (rowIterator.hasNext())
            {
                if(i > 10)
                    break;
                Row row = rowIterator.next();
                //XSSFSheet templateSheet = workbook.cloneSheet(1);
                for(int j = 0; j < 23; j++) {
                    CopyRow.copyRow(workbook, newSheet, j, rowid);
                    rowid++;
                }
                updateDetails(workbook, sheet, newSheet, i, i);
                i++;
                rowid += 3;
            }
            FileOutputStream out = new FileOutputStream(new File("updated_list.xlsx"));
            workbook.write(out);
            out.close();
            file.close();
        }
        catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}