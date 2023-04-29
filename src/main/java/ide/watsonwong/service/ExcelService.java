package ide.watsonwong.service;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelService {

    private int numColumn;

    public Workbook readExcel(String fileLocation) throws Exception {
        FileInputStream file = null;
        try {
            file = new FileInputStream(new File(fileLocation));
            Workbook workbook = new XSSFWorkbook(file);
            return workbook;
        } catch (FileNotFoundException ef) {
            ef.printStackTrace();
            throw ef;
        } catch (IOException eio) {
            eio.printStackTrace();
            throw eio;
        }
    }

    public Workbook readExcel(File file) throws Exception {
        try {
            Workbook workbook = new XSSFWorkbook(file);
            return workbook;
        } catch (IOException eio) {
            eio.printStackTrace();
            throw eio;
        } catch (InvalidFormatException eif) {
            eif.printStackTrace();
            throw eif;
        }
    }

    public Sheet openSheet(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        return sheet;
    }

    public Sheet openSheet(Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(0);
        return sheet;
    }



    public List<Object> getFirstRow(Sheet sheet, int before, int after) {
        System.out.println("<<<getFirstRow>>>" );
        Row firstRow = sheet.getRow(0);
        numColumn =firstRow.getPhysicalNumberOfCells();
        List<Object> titles = new ArrayList<Object>();
        List<String> item = new ArrayList<String>();
        int cellNum = 1;
        for(Cell cell:firstRow) {
            CellType ct = cell.getCellType();
//            System.out.println(cell.getRichStringCellValue() );
            if(cellNum>before && cellNum <= numColumn-after) {
                item.add(cell.getRichStringCellValue().getString());
//                System.out.println(numColumn + "--------------item--------------" + cellNum);
                if(cellNum == numColumn - after) {
                    titles.add(item);
                    System.out.println("===STOP===");
                }
            } else {
                titles.add(cell.getRichStringCellValue().getString());
//                System.out.println(numColumn + "-------------basic--------------" + cellNum);
            }
            cellNum ++;

        }
        System.out.println("<<<getFirstRow>>> -- End" );
        return titles;
    }

    public List<List<HashMap>> getDataRow(Sheet sheet, int before, int after) {
        System.out.println("<<<getFirstRow>>>" + numColumn);
        List<List<HashMap>> data = new ArrayList<List<HashMap>>();
        for (Row row : sheet) {
            if(row.getRowNum() == 0) {
                System.out.println("It is first row");
            } else {
                List<HashMap> rowData = new ArrayList<HashMap>();
                List<Integer> itemQty = new ArrayList<Integer>();
                for (int cellNum = 1; cellNum <= numColumn; cellNum ++) {
                    Cell cell = row.getCell(cellNum-1, Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
                    HashMap<String, Object> map = new HashMap<String,Object>();
                    try {
                        CellType ct;
                        try {
                            ct = cell.getCellType();
//                            System.out.println((row.getRowNum() + 1) + "||" + cellNum + "||" + ct.toString());
                        } catch (NullPointerException e) {
//                            System.out.println((row.getRowNum() + 1) + "||" + cellNum + "|| null");
                            throw e;
                        }
                        switch (ct) {
                            case STRING: {
                                String str = cell.getRichStringCellValue().getString();
//                                System.out.println("String value: " + str);
                                map.put("string", str);
                                rowData.add(map);
//                                System.out.println(numColumn + "-------------basic--------------" + cellNum);
                                break;
                            }
                            case NUMERIC: {
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    Date date = cell.getDateCellValue();
//                                    System.out.println("Date value: " + date);
                                    map.put("date", date);
                                    rowData.add(map);
//                                    System.out.println(numColumn + "-------------basic--------------" + cellNum);
                                } else if (cellNum > 5 && cellNum <= numColumn - after) {
                                    Double num = cell.getNumericCellValue();
//                                    System.out.println("Number value: " + num);
                                    itemQty.add(num.intValue());
//                                    System.out.println(numColumn + "--------------item--------------" + cellNum);
                                    if(cellNum == numColumn - after) {
                                        map.put("item", itemQty);
                                        rowData.add(map);
//                                        System.out.println("===STOP===");
                                    }
                                } else {
                                    cell.setCellType(CellType.STRING);
                                    String str = cell.getRichStringCellValue().getString();
//                                    System.out.println("String value: " + str);
                                    map.put("string", str);
                                    rowData.add(map);
//                                    System.out.println(numColumn + "-------------basic--------------" + cellNum);
                                }
                                break;
                            }
                            case BOOLEAN:
                                break;
                            case FORMULA: {
                                String str = cell.getCellFormula();
//                                System.out.println("Formula value: " + str);
                                map.put("formula", str);
                                rowData.add(map);
//                                System.out.println(numColumn + "-------------basic--------------" + cellNum);
                                break;
                            }
                            case BLANK: {
                                String str = "";
//                                System.out.println("Formula value: " + str);
                                map.put("blank", str);
                                rowData.add(map);
//                                System.out.println(numColumn + "-------------basic--------------" + cellNum);
                                break;
                            }
                            default:
                                break;
                        }
                    } catch (NullPointerException eN) {
//                        System.out.println("Null value");
                        String str = "";
                        map.put("null", str);
                        rowData.add(map);
//                        System.out.println(numColumn + "-------------basic--------------" + cellNum);
                    }

                }
                data.add(rowData);
            }
        }

        return data;
    }



    private void workBookToFile(XSSFWorkbook workbook,String filePath) {
        try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
            workbook.write(outputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void writeExcel(List<Object> firstRow, List<List<HashMap>> data,
                           String sheetName, String firstCol, String secondCol, int before, int after,
                           String filePath) {
        Sheet sheet = null;
//        try {
//            sheet = this.openSheet(this.readExcel(filePath));
//        } catch (Exception e) {
//            e.printStackTrace();
//        }

        XSSFWorkbook workbook = new XSSFWorkbook();
        String idName = sheetName.substring(0, sheetName.indexOf("-"));
        String formName = sheetName.substring(sheetName.indexOf("-") + 1, sheetName.length());
        sheet = workbook.createSheet(sheetName);
        int rowIndex =0;
        Row titleRow = sheet.createRow(rowIndex++);
        int cellIndex = 0;
        Cell cell1 = titleRow.createCell(cellIndex++);
        cell1.setCellValue(firstCol);
        Cell cell2 = titleRow.createCell(cellIndex++);
        cell2.setCellValue(secondCol);
        List<String> items = null;
        for(Object title: firstRow) {
            Cell cell =titleRow.createCell(cellIndex++);
            if(title instanceof List) {
                cell.setCellValue("Item");
                items = (List<String>)title;
                Cell cellQty =titleRow.createCell(cellIndex++);
                cellQty.setCellValue("Item Qty");
            } else {
                cell.setCellValue(title.toString());
            }
        }



        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat((short) 14);
        for(List<HashMap> dataList: data) {
            System.out.println("");
            ArrayList<Integer> itemQtyL = (ArrayList<Integer>) dataList.get(before).get("item");
            for(int count =0; count < items.size(); count++) {
                String itemName = items.get(count);
                int itemQty = itemQtyL.get(count);
                Row dataRow = sheet.createRow(rowIndex);
                int dataIndex = 0;
                Cell data1 = dataRow.createCell(dataIndex++);
                data1.setCellValue(idName + "-" + rowIndex++);
                Cell data2 = dataRow.createCell(dataIndex++);
                data2.setCellValue(formName);
                for(int cellCount = 0; cellCount < dataList.size() ; cellCount ++) {
                    if(cellCount != before) {
                        Cell dataCell = dataRow.createCell(dataIndex++);
                        HashMap<String, Object> cellData = dataList.get(cellCount);
                        for ( String key : cellData.keySet() ) {
                            if("date".equalsIgnoreCase(key)) {
                                Date date = (Date) cellData.get(key);
                                dataCell.setCellValue(date);
                                dataCell.setCellStyle(cellStyle);
                            }else if("number".equalsIgnoreCase(key)) {
                                Double doub = (Double) cellData.get(key);
                                dataCell.setCellValue(doub);
                            }else if("string".equalsIgnoreCase(key) ) {
                                String str = (String) cellData.get(key);
                                dataCell.setCellValue(str);
                            }else {
                                dataCell.setCellValue("");
                            }
                            System.out.println("--" + key + "||" + cellData.get(key) );
                        }
                        System.out.println("--------------------------------");
                    } else {
                        Cell dataI = dataRow.createCell(dataIndex++);
                        dataI.setCellValue(itemName);
                        Cell dataQ = dataRow.createCell(dataIndex++);
                        dataQ.setCellValue(itemQty);
                    }

                }
            }



        }

        this.workBookToFile(workbook, filePath);
    }


}
