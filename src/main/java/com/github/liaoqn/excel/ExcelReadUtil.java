package com.github.liaoqn.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.builder.ExcelReaderSheetBuilder;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.github.liaoqn.excel.dto.ColumnData;
import com.github.liaoqn.excel.dto.ExcelReadContext;
import com.github.liaoqn.excel.dto.ExcelSheet;
import com.github.liaoqn.excel.dto.SheetColumnData;
import com.github.liaoqn.excel.enums.FileType;
import com.github.liaoqn.excel.listener.ExcelHeaderListener;
import com.github.liaoqn.excel.listener.ExcelMetadataListener;
import com.github.liaoqn.excel.listener.ExcelPageReadListener;
import com.github.liaoqn.excel.listener.ExcelReadListener;
import org.apache.commons.collections4.CollectionUtils;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import java.util.stream.Collectors;

/**
 * 读取Excel
 *
 * @author liaoqn
 * @date 2023/4/14
 */
public class ExcelReadUtil {

    /**
     * 读取excel sheet页数据
     *
     * @param excelReadContext
     * @return excel数据，结果与columnDataList顺序保持一致，并按照columnDataList传入的类型转换数据类型，类型转换失败取null
     */
    public static List<ArrayList<Object>> readSheet(ExcelReadContext excelReadContext) throws IOException {
        if (excelReadContext.getInputStream() == null) {
            return new ArrayList<>();
        }

        Integer offset = excelReadContext.getOffset();
        if (offset == null || offset < 1) {
            offset = 1;
        }

        ExcelReadListener listener = new ExcelReadListener(excelReadContext.getColumnDataList(), excelReadContext.getSize());
        try (InputStream inputStream = convertFileInputStream(excelReadContext.getInputStream())) {
            ExcelReaderSheetBuilder readerSheetBuilder = EasyExcel.read(inputStream, listener)
                    .autoCloseStream(false)
                    .excelType(convertExcelType(excelReadContext.getFileType()))
                    .sheet(excelReadContext.getSheetNo())
                    .headRowNumber(offset);
            readerSheetBuilder.doRead();
            return listener.getDataList();
        }
    }

    /**
     * 分页读取
     *
     * @param is                  excel文件流
     * @param fileType            excel类型
     * @param sheetColumnDataList Excel sheet页列信息
     * @param pageSize            分页大小
     * @param pageConsumer        读取到pageSize条数据后，触发回调
     */
    public static void readPage(InputStream is, FileType fileType, List<SheetColumnData> sheetColumnDataList,
                                int pageSize, Consumer<List<ArrayList<Object>>> pageConsumer) throws IOException {
        if (is == null) {
            return;
        }

        List<ReadSheet> readSheets = new ArrayList<>();
        for (SheetColumnData sheetColumnData : sheetColumnDataList) {
            ExcelPageReadListener listener = new ExcelPageReadListener(sheetColumnData.getColumnDataList(), pageSize, pageConsumer);

            ReadSheet readSheet = EasyExcel.readSheet(sheetColumnData.getSheetNo())
                    .headRowNumber(1)
                    .registerReadListener(listener).build();
            readSheets.add(readSheet);
        }

        ExcelTypeEnum excelType = convertExcelType(fileType);
        try (InputStream inputStream = convertFileInputStream(is); ExcelReader excelReader = EasyExcel.read(inputStream).excelType(excelType).build()) {
            excelReader.read(readSheets);
        }
    }

    /**
     * 获取excel的sheet页信息
     *
     * @param is
     * @param fileType
     * @return
     */
    public static List<ExcelSheet> readSheetMetaData(InputStream is, FileType fileType) throws IOException {
        try (InputStream inputStream = convertFileInputStream(is);
             ExcelReader excelReader = EasyExcel.read(inputStream).excelType(convertExcelType(fileType)).build()) {
            List<ReadSheet> readSheets = excelReader.excelExecutor().sheetList();
            excelReader.finish();

            return CollectionUtils.emptyIfNull(readSheets).stream().map(sheet -> {
                ExcelSheet excelSheet = new ExcelSheet();
                excelSheet.setSheetNo(sheet.getSheetNo());
                excelSheet.setSheetName(sheet.getSheetName());
                return excelSheet;
            }).collect(Collectors.toList());
        }
    }

    /**
     * 读取excel sheet页的列信息，并推测字段类型
     *
     * @param is                excel文件流
     * @param fileType
     * @param sheetNo
     * @param excelMetadataRows 使用前多少行用于推测字段类型，默认前10行
     * @return
     */
    public static SheetColumnData readColumnMetaData(InputStream is, FileType fileType, int sheetNo, Integer excelMetadataRows) throws IOException {
        List<SheetColumnData> sheetColumnDataList = readColumnMetaData(is, fileType, Collections.singletonList(sheetNo), excelMetadataRows);
        if (CollectionUtils.isEmpty(sheetColumnDataList)) {
            return null;
        }

        return sheetColumnDataList.get(0);
    }

    /**
     * 读取excel sheet页的列信息
     *
     * @param is                excel文件流
     * @param fileType
     * @param sheetNos
     * @param excelMetadataRows 使用前多少行用于推测字段类型，默认前10行
     * @return
     */
    public static List<SheetColumnData> readColumnMetaData(InputStream is, FileType fileType, List<Integer> sheetNos, Integer excelMetadataRows) throws IOException {
        List<SheetColumnData> sheetColumnDataList = new ArrayList<>();
        if (is == null || CollectionUtils.isEmpty(sheetNos)) {
            return sheetColumnDataList;
        }

        List<ReadSheet> readSheets = new ArrayList<>();
        List<ExcelMetadataListener> listeners = new ArrayList<>();
        for (Integer sheetNo : sheetNos) {
            if (sheetNo == null || sheetNo < 0) {
                continue;
            }

            ExcelMetadataListener listener = new ExcelMetadataListener(excelMetadataRows, sheetNo);
            ReadSheet readSheet = EasyExcel.readSheet(sheetNo).headRowNumber(1).registerReadListener(listener).build();
            readSheets.add(readSheet);
            listeners.add(listener);
        }

        try (InputStream inputStream = convertFileInputStream(is);
             ExcelReader excelReader = EasyExcel.read(inputStream).excelType(convertExcelType(fileType)).build()) {
            excelReader.read(readSheets);
        }

        for (ExcelMetadataListener listener : listeners) {
            SheetColumnData sheetColumnData = new SheetColumnData();
            sheetColumnData.setSheetNo(listener.getSheetNo());
            sheetColumnData.setColumnDataList(listener.getColumnDataList());
            sheetColumnDataList.add(sheetColumnData);
        }

        return sheetColumnDataList;
    }

    public static Map<Integer, List<ColumnData>> readColumnHeader(InputStream is, FileType fileType, List<Integer> sheetNos) throws IOException {
        Map<Integer, List<ColumnData>> columnHeaders = new LinkedHashMap<>();
        if (is == null || CollectionUtils.isEmpty(sheetNos)) {
            return columnHeaders;
        }

        List<ReadSheet> readSheets = new ArrayList<>();
        Map<Integer, ExcelHeaderListener> listenerMap = new LinkedHashMap<>();
        for (Integer sheetNo : sheetNos) {
            if (sheetNo == null || sheetNo < 0) {
                continue;
            }

            ExcelHeaderListener listener = new ExcelHeaderListener();
            ReadSheet readSheet = EasyExcel.readSheet(sheetNo).headRowNumber(1).registerReadListener(listener).build();
            readSheets.add(readSheet);
            listenerMap.put(sheetNo, listener);
        }

        ExcelTypeEnum excelType = convertExcelType(fileType);
        try (InputStream inputStream = convertFileInputStream(is);
             ExcelReader excelReader = EasyExcel.read(inputStream).excelType(excelType).build()) {
            excelReader.read(readSheets);
        }

        listenerMap.forEach((sheetNo, listener) -> columnHeaders.put(sheetNo, listener.getColumnDataList()));
        return columnHeaders;
    }

    public static FileType convertMimeToFileType(String mime) {
        if ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet".equals(mime)) {
            return FileType.XLSX;
        }
        if ("application/vnd.ms-excel".equals(mime)) {
            return FileType.XLS;
        }
        if ("text/csv".equals(mime)) {
            return FileType.CSV;
        }
        return null;
    }

    private static ExcelTypeEnum convertExcelType(FileType fileType) {
        switch (fileType) {
            case XLS:
                return ExcelTypeEnum.XLS;
            case CSV:
                return ExcelTypeEnum.CSV;
            case XLSX:
            default:
                return ExcelTypeEnum.XLSX;
        }
    }

    /**
     * poi使用FileInputStream读取excel会使用堆外内存，并且不会直接释放，容易出现堆外内存泄漏。这里转成BufferedInputStream
     *
     * @param inputStream
     * @return
     */
    private static InputStream convertFileInputStream(InputStream inputStream) {
        if (inputStream instanceof FileInputStream) {
            return new BufferedInputStream(inputStream);
        }
        return inputStream;
    }
}
