package com.github.liaoqn.excel;

import com.github.liaoqn.excel.dto.ColumnData;
import com.github.liaoqn.excel.dto.ExcelReadContext;
import com.github.liaoqn.excel.dto.ExcelSheet;
import com.github.liaoqn.excel.dto.SheetColumnData;
import com.github.liaoqn.excel.enums.FileType;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.builder.ToStringBuilder;
import org.apache.commons.lang3.builder.ToStringStyle;
import org.junit.Assert;
import org.junit.Test;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * 读取excel
 *
 * @author liaoqn
 * @date 2023/4/18
 */
@Slf4j
public class ExcelReadUtilTest {

    @Test
    public void readSheetXlsxTest() throws IOException {
        Resource resource = new ClassPathResource("excel/test001.xlsx");

        SheetColumnData sheetColumnData = readSheetColumnData(resource, FileType.XLSX, 0);

        try (InputStream is = resource.getInputStream()) {
            int pageSize = 5;
            ExcelReadContext readContext = new ExcelReadContext();
            readContext.setInputStream(is);
            readContext.setFileType(FileType.XLSX);
            readContext.setSheetNo(0);
            readContext.setColumnDataList(sheetColumnData.getColumnDataList());
            readContext.setOffset(1);
            readContext.setSize(pageSize);

            List<ArrayList<Object>> arrayLists = ExcelReadUtil.readSheet(readContext);
            Assert.assertEquals(arrayLists.size(), pageSize);
        }
    }

    @Test
    public void readPageXlsxTest() throws IOException {
        Resource resource = new ClassPathResource("excel/test001.xlsx");

        List<SheetColumnData> sheetColumnDataList = readSheetColumnData(resource, FileType.XLSX, Arrays.asList(0, 1));

        try (InputStream is = resource.getInputStream()) {
            AtomicBoolean success = new AtomicBoolean(false);
            ExcelReadUtil.readPage(is, FileType.XLSX, sheetColumnDataList, 2, dataList -> {
                success.set(true);
                log.info(ToStringBuilder.reflectionToString(dataList, ToStringStyle.JSON_STYLE, true, Object.class));
            });
            Assert.assertTrue(success.get());
        }
    }

    @Test
    public void readPageOnlyTimeTest() throws IOException {
        Resource resource = new ClassPathResource("excel/only_time.xlsx");

        List<SheetColumnData> sheetColumnDataList = readSheetColumnData(resource, FileType.XLSX, Arrays.asList(0, 1));

        try (InputStream is = resource.getInputStream()) {
            AtomicBoolean success = new AtomicBoolean(false);
            ExcelReadUtil.readPage(is, FileType.XLSX, sheetColumnDataList, 2, dataList -> {
                success.set(true);
                log.info(ToStringBuilder.reflectionToString(dataList, ToStringStyle.JSON_STYLE, true, Object.class));
            });
            Assert.assertTrue(success.get());
        }
    }

    @Test
    public void readSheetMetaDataXlsxTest() throws IOException {
        Resource resource = new ClassPathResource("excel/test001.xlsx");
        try (InputStream is = resource.getInputStream()) {
            List<ExcelSheet> excelSheets = ExcelReadUtil.readSheetMetaData(is, FileType.XLSX);
            Assert.assertEquals(excelSheets.size(), 2);
        }
    }

    @Test
    public void readColumnMetaDataXlsxTest() throws IOException {
        Resource resource = new ClassPathResource("excel/test001.xlsx");
        try (InputStream is = resource.getInputStream()) {
            SheetColumnData sheetColumnData = ExcelReadUtil.readColumnMetaData(is, FileType.XLSX, 0, 10);
            Assert.assertNotNull(sheetColumnData);
            log.info(ToStringBuilder.reflectionToString(sheetColumnData));
            Assert.assertEquals(sheetColumnData.getColumnDataList().size(), 7);
        }
    }

    @Test
    public void readXlsxColumnHeaderXlsxTest() throws IOException {
        Resource resource = new ClassPathResource("excel/test001.xlsx");
        try (InputStream is = resource.getInputStream()) {
            Map<Integer, List<ColumnData>> columnHeaders = ExcelReadUtil.readColumnHeader(is, FileType.XLSX, Arrays.asList(1, 0));
            log.info(ToStringBuilder.reflectionToString(columnHeaders));
            Assert.assertEquals(columnHeaders.size(), 2);
        }
    }

    @Test
    public void readSheetXlsTest() throws IOException {
        Resource resource = new ClassPathResource("excel/test001.xls");

        SheetColumnData sheetColumnData = readSheetColumnData(resource, FileType.XLS, 0);

        try (InputStream is = resource.getInputStream()) {
            int pageSize = 5;
            ExcelReadContext readContext = new ExcelReadContext();
            readContext.setInputStream(is);
            readContext.setFileType(FileType.XLS);
            readContext.setSheetNo(0);
            readContext.setColumnDataList(sheetColumnData.getColumnDataList());
            readContext.setOffset(1);
            readContext.setSize(pageSize);
            List<ArrayList<Object>> arrayLists = ExcelReadUtil.readSheet(readContext);
            Assert.assertEquals(arrayLists.size(), pageSize);
        }
    }

    @Test
    public void readPageXlsTest() throws IOException {
        Resource resource = new ClassPathResource("excel/test001.xls");

        List<SheetColumnData> sheetColumnDataList = readSheetColumnData(resource, FileType.XLS, Arrays.asList(0, 1));

        try (InputStream is = resource.getInputStream()) {
            AtomicBoolean success = new AtomicBoolean(false);
            ExcelReadUtil.readPage(is, FileType.XLS, sheetColumnDataList, 2, dataList -> {
                success.set(true);
                log.info(ToStringBuilder.reflectionToString(dataList, ToStringStyle.JSON_STYLE, true, Object.class));
            });
            Assert.assertTrue(success.get());
        }
    }

    @Test
    public void readSheetMetaDataXlsTest() throws IOException {
        Resource resource = new ClassPathResource("excel/test001.xls");
        try (InputStream is = resource.getInputStream()) {
            List<ExcelSheet> excelSheets = ExcelReadUtil.readSheetMetaData(is, FileType.XLS);
            Assert.assertEquals(excelSheets.size(), 2);
        }
    }

    @Test
    public void readColumnMetaDataXlsTest() throws IOException {
        Resource resource = new ClassPathResource("excel/test001.xls");
        try (InputStream is = resource.getInputStream()) {
            SheetColumnData sheetColumnData = ExcelReadUtil.readColumnMetaData(is, FileType.XLS, 0, 10);
            Assert.assertNotNull(sheetColumnData);
            log.info(ToStringBuilder.reflectionToString(sheetColumnData));
            Assert.assertEquals(sheetColumnData.getColumnDataList().size(), 7);
        }
    }

    @Test
    public void readColumnHeaderXlsTest() throws IOException {
        Resource resource = new ClassPathResource("excel/test001.xls");
        try (InputStream is = resource.getInputStream()) {
            Map<Integer, List<ColumnData>> columnHeaders = ExcelReadUtil.readColumnHeader(is, FileType.XLS, Arrays.asList(1, 0));
            log.info(ToStringBuilder.reflectionToString(columnHeaders));
            Assert.assertEquals(columnHeaders.size(), 2);
        }
    }

    @Test
    public void readSheetCsvTest() throws IOException {
        Resource resource = new ClassPathResource("excel/test001.csv");

        SheetColumnData sheetColumnData = readSheetColumnData(resource, FileType.CSV, 0);

        try (InputStream is = resource.getInputStream()) {
            int pageSize = 5;
            ExcelReadContext readContext = new ExcelReadContext();
            readContext.setInputStream(is);
            readContext.setFileType(FileType.CSV);
            readContext.setSheetNo(0);
            readContext.setColumnDataList(sheetColumnData.getColumnDataList());
            readContext.setOffset(1);
            readContext.setSize(pageSize);
            List<ArrayList<Object>> arrayLists = ExcelReadUtil.readSheet(readContext);
            Assert.assertEquals(arrayLists.size(), pageSize);
        }
    }

    @Test
    public void readPageCsvTest() throws IOException {
        Resource resource = new ClassPathResource("excel/test001.csv");

        List<SheetColumnData> sheetColumnDataList = readSheetColumnData(resource, FileType.CSV, Arrays.asList(0, 1));

        try (InputStream is = resource.getInputStream()) {
            AtomicBoolean success = new AtomicBoolean(false);
            ExcelReadUtil.readPage(is, FileType.CSV, sheetColumnDataList, 2, dataList -> {
                success.set(true);
                log.info(ToStringBuilder.reflectionToString(dataList, ToStringStyle.JSON_STYLE, true, Object.class));
            });
            Assert.assertTrue(success.get());
        }
    }

    @Test
    public void readSheetMetaDataCsvTest() throws IOException {
        Resource resource = new ClassPathResource("excel/test001.csv");
        try (InputStream is = resource.getInputStream()) {
            List<ExcelSheet> excelSheets = ExcelReadUtil.readSheetMetaData(is, FileType.CSV);
            Assert.assertEquals(excelSheets.size(), 1);
        }
    }

    @Test
    public void readColumnMetaDataCsvTest() throws IOException {
        Resource resource = new ClassPathResource("excel/test001.csv");
        try (InputStream is = resource.getInputStream()) {
            SheetColumnData sheetColumnData = ExcelReadUtil.readColumnMetaData(is, FileType.CSV, 0, 10);
            Assert.assertNotNull(sheetColumnData);
            log.info(ToStringBuilder.reflectionToString(sheetColumnData));
            Assert.assertEquals(sheetColumnData.getColumnDataList().size(), 7);
        }
    }

    private SheetColumnData readSheetColumnData(Resource resource, FileType fileType, int sheetNo) throws IOException {
        SheetColumnData sheetColumnData;
        try (InputStream is = resource.getInputStream()) {
            sheetColumnData = ExcelReadUtil.readColumnMetaData(is, fileType, sheetNo, 10);
        }
        Assert.assertNotNull(sheetColumnData);
        log.info("sheetColumnData={}", ToStringBuilder.reflectionToString(sheetColumnData, ToStringStyle.JSON_STYLE, true, Object.class));
        return sheetColumnData;
    }

    private List<SheetColumnData> readSheetColumnData(Resource resource, FileType fileType, List<Integer> sheetNos) throws IOException {
        List<SheetColumnData> sheetColumnDataList;
        try (InputStream is = resource.getInputStream()) {
            sheetColumnDataList = ExcelReadUtil.readColumnMetaData(is, fileType, sheetNos, 10);
        }
        Assert.assertTrue(CollectionUtils.isNotEmpty(sheetColumnDataList));
        log.info("sheetColumnDataList={}", ToStringBuilder.reflectionToString(sheetColumnDataList, ToStringStyle.JSON_STYLE, true, Object.class));
        return sheetColumnDataList;
    }
}
