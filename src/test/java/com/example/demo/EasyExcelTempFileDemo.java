package com.example.demo;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.util.FileUtils;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.handler.WorkbookWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import com.alibaba.excel.write.handler.context.WorkbookWriteHandlerContext;
import com.alibaba.excel.util.ListUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.Date;
import java.util.List;

public class EasyExcelTempFileDemo {

    // ====================== 不压缩临时文件测试 ======================
    @Test
    public void uncompressedTemporaryFile() throws Exception {
        System.out.println("临时文件存储路径: " +FileUtils.getPoiFilesPath());
//        log.info("临时文件存储路径: {}", FileUtils.getPoiFilesPath());

        File outputFile = createOutputFile("uncompressed_" + System.currentTimeMillis() + ".xlsx");

        ExcelWriter excelWriter = null;
        try {
            excelWriter = EasyExcel.write(outputFile, DemoData.class).build();
            WriteSheet writeSheet = EasyExcel.writerSheet("测试数据").build();
            writeData(excelWriter, writeSheet, 10000); // 写入10万条数据（分10000次，每次10条）
//            log.info("未压缩临时文件写入完成");
            System.out.println("压缩临时文件写入完成");
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }
    }

    // ====================== 压缩临时文件测试 ======================
    @Test
    public void compressedTemporaryFile() throws Exception {
        System.out.println("临时文件存储路径: " +FileUtils.getPoiFilesPath());
        File outputFile = createOutputFile("compressed_" + System.currentTimeMillis() + ".xlsx");

        ExcelWriter excelWriter = null;
        try {
            excelWriter = EasyExcel.write(outputFile, DemoData.class)
                    .registerWriteHandler(new CompressTempFileHandler()) // 注册压缩处理器
                    .build();
            WriteSheet writeSheet = EasyExcel.writerSheet("测试数据").build();
            writeData(excelWriter, writeSheet, 10000); // 写入10万条数据
//            log.info("压缩临时文件写入完成");
            System.out.println("压缩临时文件写入完成");
        } finally {
            if (excelWriter != null) {
                excelWriter.finish();
            }
        }
    }

    // ====================== 通用工具方法 ======================
    private String getPoiTempPath() {
        // POI临时文件路径（Windows可能为java.io.tmpdir/poifiles，Linux/mac为/var/folders/.../poifiles）
        return System.getProperty("java.io.tmpdir") + File.separator + "poifiles";
    }

    private File createOutputFile(String fileName) {
        // 创建输出文件（可替换为实际存储路径，测试时建议使用临时目录）
        File outputDir = new File("target/test-output");
        outputDir.mkdirs();
        return new File(outputDir, fileName);
    }

    private void writeData(ExcelWriter excelWriter, WriteSheet writeSheet, int pageCount) {
        for (int i = 0; i < pageCount; i++) {
            excelWriter.write(data(), writeSheet);
        }
    }

    private List<DemoData> data() {
        List<DemoData> list = ListUtils.newArrayList();
        for (int i = 0; i < 10; i++) { // 每次写入10条数据
            DemoData data = new DemoData();
            data.setString("字符串_" + i);
            data.setDate(new Date());
            data.setDoubleData(0.56);
            list.add(data);
        }
        return list;
    }

    // ====================== 压缩处理器（内部类） ======================
    private static class CompressTempFileHandler implements WorkbookWriteHandler {
        @Override
        public void afterWorkbookCreate(WorkbookWriteHandlerContext context) {
            WriteWorkbookHolder workbookHolder = context.getWriteWorkbookHolder();
            if (workbookHolder.getWorkbook() instanceof SXSSFWorkbook) {
                SXSSFWorkbook sxssfWorkbook = (SXSSFWorkbook) workbookHolder.getWorkbook();
                sxssfWorkbook.setCompressTempFiles(true); // 关键配置：启用临时文件压缩
            }
        }
    }

    // ====================== 数据实体类 ======================
    private static class DemoData {
        private String string;
        private Date date;
        private Double doubleData;

        // Getter/Setter
        public String getString() { return string; }
        public void setString(String string) { this.string = string; }
        public Date getDate() { return date; }
        public void setDate(Date date) { this.date = date; }
        public Double getDoubleData() { return doubleData; }
        public void setDoubleData(Double doubleData) { this.doubleData = doubleData; }
    }
}