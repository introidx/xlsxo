# XLSXO

Generate an XLSX file effortlessly.

## Install Dependencies

Add this to your `pom.xml` file

```
<dependency>
	<groupId>org.apache.poi</groupId>
	<artifactId>poi</artifactId>
	<version>5.0.0</version>
</dependency>

<dependency>
   <groupId>org.apache.poi</groupId>
   <artifactId>poi-ooxml</artifactId>
   <version>5.0.0</version>
</dependency>

```

## Usage
Create a Java class `XlsxGenerator.java` and add this code, call the method 
`generate` with your list of object and Name of your xlsx sheet, this will return you a byte[]

```
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.List;

public class XlsxGenerator {

    public static <T> byte[] generate(List<T> list, String sheetName) throws IOException {
        try {
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet(sheetName);

            Row headerRow = sheet.createRow(0);
            String[] headers = getHeaders(list.get(0));

            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            // Create data rows
            for (int rowIndex = 0; rowIndex < list.size(); rowIndex++) {
                Row dataRow = sheet.createRow(rowIndex + 1);
                T item = list.get(rowIndex);
                Field[] fields = item.getClass().getDeclaredFields();

                for (int colIndex = 0; colIndex < fields.length; colIndex++) {
                    fields[colIndex].setAccessible(true);
                    Cell cell = dataRow.createCell(colIndex);
                    try {
                        Object value = fields[colIndex].get(item);
                        if (value != null) {
                            cell.setCellValue(value.toString());
                        }
                    } catch (IllegalAccessException e) {
                        throw new IOException("Error accessing field value", e);
                    }
                }
            }

            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            workbook.write(bos);
            bos.close();
            workbook.close();
            return bos.toByteArray();
        } catch (Exception e) {
            throw new IOException("Error while converting data to Excel sheet", e);
        }
    }

    private static <T> String[] getHeaders(T object) {
        Field[] fields = object.getClass().getDeclaredFields();
        String[] headers = new String[fields.length];
        for (int i = 0; i < fields.length; i++) {
            headers[i] = convertToReadableFormat(fields[i].getName());
        }
        return headers;
    }

    private static String convertToReadableFormat(String fieldName) {
        StringBuilder readableName = new StringBuilder();
        for (char c : fieldName.toCharArray()) {
            if (Character.isUpperCase(c)) {
                readableName.append(" ");
            }
            readableName.append(c);
        }
        return readableName.substring(0, 1).toUpperCase() + readableName.substring(1);
    }
}
```

## Generate xlsx file

Call it like this
```
@GetMapping("/download-report")
    public ResponseEntity<ByteArrayResource> downloadTDSReport(@RequestHeader(COMPANY_ID_HEADER) String companyId) {
        byte[] bytes = reportService.downloadTDSReport(Long.parseLong(companyId));

        ByteArrayResource resource = new ByteArrayResource(bytes);
        return ResponseEntity.ok()
                .contentLength(bytes.length)
                .contentType(org.springframework.http.MediaType.APPLICATION_OCTET_STREAM)
                .header("Content-disposition", "attachment; filename=\"tds-report.xlsx\"")
                .body(resource);
    }
```

