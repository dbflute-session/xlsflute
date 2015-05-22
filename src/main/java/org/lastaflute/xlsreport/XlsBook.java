/*
 * Copyright 2014-2015 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
 * either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */
package org.lastaflute.xlsreport;

import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.lastaflute.xlsreport.exception.ReportingXlsWriteFailureException;

/**
 * @author jflute
 */
public class XlsBook {

    protected final HSSFWorkbook workbook;

    public XlsBook(HSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public void write(OutputStream stream) {
        try {
            workbook.write(stream);
        } catch (IOException e) {
            String msg = "Failed to write the xls book: " + workbook;
            throw new ReportingXlsWriteFailureException(msg, e);
        } finally {
            try {
                stream.close();
            } catch (IOException ignored) {}
        }
    }

    public void protect(String username, String password) {
        workbook.writeProtectWorkbook(password, username);
    }

    public HSSFWorkbook getWorkbook() {
        return workbook;
    }
}
