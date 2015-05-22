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

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;

import org.dbflute.utflute.core.PlainTestCase;

/**
 * @author jflute
 */
public class XlsReporterTest extends PlainTestCase {

    public void test_demo() throws Exception {
        // ## Arrange ##
        Map<String, Object> dataMap = newHashMap();
        XlsBook book = new XlsReporter().reportBook("report/simple-report.xls", dataMap);
        File output = new File(getTestCaseBuildDir().getParentFile().getCanonicalPath() + "/simple-output.xls");
        if (output.exists()) {
            output.delete();
        }
        assertFalse(output.exists());

        // ## Act ##
        book.write(new FileOutputStream(output));

        // ## Assert ##
        assertTrue(output.exists());
        output.delete();
    }
}
