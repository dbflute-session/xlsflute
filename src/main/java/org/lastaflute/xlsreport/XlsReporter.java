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
import java.io.InputStream;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.lastaflute.xlsreport.exception.ReportingXlsParseFailureException;
import org.lastaflute.xlsreport.exception.ReportingXlsProcessFailureException;
import org.lastaflute.xlsreport.exception.ReportingXlsTemplateNotFoundException;
import org.seasar.fisshplate.exception.FPMergeException;
import org.seasar.fisshplate.exception.FPParseException;
import org.seasar.fisshplate.template.FPTemplate;

/**
 * @author jflute
 */
public class XlsReporter {

    // TODO jflute lastaflute: [E] function: making xlsreport (2015/05/22)
    public XlsBook reportBook(String templateName, Map<String, Object> dataMap) {
        final FPTemplate template = newFPTemplate();
        final HSSFWorkbook hssf;
        try {
            hssf = template.process(toStream(templateName), dataMap);
        } catch (FPMergeException e) {
            String msg = "Failed to merge the template: " + templateName + ", data=" + dataMap;
            throw new ReportingXlsParseFailureException(msg, e);
        } catch (FPParseException e) {
            String msg = "Failed to parse the template: " + templateName + ", data=" + dataMap;
            throw new ReportingXlsParseFailureException(msg, e);
        } catch (IOException e) {
            String msg = "Failed to process the template: " + templateName + ", data=" + dataMap;
            throw new ReportingXlsProcessFailureException(msg, e);
        }
        return newXlsBook(hssf);
    }

    protected FPTemplate newFPTemplate() {
        return new FPTemplate();
    }

    protected InputStream toStream(String templateName) {
        final ClassLoader loader = Thread.currentThread().getContextClassLoader();
        final InputStream stream = loader.getResourceAsStream(templateName);
        if (stream == null) {
            String msg = "Not found the template in the classpath: " + templateName + " loader=" + loader;
            throw new ReportingXlsTemplateNotFoundException(msg);
        }
        return stream;
    }

    protected XlsBook newXlsBook(HSSFWorkbook hssfWorkbook) {
        return new XlsBook(hssfWorkbook);
    }
}
