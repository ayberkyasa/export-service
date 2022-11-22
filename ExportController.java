package com.msu.tt.claim.controller;

import com.itextpdf.text.DocumentException;
import com.msu.tt.base.BaseDataResponse;
import com.msu.tt.claim.util.Export;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import lombok.RequiredArgsConstructor;
import org.apache.commons.math3.analysis.function.Exp;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;

@RestController
@RequiredArgsConstructor
@Api(value = "dosya indirme işlemi için controller")
@RequestMapping(value = "/api/v1")
public class ExportController {

    @PostMapping(value = "/export/form-to-excel")
    @ApiOperation(value = "Form'u Excel'e kaydetme")
    public void exportFormToExcel(@RequestBody HashMap<String, Object> map, HttpServletResponse response) throws IOException, DocumentException {
        Export.exportToExcel(map, response);
    }

    @PostMapping(value = "/export/datatable-to-excel")
    @ApiOperation(value = "Datatable'ı Excel'e kaydetme")
    public void exportDatatableToExcel(@RequestBody List<HashMap<String, Object>> mapList, HttpServletResponse response) throws IOException {
        Export.exportToExcel(mapList, response);
    }

    @PostMapping(value = "/export/datatables-to-excel")
    @ApiOperation(value = "Datatable'ları Excel'e kaydetme")
    public void exportDatatablesToExcel(@RequestBody List<List<HashMap<String, Object>>> mapList, HttpServletResponse response) throws IOException {
        Export.exportMultipleDatatableToExcel(mapList, response);
    }

    @PostMapping(value = "/export/form-to-pdf")
    @ApiOperation(value = "Form'u PDF'e kaydetme")
    public void exportFormToPdf(@RequestBody HashMap<String, Object> map, HttpServletResponse response) throws IOException, DocumentException {
        Export.exportToPdf(map, response);
    }

    @PostMapping(value = "/export/datatable-to-pdf")
    @ApiOperation(value = "Datatable'ı PDF'e kaydetme")
    public void exportDatatableToPdf(@RequestBody List<HashMap<String, Object>> mapList, HttpServletResponse response) throws IOException, DocumentException {
        Export.exportToPdf(mapList, response);
    }
}
