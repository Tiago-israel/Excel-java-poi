package br.com.poi.controller;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import br.com.poi.model.Pessoa;

@RestController
@RequestMapping("pessoas")
public class PessoaController {

	@PostMapping
	public byte[] gerarRelatorio(@RequestBody List<Pessoa> lista) throws IOException {
		String sheetName = "teste";
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
		HSSFSheet sheet = hssfWorkbook.createSheet(sheetName);

		for (int i = 0; i < lista.size(); i++) {
			HSSFRow row = sheet.createRow(i);
			for (int j = 0; j < 2; j++) {
				HSSFCell cell = row.createCell(j);
				switch (j) {
				case 0:
					if (i == 0) {
						cell.setCellValue("Nome");
					} else {
						cell.setCellValue(lista.get(i).nome);
					}
					break;

				case 1:
					if (i == 0) {
						cell.setCellValue("CPF");
					} else {
						cell.setCellValue(lista.get(i).cpf);
					}
					break;

				default:
					break;
				}
			}
		}

		FileOutputStream fileOut = new FileOutputStream(File.createTempFile("teste", ".xls"));
		ByteArrayOutputStream output = new ByteArrayOutputStream();
		hssfWorkbook.write(output);
		hssfWorkbook.write(fileOut);
		fileOut.flush();
		fileOut.close();
		output.close();
		return output.toByteArray();

	}

}
