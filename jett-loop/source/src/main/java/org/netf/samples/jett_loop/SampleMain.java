package org.netf.samples.jett_loop;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;

import net.sf.jett.transform.ExcelTransformer;

/**
 * Template Engine JETT Loop Sample
 */
public class SampleMain {

	public static void main(String[] args) throws Exception {

		ClassLoader loader = Thread.currentThread().getContextClassLoader();
		String path = loader.getResource("output").toURI().getPath();

		ExcelTransformer transformer = new ExcelTransformer();
		Map<String, Object> param = new HashMap<>();

		try (InputStream in = loader.getResourceAsStream("template/sample-template.xlsx");
				OutputStream out = new FileOutputStream(new File(path, "sample.xlsx"));) {

			Workbook workbook = transformer.transform(in,
					Arrays.asList("loop"), // テンプレートのシート名
					Arrays.asList("result"), // 出力ファイルのシート名
					Arrays.asList(param) // パラメータ
			);

			workbook.write(out);

			workbook.close();
		}

	}

}
