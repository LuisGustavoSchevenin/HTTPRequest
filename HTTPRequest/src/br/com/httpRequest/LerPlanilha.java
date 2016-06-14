package br.com.httpRequest;

import java.io.File;
import java.io.IOException;

import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;


public class LerPlanilha {

	private int contador;
	private Cell[] celulasColLogo;
	private Cell[] celulasColRet;
	//private String counteudoColLogo;
	//private String counteudoColRet;
	private int linhasPlanilha;
	private int retornoRequest;
	private String codRetorno;
	Workbook planilha;
	File arquivo;
	WorkbookSettings ws;
	WritableWorkbook wbCopia;
	WritableSheet escreverExcel;
	WritableCell escreverCelula;
	String rm1 = "https://www";
	String rm2 = "http://www";
	String rm3 = "https://";
	String rm4 = "http://";
	String rm5 = "www";

	public LerPlanilha() {
		try {
			init();
		} catch (BiffException e) {
			e.printStackTrace();
			;
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public void init() throws BiffException, IOException {

		try {
			arquivo = new File("FilesExcel/NovembroRetornoURL2003.xls");
			planilha = Workbook.getWorkbook(arquivo);
			wbCopia = Workbook.createWorkbook(new File("RetornoURL.xls"), planilha);
			escreverExcel = wbCopia.getSheet(0);

			//ws = new WorkbookSettings();
			
			//PEGA A PRIMEIRA ABA DA PLANILHA
			Sheet abaPlanilha = planilha.getSheet(0);

			//PEGA O NÚMERO DE LINHAS DA PLANILHA
			linhasPlanilha = abaPlanilha.getRows();

			celulasColLogo = new Cell[linhasPlanilha];
			celulasColRet = new Cell[linhasPlanilha];
			
			for (contador = 1; contador < linhasPlanilha; contador++) {

				celulasColLogo[contador] = abaPlanilha.getCell(3, contador);
				celulasColRet[contador] = abaPlanilha.getCell(4, contador);

				//counteudoColLogo = celulasColLogo[contador].getContents();
				//counteudoColRet = celulasColRet[contador].getContents();
			}
		} catch (Exception e) {
			throw e;
		}
	}

	public void doRequestURL() throws IOException, WriteException {

		try {
			for (contador = 1062 ; contador < linhasPlanilha ; contador++) {

				URL urlRet = new URL(celulasColRet[contador].getContents().toString());

				HttpURLConnection conURL = (HttpURLConnection) urlRet.openConnection();

				conURL.setRequestMethod("GET");
				conURL.setConnectTimeout(90000);
				conURL.setReadTimeout(90000);

				try {
					conURL.connect();
					retornoRequest = conURL.getResponseCode();
					escreveRetornoRequestURL(retornoRequest, contador);
				} catch (MalformedURLException e) {
					continue;
				} catch (Exception e) {
					continue;
				}
				
				System.out.println(contador);
				//System.out.println(urlRet.toString());
				//System.out.println(conURL.getResponseCode());
			}
		} catch (Exception e) {
			throw e;
		} finally {
			
			wbCopia.write();
			System.out.println("ESCRITO");
			wbCopia.close();
			System.out.println("FECHADO");
		}
	}

	private void escreveRetornoRequestURL(int codigoRetorno, int contador) throws Exception {

		try {
			planilha.close();
			
			//for (contador = 1; contador < linhasPlanilha ; contador++) {
				if (codigoRetorno > 0) {

					codRetorno = Integer.toString(codigoRetorno);
					Label texto = new Label(6, contador, codRetorno);
					escreverExcel.addCell(texto);

				} else {
					Label texto = new Label(6, contador, "Ñ OK");
					escreverExcel.addCell(texto);
				}

			//}
		} catch (Exception e) {
			throw e;
		}/* finally {
			System.out.println("Copia Realizada");
		}*/
	}
	
	public void diffURL() throws Exception {

		try {
			planilha.close();
			
			for (contador = 1; contador < linhasPlanilha ; contador++) {
				if (celulasColLogo[contador].getContents() != celulasColRet[contador].getContents()) {

					WritableCellFormat escreverLinha = new WritableCellFormat();
					Label texto = new Label(5, contador, "DIFERENTE", escreverLinha);
					escreverExcel.addCell(texto);

				} else {
					WritableCellFormat escreverLinha = new WritableCellFormat();
					Label texto = new Label(5, contador, "IGUAL", escreverLinha);
					escreverExcel.addCell(texto);
				}

				wbCopia.write();
			}
		} catch (Exception e) {
			throw e;
		} finally {
			wbCopia.close();
			System.out.println("Copia Realizada");
		}
	}
}