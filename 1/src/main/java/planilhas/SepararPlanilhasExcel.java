package planilhas;

import com.github.pjfanning.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.ss.usermodel.DateUtil;
import java.text.SimpleDateFormat;
import java.text.DecimalFormat;

import java.io.*;
import java.util.Calendar;
import java.util.Date;

public class SepararPlanilhasExcel {

    public static void main(String[] args) {
        String caminhoArquivoOriginal = "C:/users/ee/downloads/Fonte de Dados - HAM.xlsx";
        String pastaDestino = "C:/users/ee/documents/";

        File arquivo = new File(caminhoArquivoOriginal);
        if (!arquivo.exists()) {
            System.out.println("Arquivo não encontrado: " + arquivo.getAbsolutePath());
            return;
        } else {
            System.out.println("Arquivo encontrado com sucesso!");
        }

        try (InputStream is = new FileInputStream(caminhoArquivoOriginal);
             Workbook workbookOriginal = StreamingReader.builder()
                     .rowCacheSize(100)
                     .bufferSize(4096)
                     .open(is)) {

            for (Sheet planilhaOriginal : workbookOriginal) {
                String nomePlanilha = planilhaOriginal.getSheetName();
                System.out.println("Processando: " + nomePlanilha);

                Workbook novoWorkbook = new SXSSFWorkbook();
                Sheet novaPlanilha = novoWorkbook.createSheet(nomePlanilha);

                copiarConteudo(planilhaOriginal, novaPlanilha);

                String novoArquivo = pastaDestino + nomePlanilha + ".xlsx";
                try (FileOutputStream fos = new FileOutputStream(novoArquivo)) {
                    novoWorkbook.write(fos);
                }

                novoWorkbook.close();
                System.out.println("Planilha '" + nomePlanilha + "' salva em: " + novoArquivo);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void copiarConteudo(Sheet origem, Sheet destino) {
        int linhaDestinoIndex = 1; // começa em 1 para deixar espaço para o cabeçalho

        for (Row linhaOrigem : origem) {
            if (linhaOrigem.getRowNum() == 0) {
                // Copia sempre o cabeçalho
                Row linhaDestino = destino.createRow(0);
                for (Cell celulaOrigem : linhaOrigem) {
                    Cell celulaDestino = linhaDestino.createCell(celulaOrigem.getColumnIndex());
                    copiarValorCelula(celulaOrigem, celulaDestino);
                }
                continue;
            }

            Cell celulaData = linhaOrigem.getCell(2); // Coluna C = índice 2
            if (celulaData == null || celulaData.getCellType() != CellType.NUMERIC ||
                !DateUtil.isCellDateFormatted(celulaData)) {
                continue;
            }

            Date data = celulaData.getDateCellValue();
            Calendar limite = Calendar.getInstance();
            limite.set(2023, Calendar.JANUARY, 1);

            if (data.before(limite.getTime())) {
                continue;
            }

            // Copia linha válida
            Row linhaDestino = destino.createRow(linhaDestinoIndex++);
            for (Cell celulaOrigem : linhaOrigem) {
                Cell celulaDestino = linhaDestino.createCell(celulaOrigem.getColumnIndex());
                copiarValorCelula(celulaOrigem, celulaDestino);
            }
        }
    }

           
        
    

    private static void copiarValorCelula(Cell origem, Cell destino) {
        CellType tipo = origem.getCellType();
        if (tipo == CellType.FORMULA) {
            tipo = origem.getCachedFormulaResultType(); // pega o tipo do resultado da fórmula
        }

        switch (tipo) {
        
            case STRING:
                destino.setCellValue(origem.getStringCellValue());
                break;
                
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(origem)) {
                    SimpleDateFormat formatoData = new SimpleDateFormat("dd/MM/yyyy");
                    destino.setCellValue(formatoData.format(origem.getDateCellValue()));
                } else {
                    int coluna = origem.getColumnIndex();
                    if (coluna == 0 || coluna == 7) {
                        // Coluna H: formatar como inteiro sem decimais
                        destino.setCellValue((int) origem.getNumericCellValue());
                    } else {
                        DecimalFormat formatoMoeda = new DecimalFormat("#,##0.00");
                        destino.setCellValue(formatoMoeda.format(origem.getNumericCellValue()));
                    }
                }
                break;
                
            case BOOLEAN:
                destino.setCellValue(origem.getBooleanCellValue());
                break;
            case BLANK:
                destino.setBlank();
                break;
            default:
                destino.setCellValue(origem.toString());
        }
    }
}