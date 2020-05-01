package leitor;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;

/**
 *
 * @author Diony
 */
public class Leitor {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // Exemplo de Leitura
        String mapa, nomeTimeA, nomeTimeB, vencedor;
        mapa = readCellData(2, 0, 1);
        nomeTimeA = readCellData(3, 0, 1);
        nomeTimeB = readCellData(3, 0, 2);
        vencedor = readCellData(0, 20, 1);
        System.out.println("Mapa: " + mapa + "\nTime 1: " + nomeTimeA + "\tTime 2: " + nomeTimeB + "\nVencedor: " + vencedor);
        System.out.println(readCellData(0,8,1));
        /* Exemplo de Escrita
        writeCellData(0, 2, 1, "XxXx");
        writeCellData(0, 3, 1, "XxXx");
        writeCellData(0, 4, 1, "XxXx");
        */

    }

    //lembrando que a linha, tabela e coluna começam por 0
    public static String readCellData(int tabela, int coluna, int linha) {
        String saida = "";
        XSSFWorkbook wb = null;
        try {
            FileInputStream arquivo = new FileInputStream("Exemplo\\testeA.xlsx");
            wb = new XSSFWorkbook(arquivo);
        } catch (IOException e) {
        }
        Sheet sheet = wb.getSheetAt(tabela);
        Row row = sheet.getRow(linha);
        Cell cell = row.getCell(coluna);
        DataFormatter formatter = new DataFormatter();
        saida = formatter.formatCellValue(cell);
        // Alguns lugares fala para usar isso, deve evitar problemas com memoria mas funciona sem
        try {
            wb.close();
        } catch (IOException e) {
        }

        return saida;
    }
    
    /*
    Se o arquivo ja estiver aberto por outro programa retorna erro
    */
    public static void writeCellData(int tabela, int coluna, int linha, String conteudo) {
        XSSFWorkbook wb = null;
        Row row;
        try {
            FileInputStream arquivo = new FileInputStream("Exemplo\\testeA.xlsx");
            wb = new XSSFWorkbook(arquivo);
        } catch (IOException e) {
        }
        /* Para criar uma nova tabela(aba)
        Sheet sheet = wb.createSheet("Employee");
         */
        Sheet sheet = wb.getSheetAt(tabela);
        row = sheet.getRow(linha);
        //Conferir se a linha existe
        if(row == null){
            row = sheet.createRow(linha);
        }
        Cell cell = row.getCell(coluna);
        //Conferir se a celula existe
        if(cell == null){
            cell = row.createCell(coluna);
        }
        
        cell.setCellType(CellType.STRING);
        cell.setCellValue(conteudo);
        try {
            FileOutputStream arquivo = new FileOutputStream("Exemplo\\testeA.xlsx");
            wb.write(arquivo);
            wb.close();
        } catch (IOException e) {
        }
        System.out.println("Escrito com Sucesso");
    }

    

}
