/*
 * Excel.java
 *
 * Created on 29 juni 2007, 8:06
 *
 * To change this template, choose Tools | Template Manager
 * and open the template in the editor.
 */

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;

/**
 *
 * @author TammesC
 */
public class Excel {
    
    private String sheetDirectory = "";     //"C:\Documents and Settings\TammesC\Mijn documenten\Uren\"
    private String sheetFile = null;   // xls-bestand
    private FileInputStream sheetPath = null;   // xls-bestand incl. path
    private int regelVan = 0;       // eerste dataregel
    private int regelTm = 0;        // laatste dataregel
    private HSSFWorkbook werkboek;  // werkboek
    private HSSFSheet werkblad;     // werkblad

    public Excel() {
    }

    /** Creates a new instance of Excel */
    public Excel(String xlsDir, String xlsFile) {
        try {
            sheetDirectory = xlsDir;
            sheetFile = xlsFile;
            sheetPath = new FileInputStream(new File(xlsDir + "/" + xlsFile));
            werkboek = new HSSFWorkbook(sheetPath);
            werkblad = werkboek.getSheetAt(0);
        } catch(Exception e) {
            System.out.println(e.getMessage());
        }

    }

    public FileInputStream getSheetPath() {
        return sheetPath;
    }

    public HSSFSheet getWerkblad() {
        return werkblad;
    }

    public HSSFRow getRegel(int regel) {
        return werkblad.getRow(regel);
    }

    public String getSheetDirectory() {
        return sheetDirectory;
    }

    public String getSheetFile() {
        return sheetFile;
    }

    public void setSheetDirectory(String sheetDirectory) {
        this.sheetDirectory = sheetDirectory;
    }

    public HSSFWorkbook getWerkboek() {
        return werkboek;
    }

    public void setWerkboek(HSSFWorkbook werkboek) {
        this.werkboek = werkboek;
    }

    /**
     * Sluit het spreadsheet bestand
     */
    public void sluitWerkblad() {
        try {
            sheetPath.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    /**
     * Geef aantal rijen in werkblad
     * @return
     */
    public int aantalRijen() {
        return werkblad.getLastRowNum();
    }

    /**
     * Lees numerieke waarde en geef integer terug
     * Gebruikt bij uitlezen van tijd
     * @param rij
     * @param kolom
     * @return
     */
    public int leesIntegerCel(int rij, int kolom) {
        String inhoud = leesCel(rij, kolom);
        int waarde = 0;
        if (inhoud != null) {
            waarde = Integer.parseInt(leesCel(rij, kolom).split("\\.")[0]);
        }
        return waarde;
    }

    /**
     * Lees een waarde uit een cel en geef die als string terug
     * @param rij rijnummer
     * @param kolom kolomnummer
     * @return de gevonden waarde als string
     */
    public String leesCel(int rij, int kolom) {
          
        HSSFRow row=werkblad.getRow(rij);
        Cell cell=row.getCell(kolom);
        return celWaarde(cell);
    }

    /**
     * Geef de waarde van een cel terug
     * @param cell
     * @return string
     */
    public String celWaarde(Cell cell) {

        String waarde="";
        if (cell != null) {
        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_NUMERIC:
                waarde = Double.toString(nummerNaarMinuten(cell.getNumericCellValue()));
                break;
            case HSSFCell.CELL_TYPE_FORMULA:
                waarde = Double.toString(cell.getNumericCellValue());
                break;
            case HSSFCell.CELL_TYPE_STRING:
                waarde = cell.getRichStringCellValue().toString();
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                break;
            default:
                break;
        }
        }
        return waarde;

    }

    /**
     * Omzetten van een tijdsduur in minuten naar een string (hh:mm)
     * (Een datum/tijd cel kan maximaal 23:59 groot zijn.)
     * @param tijdWaarde de numerieke celwaarde van een Excel datum/tijd cel
     * @return een string in het formaat hh:mm
     */
    public String tijdNaarTekst(int tijdWaarde) {
        int uren=0, minuten=0;

        if (tijdWaarde > 0) {
            uren=(int) tijdWaarde / 60;
            minuten=(int) (tijdWaarde - uren * 60);
        }

        return String.format("%02d:%02d",uren,minuten);
    }

    /**
     * Omzetten van een tijdsduur in minuten naar een string (hh:mm)
     * (Een datum/tijd cel kan maximaal 23:59 groot zijn.)
     * @param tijdWaarde
     * @return
     */
    public String tijdNaarTekst(double tijdWaarde) {
        long uren=0, minuten=0;

        if (tijdWaarde > 0) {
            uren=(long) tijdWaarde / 60;
            minuten =(long) ((tijdWaarde / 60 - uren) * 60);
        }

        return String.format("%02d:%02d",uren,minuten);
    }

    /**
     * Omzetten van een tijdsduur in minuten naar een string (hh:mm)
     * (Een datum/tijd cel kan maximaal 23:59 groot zijn.)
     * @param tijdTekst de tijd in tekst (hh:mm)
     * @return de tijd in minuten
     */
    public int tekstNaarTijd(String tijdTekst) {
        int uren=0, minuten=0;

        if (!tijdTekst.equals("")) {
            String[] tijd=tijdTekst.split(":");

            uren=Integer.valueOf(tijd[0]);
            minuten=Integer.valueOf(tijd[1]);
            return (uren*60)+minuten;
        } else {
            return 0;
        }

    }

    /**
     * Omzetten van een HSSFCell.getNumericCellValue naar minuten
     * @param tijdWaarde de numerieke celwaarde van een Excel datum/tijd cel
     * @return het aantal minuten
     */
    public int nummerNaarMinuten(double tijdWaarde) {

        int waarde=0;
        if (tijdWaarde>0) {
            waarde=(int) Math.round((tijdWaarde*24)*60);
        }

        return waarde;
    }

    /**
     * Omzetten van minuten naar een HSSFCell.getNumericCellValue
     * @param minuten de tijd in minuten
     * @return de getalwaarde (deel van het etmaal)
     */
    public double minutenNaarNummer(double minuten) {

        double waarde=0;
        if (minuten>0) {
            waarde=(minuten/60)/24;
        }

        return waarde;
    }


    /**
     * Geef begin en einddatum van de opgegeven week in het opgegeven jaar
     * @param weeknr
     * @param jaar
     * @return
     */
    public String[] geefWeekDatums(int weeknr, int jaar) {
        String[] dagen = new String[2];
        SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
        Calendar cal = Calendar.getInstance();
        cal.clear();
        cal.set(Calendar.WEEK_OF_YEAR, weeknr);
        cal.set(Calendar.YEAR, jaar);
        cal.set(Calendar.DAY_OF_WEEK, Calendar.MONDAY);
        dagen[0] = sdf.format(cal.getTime());
        cal.set(Calendar.DATE, cal.get(Calendar.DATE) + 6 );
        dagen[1] = sdf.format(cal.getTime());
        return dagen;

    }

}
