/*
 * Excel.java
 *
 * Created on 29 juni 2007, 8:06
 *
 * To change this template, choose Tools | Template Manager
 * and open the template in the editor.
 */

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 *
 * @author TammesC
 */
public class Excel {
    
    private final int MAX_ROWS = 64;
    private final String START_TEKST = "Project";
    private final String STOP_TEKST = "Totaal";
    private final String START_WERK = "tijd_in";
    private final String STOP_WERK = "tijd_uit";
    private final int FILE_READ = 0;
    private final int FILE_WRITE = 1;

    private String sheetDirectory = "";     //"C:\Documents and Settings\TammesC\Mijn documenten\Uren\"
//    private String sheetNaam = "Urenregistratie CT";   // hier het weeknummer en de extensie aan toevoegen
    private String sheetNaam = "CTS";   // hier het weeknummer en de extensie aan toevoegen
    private FileInputStream sheetfile = null;   // spreadsheet bestand
    private int regelVan = 0;       // eerste dataregel
    private int regelTm = 0;        // laatste dataregel
    private HSSFWorkbook werkboek;  // werkboek
    private HSSFSheet werkblad;     // werkblad
    
    /** Creates a new instance of Excel */
    public Excel(String xlsFile) {
        try {
            sheetfile = new FileInputStream(new File(xlsFile));
            werkboek = new HSSFWorkbook(sheetfile);
            werkblad = werkboek.getSheetAt(0);
        } catch(Exception e) {
            System.out.println(e.getMessage());
        }

    }

    public FileInputStream getSheetfile() {
        return sheetfile;
    }

    public void setSheetfile(FileInputStream sheetfile) {
        this.sheetfile = sheetfile;
    }

    /**
     * open een werkblad
     * @param weeknummer het weeknummer
     * @param werkblad naam van het werkblad of leeg (dan wordt het eerste werkblad geopend)
     */
    public void leesWerkblad(int weeknummer, String werkblad) {
        
        FileInputStream bestand=null;
        String pad = getCompleteSheetNaam(weeknummer);
        try {
            bestand = new FileInputStream(pad);
            setSheetfile(bestand);
            try {
                HSSFWorkbook werkboek = new HSSFWorkbook(bestand);
                setWerkboek(werkboek);
                if (werkblad.equals("")) {
                    setWerkblad(werkboek.getSheetAt(0)); 
                } else {
                    setWerkblad(werkboek.getSheet(werkblad)); 
                }
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(null, ex.toString(),"leesWerkblad",JOptionPane.ERROR_MESSAGE);
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, ex.toString(),"leesWerkblad",JOptionPane.ERROR_MESSAGE);
        }

        sluitWerkblad();

    }

    /**
     * open een werkblad
     * @param weeknummer het weeknummer
     */
    public void schrijfWerkblad(int weeknummer) {
        
        FileOutputStream bestand=null;
        try {
            bestand = new FileOutputStream(getCompleteSheetNaam(weeknummer));
            try {
                werkboek=getWerkboek();
                werkboek.write(bestand);
            } catch (IOException ex) {
                    JOptionPane.showMessageDialog(null, ex.toString(),"openWerkblad",JOptionPane.ERROR_MESSAGE);
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, ex.toString(), "openWerkblad", JOptionPane.ERROR_MESSAGE);
        }

        sluitWerkblad();
    }
        

    /**
     * sluit het spreadsheet bestand
     */
    public void sluitWerkblad() {
        try {
            sheetfile.close();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }
    
    public int aantalRijen() {
        HSSFSheet blad = getWerkblad();
        return blad.getLastRowNum();
    }
    
    public int getRegelVan() {
        return regelVan;
    }

    public void setRegelVan(int regelVan) {
        this.regelVan = regelVan;
    }

    public int getRegelTm() {
        return regelTm;
    }

    public void setRegelTm(int regelTm) {
        this.regelTm = regelTm;
    }

    public HSSFSheet getWerkblad() {
        return werkblad;
    }

    public void setWerkblad(HSSFSheet werkblad) {
        this.werkblad = werkblad;
    }

    public HSSFRow getRegel(int regel) {
        return werkblad.getRow(regel);
    }


    /**
     * lees de projectnamen uit het werkblad
     * @return List<String> met projectnamen
     */
    public List<String> leesProjecten() {
        
        // zoek eerste projectregel op
        int rijnum=werkblad.getFirstRowNum();
        while (rijnum<=werkblad.getLastRowNum()) {
            if (leesCel(rijnum,(short) 0).equals(START_TEKST)) {
                break;
            } else {
                rijnum++;
            }
        }
        rijnum++;
        
        // lees de projectnamen en zet ze in een lijst
        List< String > projecten = new ArrayList< String >();
        String waarde="";
        while (rijnum<=werkblad.getLastRowNum()) {
            if (leesCel(rijnum,(short) 0).equals(STOP_TEKST)) {
              break;
            } else {
                waarde=leesCel(rijnum,(short) 2);
                waarde=leesCel(rijnum,(short) 1);

                if (!waarde.equals("")) {
                    projecten.add(waarde);
                }
                rijnum++;
            }
        }

        return projecten;
    
    }

    /**
     * Lees tijd-in en tijd uit per werkdag
     * @return
     */
    public List<Werkdag> leesWerkTijden() {

        Iterator<Row> rowIterator = werkblad.iterator();
        List<Werkdag> werkdagen = new ArrayList<Werkdag>();
        while(rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(0);
            if (celWaarde(cell).equals("tijd_in")) {
                int rij = row.getRowNum();
                for (int dag = Weekdagen.MA.get(); dag <= Weekdagen.VR.get(); dag++) {

                    int in = leesIntegerCel(rij, dag);
                    int uit = leesIntegerCel(rij + 1, dag);
                    Werkdag werkdag = new Werkdag(dag, in, uit);
                    werkdagen.add(werkdag);
                }
            }
        }
        return werkdagen;
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


    public int zoekProjectregel(String project) {

        Iterator<Row> rowIterator = werkblad.iterator();
        int rij = -1;
        while(rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(1);
            if (celWaarde(cell).toLowerCase().equals(project.toLowerCase())) {
                rij = row.getRowNum();
                break;
            }
        }
        return rij;
    }


/*
        // zoek eerste projectregel op
        int rijnum=blad.getFirstRowNum();
        while (rijnum<=blad.getLastRowNum()) {
            if (leesCel(rijnum,(short) 0).equals(START_TEKST)) {
                break;
            } else {
                rijnum++;
            }
        }
        rijnum++;

        // lees de projectnamen en zet ze in een lijst
        List< String > projecten = new ArrayList< String >();
        String waarde="";
        while (rijnum<=blad.getLastRowNum()) {
            if (leesCel(rijnum,(short) 0).equals(STOP_TEKST)) {
                break;
            } else {
                waarde=leesCel(rijnum,(short) 2);
                waarde=leesCel(rijnum,(short) 1);

                if (!waarde.equals("")) {
                    projecten.add(waarde);
                }
                rijnum++;
            }
        }

        return projecten;
*/


    /**
     * schrijf de projectduur naar het werkblad
     * @param weeknummer het weeknummer
     * @param werkblad de naam van het werkklad (anders het eerste werkblad)
     */
/*
    public void schrijfProjectDuur(int weeknummer, String werkblad) {
        UrenLog uren=new UrenLog();
        Calendar nu=Calendar.getInstance();
        
        // bepaal de juiste kolom
        int kolom=nu.get(Calendar.DAY_OF_WEEK);
        if (kolom==1) {
            kolom=8;    //zondag wijkt af
        }
            
        leesWerkblad(weeknummer,"");
        HSSFSheet blad=getWerkblad();
        
        // zoek eerste projectregel op
        int rijnum=blad.getFirstRowNum();
        while (rijnum<=blad.getLastRowNum()) {
            if (leesCel(rijnum,(short) 0).equals(START_TEKST)) {
                break;
            } else {
                rijnum++;
            }
        }
        rijnum++;

        // lees de projectnamen uit de sheet, haal de projectduur erbij en schrijf deze weg
        String project="";
        int waarde=0;
        while (rijnum<=blad.getLastRowNum()) {
            if (leesCel(rijnum,(short) 0).equals(STOP_TEKST)) {
              break;
            } else {
                project=leesCel(rijnum,(short) 1);
                
                if (uren.projectLijst.containsKey(project)) {
                    waarde = (Integer) uren.projectLijst.get(project);
                    if (waarde>0) {
                        schrijfDuurInCel(rijnum,(short) kolom, (double) waarde);
                    } else {
                        schrijfDuurInCel(rijnum,(short) kolom,0);
                    }
                    
                }
                
                rijnum++;
                }
        }
        
        schrijfWerkblad(weeknummer);
        JOptionPane.showMessageDialog(null, String.format("Gegevens weggeschreven naar:\n %s",getCompleteSheetNaam(weeknummer)));
    }
*/
    /**
     * lees een waarde uit een cel en geef die als string terug
     * @param rij rijnummer
     * @param kolom kolomnummer
     * @return de gevonden waarde als string
     */
    public String leesCel(int rij, int kolom) {
          
        HSSFRow row=werkblad.getRow(rij);
        Cell cell=row.getCell(kolom);
        return celWaarde(cell);
    }

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
     * schrijf de projectduur naar een cel
     * @param rij rijnummer
     * @param kolom kolomnummer
     * @param waarde de tijd in minuten
     */
    public void schrijfDuurInCel(int rij, int kolom, double waarde) {
          
        HSSFRow row=werkblad.getRow(rij);
        HSSFCell cell=row.getCell(kolom);

        if (waarde==0) {
            cell.setCellType(HSSFCell.CELL_TYPE_BLANK);
        } else {
            switch (cell.getCellType()) {
                case HSSFCell.CELL_TYPE_NUMERIC:
                    cell.setCellValue(minutenNaarNummer(waarde));
                    break;
                case HSSFCell.CELL_TYPE_FORMULA:
                    break;
                case HSSFCell.CELL_TYPE_STRING:
                    HSSFRichTextString wrde = new HSSFRichTextString(tijdNaarTekst((int) waarde));
                    cell.setCellValue(wrde);
                    break;
                case HSSFCell.CELL_TYPE_BLANK:
                    cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                    cell.setCellValue(minutenNaarNummer(waarde));
                    break;
            }
        }
        
    }

    public String getSheetDirectory() {
        return sheetDirectory;
    }

    public void setSheetDirectory(String sheetDirectory) {
        this.sheetDirectory = sheetDirectory;
    }

    public String getSheetNaam() {
        return sheetNaam;
    }

    public void setSheetNaam(String sheetNaam) {
        this.sheetNaam = sheetNaam;
    }

    /**
     * geef de volledige naam van het sheet, incl. pad
     * @return volledige pad van de sheet
     */
    public String getCompleteSheetNaam() {
        int weeknummer=Calendar.getInstance().get(Calendar.WEEK_OF_YEAR);
        return getCompleteSheetNaam(weeknummer);
    }
    
    /**
     * geef de volledige naam van het sheet, incl. pad
     * @param weeknummer weeknummer
     * @return volledige pad van de sheet
     */
    public String getCompleteSheetNaam(int weeknummer) {
        return String.format("%s%s %d.xls",getSheetDirectory(),getSheetNaam(),weeknummer);
    }
    

    public HSSFWorkbook getWerkboek() {
        return werkboek;
    }

    public void setWerkboek(HSSFWorkbook werkboek) {
        this.werkboek = werkboek;
    }

    /**
     * omzetten van een tijdsduur in minuten naar een string (hh:mm)
     * (Een datum/tijd cel kan maximaal 23:59 groot zijn.)
     * @param tijdWaarde de numerieke celwaarde van een Excel datum/tijd cel
     * @return een string in het formaat hh:mm
     */
    public String tijdNaarTekst(int tijdWaarde) {
        int uren=0, minuten=0;

        if (tijdWaarde>0) {
            uren=(int) tijdWaarde/60;
            minuten=(int) (tijdWaarde-uren*60);
        }

        return String.format("%02d:%02d",uren,minuten);
    }

    public String tijdNaarTekst(double tijdWaarde) {
        long uren=0, minuten=0;

        if (tijdWaarde>0) {
            uren=(long) tijdWaarde/60;
            minuten =(long) ((tijdWaarde / 60 - uren) * 60);
        }

        return String.format("%02d:%02d",uren,minuten);
    }

    /**
     * omzetten van een tijdsduur in minuten naar een string (hh:mm)
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
     * omzetten van een HSSFCell.getNumericCellValue naar minuten
     * @param tijdWaarde de numerieke celwaarde van een Excel datum/tijd cel
     * @return het aantal minuten
     */
    public double nummerNaarMinuten(double tijdWaarde) {

        double waarde=0;
        if (tijdWaarde>0) {
            waarde=(tijdWaarde*24)*60;
        }

        return waarde;
    }

    /**
     * omzetten van minuten naar een HSSFCell.getNumericCellValue
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
     * totaliseer de projectduur
     * @return het totale aantal minuten
     */
    public int totaliseerDuur(String project) {

        int rij = zoekProjectregel(project);

        int totaal = 0;
        if (rij >= 0) {
            for (int dag = Weekdagen.MA.get(); dag <= Weekdagen.VR.get(); dag++) {
                String waarde = leesCel(rij, dag);
                if (waarde != "") {
                    totaal += Integer.parseInt(leesCel(rij, dag).split("\\.")[0]);
                }
            }
        }

        return totaal;

    }

    public float totaliseerDagtotaal() {

        Iterator<Row> rowIterator = werkblad.iterator();
        int rij = -1;
        while(rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(0);
            if (celWaarde(cell).toLowerCase().equals("dagtotaal")) {
                rij = row.getRowNum();
                break;
            }
        }

        float totaal = 0;
        if (rij >= 0) {
            String waarde = leesCel(rij, Weekdagen.TOTAAL.get());
            if (waarde != "") {
                totaal = Float.parseFloat(leesCel(rij, Weekdagen.TOTAAL.get()));
            }
        }

        return totaal;

    }

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
