package nl.ctammes.exceluren;

import nl.ctammes.common.Diversen;
import nl.ctammes.common.Excel;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import javax.swing.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 20-3-13
 * Time: 12:13
 * To change this template use File | Settings | File Templates.
 */
public class ExcelUren extends Excel {

    public static final int MAX_ROWS = 64;
    public static final String START_TEKST = "Project";      // projecten/tellers beginnen hierna (kolom A)
    public static final String STOP1 = "Dagtotaal";          // tellers stoppen hiervoor (kolom A)
    public static final String START1 = "Algemeen";          // tellers beginnen hier weer (kolom A)
    public static final String STOP_TEKST = "Totaal";        // projecten/tellers eindigen hiervoor (kolom A)
    public static final String START_WERK = "tijd_in";       // regel met start werktijd (kolom A)
    public static final String STOP_WERK = "tijd_uit";       // regel met stop werktijd (kolom A)
    public static final int URENPERDAG = 9;                  // aantal gewerkte uren per dag
    public static final int DAGENPERWEEK = 4;                // aantal gewerkte dagen per week
    public static final String XLSMASK = "(.+)(\\d{2})(\\.xls)";  // filemask voor uren files
    public static final String XLSTEMPLATE = "CTS%02d.xls";  // filemask voor uren files

    public ExcelUren(String xlsPath) {
        super(Diversen.splitsPad(xlsPath)[0], Diversen.splitsPad(xlsPath)[1]);
    }

    public ExcelUren(String xlsDir, String xlsNaam) {
        super(xlsDir, xlsNaam);
    }

    /**
     * Zoek de rij met de opgegeven projectnaam
     * @param project
     * @return
     */
    public int zoekProjectregel(String project) {

        Iterator<Row> rowIterator = getWerkblad().iterator();
        int rij = -1;
        while(rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(0);
            if (celWaarde(cell).toLowerCase().equals(project.toLowerCase())) {
                rij = row.getRowNum();
                break;
            }
        }
        return rij;
    }

    /**
     * Zoek de rij met de tijd_in
     * @return rijnummer
     */
    public int zoekTijdinRegel() {

        Iterator<Row> rowIterator = getWerkblad().iterator();
        int rij = -1;
        while(rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(0);
            if (celWaarde(cell).toLowerCase().equals("tijd_in")) {
                rij = row.getRowNum();
                break;
            }
        }
        return rij;
    }

    /**
     * Zoek de rij met de opgegeven taaknaam
     * @param taak
     * @return rijnummer
     */
    public int zoekTaakregel(String taak) {

        Iterator<Row> rowIterator = getWerkblad().iterator();
        int rij = -1;
        while(rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(1);
            if (celWaarde(cell).toLowerCase().equals(taak.toLowerCase())) {
                rij = row.getRowNum();
                break;
            }
        }
        return rij;
    }

    /**
     * Geef de weekduur van een project in minuten
     * (0,375 is 9 uur)
     * @param project
     * @return
     */
    public float taakDuur(String project) {

        int rij = zoekTaakregel(project);

        float totaal = 0;
        if (rij >= 0) {
            String waarde = leesCel(rij, Weekdagen.TOTAAL.get());
            if (!waarde.equals("")) {
                totaal = Float.parseFloat(waarde) * 24 * 60;
            }
        }

        return totaal;

    }

    /**
     * Geef gewerkte uren van dit werkblad
     * @return
     */
    public float dagTotaal() {

        int rij = zoekProjectregel("dagtotaal");

        float totaal = 0;
        if (rij >= 0) {
            String waarde = leesCel(rij, Weekdagen.TOTAAL.get());
            if (!waarde.equals("")) {
                totaal = Float.parseFloat(waarde) * 24;
            }
        }

        return totaal;
    }

    /**
     * Geef overzicht van datum en verlofuren
     * @param weeknr
     * @param jaar
     * @return
     */
    public List<Verlofdag> verlofPerDag(int weeknr, int jaar) {

        List<Verlofdag> verlof = new ArrayList<Verlofdag>();
        int rij = zoekTaakregel("verlof");
        if (rij >= 0) {
            for (int dag = Weekdagen.MA.get(); dag <= Weekdagen.VR.get(); dag++) {
                String waarde = leesCel(rij, dag);
                if (!waarde.equals("")) {
                    verlof.add(new Verlofdag(dag, Diversen.datumUitWeekDag(weeknr, dag, jaar), Float.parseFloat(waarde)));
                }
            }

        }

        return verlof;
    }

    /**
     * Lees tijd-in en tijd uit per werkdag
     * @return
     */
    public List<Werkdag> leesWerkTijden() {

        Iterator<Row> rowIterator = getWerkblad().iterator();
        List<Werkdag> werkdagen = new ArrayList<Werkdag>();
        while(rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(0);
            if (celWaarde(cell).equals(START_WERK)) {
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
     * Geef het regelnummer van het eerste project
     * @return
     */
    private Integer eersteProjectregel() {
        // zoek eerste projectregel op
        int rijnum=getWerkblad().getFirstRowNum();
        while (rijnum<=getWerkblad().getLastRowNum()) {
            if (leesCel(rijnum, (short) 0).equals(START_TEKST)) {
                break;
            } else {
                rijnum++;
            }
        }
        return ++rijnum;
    }

    /**
     * lees de projectnamen uit het werkblad
     * @return List<String> met projectnamen
     */
    public List<String> leesProjecten() {

        // zoek eerste projectregel op
        int rijnum= eersteProjectregel();

        // lees de projectnamen en zet ze in een lijst
        List< String > projecten = new ArrayList< String >();
        String waarde="";
        while (rijnum<=getWerkblad().getLastRowNum()) {
            if (leesCel(rijnum, (short) 0).equals(STOP_TEKST)) {
                break;
            } else {
                waarde=leesCel(rijnum, (short) 1);

                if (!waarde.equals("")) {
                    projecten.add(waarde);
                }
                rijnum++;
            }
        }

        return projecten;

    }

    public List<String> leesPOProjecten() {

        // zoek eerste projectregel op
        int rijnum= eersteProjectregel();

        // lees de projectnamen en zet ze in een lijst
        List< String > projecten = new ArrayList<String>();
        String waarde="";
        while (rijnum<=getWerkblad().getLastRowNum()) {
            if (leesCel(rijnum, (short) 0).equals(STOP_TEKST)) {
                break;
            } else {
                if (leesCel(rijnum, (short) 0).equals("PO")) {
                    waarde = leesCel(rijnum, (short) 1);

                    if (!waarde.equals("")) {
                        projecten.add(waarde);
                    }
                }
                rijnum++;
            }
        }

        return projecten;

    }

    /**
     * Geeft een lijst van alle projecten waar een totaal ingevuld is
     * Let op: voor de weekdagen geldt een andere opmaak dan voor het weektotaal!
     * @param kol dag van de week
     * @return
     */
    public Map projectenMetTotaal(int kol) {
        String naam = "";
        String tijd = "";
        long uren, minuten = 0;
        Map<String, String> results = new HashMap<String, String>();

        // zoek eerste projectregel op
        int rijnum= eersteProjectregel();

        Iterator<Row> rowIterator = getWerkblad().iterator();
        while (rijnum<=getWerkblad().getLastRowNum()) {
            if (leesCel(rijnum, 0).equals(STOP_TEKST)) {
                break;
            } else {
                naam = leesCel(rijnum, 1);
                tijd = leesCel(rijnum, kol);
                if (!naam.equals("")  && !tijd.equals("") && (Double.parseDouble(tijd) > 0.0)) {
                    if (kol == Weekdagen.TOTAAL.get()) {
                        uren = (long) (Double.parseDouble(tijd) * 24);
                        minuten = (long) ((Double.parseDouble(tijd) * 24 - uren) * 60);
                    } else {
                        uren = (long) (Double.parseDouble(tijd) / 60);
                        minuten = (long) ((Double.parseDouble(tijd) / 60 - uren) * 60);
                    }
                    results.put(naam, String.format("%02d:%02d", uren, minuten));
                }
                rijnum++;
            }
        }
        return results;
    }

    /**
     * Geeft aan of er deze week tijden zijn ingevuld
     * @param taak
     * @return
     */
    public boolean zijnTaakDagenGevuld(String taak) {
        // zoek projectregel op
        int rijnum = zoekTaakregel(taak);

        boolean result = false;
        if (rijnum >= 0) {
            for (int kol = Weekdagen.MA.get(); kol <= Weekdagen.ZO.get(); kol++) {
                String tijd = leesCel(rijnum, kol);
                if (!tijd.equals("") && (Double.parseDouble(tijd) > 0.0)) {
                    result = true;
                    break;
                }
            }
        }
        return result;
    }

    /**
     * Taak toevoegen in het werkblad
     * @param taak
     */
    public void taakToevoegen(String taak) {
        int waar = bepaalRijNieuweTaak(taak);
        if (waar > 0) {
            invoegenRijOnder(waar, taak);
        } else {
            invoegenRijBoven(-waar, taak);
        }
    }

    /**
     * Rij invoegen in werkblad onder het aangegeven rijnummer
     * @param rij
     */
    public void invoegenRijOnder(int rij, String taak) {
        if ( rij >= 0 ) {
            // schuif alles 1 rij op en voeg nieuwe rij in
            getWerkblad().shiftRows(rij + 1, getWerkblad().getLastRowNum(), 1);
            HSSFRow oude_rij = getWerkblad().getRow(rij);
            HSSFRow nieuwe_rij = getWerkblad().createRow(rij + 1);

            kopierenRij(oude_rij, nieuwe_rij, taak);
            if (log != null) {
                log.info(String.format("Taak '%s' toevoegen onder rij %d", taak, rij));
            }

        }
    }

    /**
     * Rij invoegen in werkblad boven het aangegeven rijnummer
     * @param rij
     */
    public void invoegenRijBoven(int rij, String taak) {
        if ( rij >= 0 ) {
            // schuif alles 1 rij op en voeg nieuwe rij in
            getWerkblad().shiftRows(rij, getWerkblad().getLastRowNum(), 1);
            HSSFRow oude_rij = getWerkblad().getRow(rij + 1);
            HSSFRow nieuwe_rij = getWerkblad().createRow(rij);

            kopierenRij(oude_rij, nieuwe_rij, taak);
            if (log != null) {
                log.info(String.format("Taak '%s' toevoegen boven rij %d", taak, rij));
            }
        }
    }

    /**
     * Kopieer gegevens van de oude rij naar de nieuwe rij
     * @param oude_rij HSSFRow
     * @param nieuwe_rij HSSFRow
     */
    private void kopierenRij(HSSFRow oude_rij, HSSFRow nieuwe_rij, String taak) {
        // kopieer celeigenschappen van oude naar nieuwe rij
        for (int kolom=0; kolom<oude_rij.getLastCellNum(); kolom++) {
            HSSFCell oude_cel = oude_rij.getCell(kolom);
            HSSFCell nieuwe_cel = nieuwe_rij.createCell(kolom);
            nieuwe_cel.setCellType(oude_cel.getCellType());
            if (kolom >= Weekdagen.MA.get() && kolom <= Weekdagen.ZO.get()) {
                nieuwe_cel.setCellType(HSSFCell.CELL_TYPE_BLANK);
            }
            nieuwe_cel.setCellStyle(oude_cel.getCellStyle());
        }
        nieuwe_rij.getCell(0).setCellValue("PO");
        nieuwe_rij.getCell(1).setCellValue(taak);
        nieuwe_rij.getCell(Weekdagen.TOTAAL.get()).setCellFormula(String.format("I%1$d+H%1$d+G%1$d+E%1$d+F%1$d+D%1$d+C%1$d", nieuwe_rij.getRowNum() + 1));

        schrijfWerkboek();

    }

    /**
     * Taak uit werkblad verwijderen
     * @param taak
     */
    public void taakVerwijderen(String taak) {
        int waar = zoekTaakregel(taak);
        if (waar >= 0) {
            wisRij(waar);
        }
    }

    /**
     * Bepaal waar een nieuwe taak in het werkblad moet worden ingevoegd
     * Een taak wordt alfabetisch toegevoegd aan de PO projecten
     * @param naam
     * @return rij waaronder de taak moet worden toegevoegd (<0: waarboven)
     */
    public int bepaalRijNieuweTaak(String naam) {
        List<String> taken = leesPOProjecten();
        taken.add(naam);
        Collections.sort(taken);

        int pos = taken.indexOf(naam);
        if (pos == 0) {
            return -zoekTaakregel(taken.get(pos + 1));
        } else {
            return zoekTaakregel(taken.get(pos - 1));
        }
    }

    /**
     * Geef de naam van het werkboek van deze week
     * @return naam van het werkboek
     */
    public static String sheetNaam() {
        int weeknummer= Calendar.getInstance().get(Calendar.WEEK_OF_YEAR);
        return sheetNaam(weeknummer);
    }

    /**
     * Geef de naam van het werkboek
     * @param weeknummer weeknummer
     * @return naam van het werkboek
     */
    public static String sheetNaam(int weeknummer) {
        return String.format(XLSTEMPLATE, weeknummer);
    }

    /**
     * Geef het weeknummer uit de filenaam terug
     * @param fileName
     * @return
     */
    public int weeknrUitFilenaam(String fileName) {
        int result = 0;
        Matcher mat = Pattern.compile(XLSMASK, Pattern.CASE_INSENSITIVE).matcher(fileName);
        if (mat.find()) {
            result = Integer.valueOf(mat.group(2));
        }
        return result;
    }

    /**
     * Bepaal de kolom ahv. de weekdag
     * Alleen zondag wijkt af
     * @param dagnr
     * @return
     */
    public static int dagKolom(int dagnr) {
        if (dagnr == 1) {
            return Weekdagen.ZO.get();
        } else {
            return dagnr;
        }
    }

    /**
     * Bepaal de kolom ahv. de weekdag voor een datum
     * @param datum
     * @return
     */
    public static int dagKolom(String datum) {
        int dagnr = Diversen.weekdagNummer(datum);
        return dagKolom(dagnr);
    }

    /**
     * Bepaal het jaar uit de directorynaam die eindigt op het jaarnnummer, anders het huidige jaar
     * @param dir
     * @return
     */
    public static int jaarUitDirnaam(String dir) {
        Calendar cal = Calendar.getInstance();
        int result = cal.get(Calendar.YEAR);    // default: dit jaar
        Pattern pat = Pattern.compile(".*(\\d{4})$");
        Matcher mat = pat.matcher(dir);
        while (mat.find()) {
            String a = mat.group(1);
            result = Integer.parseInt(mat.group(1));
        }
        return result;
    }

    /**
     * Maak een nieuw urenbestand uit een oud (huidige) bestand en zet de tijden op nul
     * @param xlsDir
     * @param weeknrOud
     * @param weeknrNieuw
     * @param dagIn
     * @param dagUit
     * @throws Exception
     */
    public static void maakNieuwBestand(String xlsDir, int weeknrOud, int weeknrNieuw, String dagIn, String dagUit) throws Exception {

        String fileOud = xlsDir + File.separatorChar + String.format(XLSTEMPLATE, weeknrOud);
        String fileNieuw = xlsDir + File.separatorChar + String.format(XLSTEMPLATE, weeknrNieuw);

        File oud = new File(fileOud);
        if (!oud.exists()) {
            JOptionPane.showMessageDialog(null, String.format("Bestand %s niet gevonden!", fileOud),"Waarschuwing",JOptionPane.WARNING_MESSAGE);
            throw new FileNotFoundException();
        }

        File nieuw = new File(fileNieuw);
        int resp = JOptionPane.YES_OPTION;
        if (nieuw.exists()) {
            // bepaal datum oude en nieuwe bestand
            Date fileOudDate = new Date(new File(fileOud).lastModified());
            Date fileNieuwDate = new Date(new File(fileNieuw).lastModified());

            SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy hh:mm:ss");
            String dateOud = sdf.format(fileOudDate);
            String dateNieuw = sdf.format(fileNieuwDate);
            // date/filedate: 07-03-2012 06:14:34 / 07-03-2012 06:06:54

            // vergelijk filedatum - huidige bestand mag niet ouder zijn dan nieuw te maken bestand
            if (fileOudDate.before(fileNieuwDate)) {
                resp = JOptionPane.showOptionDialog(null, "Huidige bestand is ouder dan te maken bestand (" + dateNieuw + "). Doorgaan?", "Bevestig keuze", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE,null,new String[] {"Ja","Nee"}, "Nee");
                if (resp == JOptionPane.YES_OPTION) {
                    resp = JOptionPane.showOptionDialog(null, fileNieuw + " bestaat al. Overschrijven?", "Bevestig keuze", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE,null,new String[] {"Ja","Nee"}, "Nee");
                    if (resp == JOptionPane.YES_OPTION) {
                        nieuw.delete();
                    }
                }
            }
        }

        // kopieer bestand naar nieuwe naam
        if (resp == JOptionPane.YES_OPTION) {
            FileUtils.copyFile(oud, nieuw);

            maakBestandLeeg(fileNieuw, weeknrNieuw, dagIn, dagUit);
        } else {
            throw new Exception("Bestand niet overschrijven");
        }
    }

    /**
     * Is vandaag een werkdag?
     * @param werkdagen
     * @return kolomnummer van de dag of 0
     */
    public static int isVandaagWerkdag(String werkdagen) {
        Calendar nu = Calendar.getInstance();

        int result = 0;
        switch (nu.DAY_OF_WEEK) {
            case Calendar.MONDAY:
                if (werkdagen.contains("ma")) {
                    result = nu.DAY_OF_WEEK;
                }
                break;
            case Calendar.TUESDAY:
                if (werkdagen.contains("di")) {
                    result = nu.DAY_OF_WEEK;
                }
                break;
            case Calendar.WEDNESDAY:
                if (werkdagen.contains("wo")) {
                    result = nu.DAY_OF_WEEK;
                }
                break;
            case Calendar.THURSDAY:
                if (werkdagen.contains("do")) {
                    result = nu.DAY_OF_WEEK;
                }
                break;
            case Calendar.FRIDAY:
                if (werkdagen.contains("vr")) {
                    result = nu.DAY_OF_WEEK;
                }
                break;
            case Calendar.SATURDAY:
                if (werkdagen.contains("za")) {
                    result = nu.DAY_OF_WEEK ;
                }
                break;
            case Calendar.SUNDAY:
                if (werkdagen.contains("zo")) {
                    result = 8;
                }
                break;
        }
        return result;
    }

    /**
     * Ligt het huidige tijdstip binnen de werktijd?
     * @param werktijdVan
     * @param werktijdTot
     * @return
     */
    public static boolean isNuWerktijd(String werkdagen, String werktijdVan, String werktijdTot) {
        Calendar nu = Calendar.getInstance();
        Calendar van = Calendar.getInstance();
        Calendar tot = Calendar.getInstance();
        String tijd[];
        boolean result = true;

        if (isVandaagWerkdag(werkdagen) > 0) {
            tijd = splitsTijd(werktijdVan);
            van.set(Calendar.HOUR_OF_DAY, Integer.parseInt(tijd[0]));
            van.set(Calendar.MINUTE, Integer.parseInt(tijd[1]));

            tijd = splitsTijd(werktijdTot);
            tot.set(Calendar.HOUR_OF_DAY, Integer.parseInt(tijd[0]));
            tot.set(Calendar.MINUTE, Integer.parseInt(tijd[1]));

            if (nu.before(van)) {
                result = false;
            }
            ;
            if (nu.after(tot)) {
                result = false;
            }
            ;
        } else {
            result = false;
        }
        return result;

    }


    /**
     * Maakt een nieuw leeg urenbestand op basis van het huidige
     * Het bestand wordt niet geopend.
     * @param xlsPath   volledige naam van xls
     */
    private static void maakBestandLeeg(String xlsPath, int weeknr, String dagIn, String dagUit) {
        ExcelUren nieuw = new ExcelUren(xlsPath);
        nieuw.schrijfCel(2, 1, "Week: " + weeknr);
        nieuw.wisUren(nieuw);
        nieuw.resetInUitTijden(nieuw, dagIn, dagUit);
        nieuw.sluitWerkboek();
    }

    /**
     * Wis de regels met gewerkte uren en sla het bestand op
     * @param nieuw
     */
    private void wisUren(ExcelUren nieuw) {
        for (int rij = nieuw.zoekProjectregel(START_TEKST) + 1; rij < nieuw.zoekProjectregel(STOP1); rij++) {
            nieuw.wisCellen(rij, Weekdagen.MA.get(), 7);
        }

        for (int rij = nieuw.zoekProjectregel(START1); rij < nieuw.zoekProjectregel(STOP_TEKST); rij++) {
            nieuw.wisCellen(rij, Weekdagen.MA.get(), 7);
        }

        nieuw.schrijfWerkboek();

    }

    /**
     * Reset de in/uit tijden
     * @param nieuw
     */
    private void resetInUitTijden(ExcelUren nieuw, String dagIn, String dagUit) {

        nieuw.schrijfTijdCellen(nieuw.zoekProjectregel(START_WERK), Weekdagen.MA.get(), 5, tekstNaarTijd(dagIn));
        nieuw.wisCellen(nieuw.zoekProjectregel(START_WERK), Weekdagen.WO.get(), 1);

        nieuw.schrijfTijdCellen(nieuw.zoekProjectregel(STOP_WERK), Weekdagen.MA.get(), 5, tekstNaarTijd(dagUit));
        nieuw.wisCellen(nieuw.zoekProjectregel(STOP_WERK), Weekdagen.WO.get(), 1);

        nieuw.schrijfWerkboek();

    }


    /**
     * Stelt een nieuwe filenaam samen uit de huidige met daarin het huidige weeknummer
     * Het weeknummer bestaat uit twee posities direct voor de punt.
     * @return
     */
    public static String maakNieuweFilenaam(String filenaam) {
        String result = "";
        Pattern pat = Pattern.compile(XLSMASK, Pattern.CASE_INSENSITIVE);
        Matcher mat = pat.matcher(filenaam);
        while (mat.find()) {
            result = String.format("%s%02d%s", mat.group(1), Diversen.weekNummer(), mat.group(3));
        }
        return result;

    }

    /**
     * bestaat het Excel bestand?
     * @return true/false
     */
    public static boolean bestaatWerkboek(String xlsDir, int weeknummer) {
        String pad = String.format("%s%s%s",xlsDir, File.separatorChar , String.format(XLSTEMPLATE, weeknummer));
        return Diversen.bestaatPad(pad);
    }



}
