package nl.ctammes.exceluren;

import nl.ctammes.common.Diversen;
import nl.ctammes.common.Excel;
import org.apache.commons.io.FileUtils;
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
    public static final String START_TEKST = "Project";     // projecten/tellers beginnen hierna (kolom A)
    public static final String STOP1 = "Dagtotaal";         // tellers stoppen hiervoor (kolom A)
    public static final String START1 = "Algemeen";         // tellers beginnen hier weer (kolom A)
    public static final String STOP_TEKST = "Totaal";       // projecten/tellers eindigen hiervoor (kolom A)
    public static final String START_WERK = "tijd_in";      // regel met start werktijd (kolom A)
    public static final String STOP_WERK = "tijd_uit";      // regel met stop werktijd (kolom A)
    public static final int URENPERDAG = 9;                 // aantal gewerkte uren per dag
    public static final int DAGENPERWEEK = 4;               // aantal gewerkte dagen per week
    public static final String URENMASK = "CTS\\d{2}\\.xls";  // filemask voor uren files
    public static final String URENTEMPLATE = "CTS%02d.xls";  // filemask voor uren files

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
    public float geefTaakDuur(String project) {

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
    public float geefDagtotaal() {

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
    public List<Verlofdag> geefVerlofPerDag(int weeknr, int jaar) {

        List<Verlofdag> verlof = new ArrayList<Verlofdag>();
        int rij = zoekTaakregel("verlof");
        if (rij >= 0) {
            for (int dag = Weekdagen.MA.get(); dag <= Weekdagen.VR.get(); dag++) {
                String waarde = leesCel(rij, dag);
                if (!waarde.equals("")) {
                    verlof.add(new Verlofdag(dag, getDatumUitWeekDag(weeknr, dag, jaar), Float.parseFloat(waarde)));
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
    private Integer getEersteProjectregel() {
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
        int rijnum=getEersteProjectregel();

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
        int rijnum=getEersteProjectregel();

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
     * Geef de volledige naam van het sheet van deze week, incl. pad
     * @return volledige pad van de sheet
     */
    public String getCompleteSheetNaam() {
        int weeknummer= Calendar.getInstance().get(Calendar.WEEK_OF_YEAR);
        return getCompleteSheetNaam(weeknummer);
    }

    /**
     * Geef de volledige naam van het sheet, incl. pad
     * @param weeknummer weeknummer
     * @return volledige pad van de sheet
     */
    public String getCompleteSheetNaam(int weeknummer) {
        return getSheetPath().toString();
    }

    /**
     * Geef het weeknummer uit de filenaam terug
     * @param fileName
     * @return
     */
    public int getWeeknrUitFilenaam(String fileName) {
        int result = 0;
        Matcher mat = Pattern.compile(".+(\\d{2})\\.xls", Pattern.CASE_INSENSITIVE).matcher(fileName);
        if (mat.find()) {
            result = Integer.valueOf(mat.group(1));
        }
        return result;
    }

    /**
     * Geeft de datum aan de hand van weeknummer, weekdag en jaar
     * @param weeknr
     * @param weekdag vb. Calendar.FRIDAY
     * @param jaar
     * @return
     */
    public static String getDatumUitWeekDag(int weeknr, int weekdag, int jaar) {
        SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
        Calendar cal = Calendar.getInstance();
//        cal.setFirstDayOfWeek(Calendar.MONDAY);
        cal.set(Calendar.YEAR, jaar);
        cal.set(Calendar.WEEK_OF_YEAR, weeknr);
        cal.set(Calendar.DAY_OF_WEEK, weekdag);
        return sdf.format(cal.getTime());
    }

    /**
     * Bepaal het jaar uit de directorynaam die eindigt op het jaarnnummer, anders het huidige jaar
     * @param dir
     * @return
     */
    public static int getJaarUitDirnaam(String dir) {
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
     * @throws Exception
     */
    public static void maakNieuwBestand(String xlsDir, int weeknrOud, int weeknrNieuw) throws Exception {

        String fileOud = xlsDir + File.separatorChar + String.format(URENTEMPLATE, weeknrOud);
        String fileNieuw = xlsDir + File.separatorChar + String.format(URENTEMPLATE, weeknrNieuw);

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
                resp = JOptionPane.showConfirmDialog(null, "Huidige bestand is ouder dan te maken bestand (" + dateNieuw + "). Doorgaan?", "Bevestig keuze", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE);
                if (resp == JOptionPane.YES_OPTION) {
                    resp = JOptionPane.showConfirmDialog(null, fileNieuw + " bestaat al. Overschrijven?", "Bevestig keuze", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE);
                    if (resp == JOptionPane.YES_OPTION) {
                        nieuw.delete();
                    }
                }
            }
        }

        // kopieer bestand naar nieuwe naam
        if (resp == JOptionPane.YES_OPTION) {
            FileUtils.copyFile(oud, nieuw);

            maakBestandLeeg(fileNieuw, weeknrNieuw);
        }
        throw new Exception("Bestand niet overschrijven");
    }

    /**
     * Maakt een urenbestand leeg en wijzig het weeknummer in de tekst
     * @param xlsPath   volledige naam van xls
     */
    private static void maakBestandLeeg(String xlsPath, int weeknr) {
        ExcelUren nieuw = new ExcelUren(xlsPath);
        nieuw.schrijfCel(2, 1, "Week: " + weeknr);
        nieuw.wisUren(nieuw);
        //TODO in/uit uren resetten?
        nieuw.sluitWerkboek();
    }

    /**
     * Wis de regels met gewerkte uren en sla het bestand op
     * @param nieuw
     */
    private void wisUren(ExcelUren nieuw) {
        for (int rij = nieuw.zoekProjectregel(START_TEKST) - 1; rij < nieuw.zoekProjectregel(STOP1); rij++) {
            nieuw.wisCellen(rij, Weekdagen.MA.get(), 5);
        }

        for (int rij = nieuw.zoekProjectregel(START1); rij < nieuw.zoekProjectregel(STOP_TEKST) - 1; rij++) {
            nieuw.wisCellen(rij, Weekdagen.MA.get(), 5);
        }

        nieuw.schrijfWerkboek();

    }


    /**
     * Stelt een nieuwe filenaam samen uit de huidige met daarin het huidige weeknummer
     * Het weeknummer bestaat uit twee posities direct voor de punt.
     * @return
     */
    public static String maakNieuweFilenaam(String filenaam) {
        String result = "";
        Pattern pat = Pattern.compile("(.+)\\d{2}(\\.xls)", Pattern.CASE_INSENSITIVE);
        Matcher mat = pat.matcher(filenaam);
        while (mat.find()) {
            result = String.format("%s%02d%s", mat.group(1), Diversen.getWeeknummer(), mat.group(2));
        }
        return result;

    }


    }
