import nl.ctammes.common.Excel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.Iterator;
import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 20-3-13
 * Time: 12:13
 * To change this template use File | Settings | File Templates.
 */
public class ExcelUren extends Excel {

    public static final int MAX_ROWS = 64;
    public static final String START_TEKST = "Project";     // projecten beginnen hierna (kolom A)
    public static final String STOP_TEKST = "Totaal";       // projecten eindigen hiervoor (kolom A)
    public static final String START_WERK = "tijd_in";      // regel met start werktijd (kolom A)
    public static final String STOP_WERK = "tijd_uit";      // regel met stop werktijd (kolom A)
    public static final int URENPERDAG = 9;                 // aantal gewerkte uren per dag
    public static final String URENMASK = "cts\\d{2}\\.xls";  // filemask voor uren files

    public ExcelUren(String logDir, String logNaam) {
        super(logDir, logNaam);
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

                    System.out.println(leesIntegerCel(rij, dag));
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
     * lees de projectnamen uit het werkblad
     * @return List<String> met projectnamen
     */
    public List<String> leesProjecten() {

        // zoek eerste projectregel op
        int rijnum=getWerkblad().getFirstRowNum();
        while (rijnum<=getWerkblad().getLastRowNum()) {
            if (leesCel(rijnum, (short) 0).equals(START_TEKST)) {
                break;
            } else {
                rijnum++;
            }
        }
        rijnum++;

        // lees de projectnamen en zet ze in een lijst
        List< String > projecten = new ArrayList< String >();
        String waarde="";
        while (rijnum<=getWerkblad().getLastRowNum()) {
            if (leesCel(rijnum, (short) 0).equals(STOP_TEKST)) {
                break;
            } else {
                waarde=leesCel(rijnum, (short) 2);
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

}
