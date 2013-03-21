import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FilenameFilter;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Logger;
import java.util.regex.Pattern;

/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 20-3-13
 * Time: 12:13
 * To change this template use File | Settings | File Templates.
 */
public class ExcelUren extends Excel {

    private final int MAX_ROWS = 64;
    private final String START_TEKST = "Project";
    private final String STOP_TEKST = "Totaal";
    private final String START_WERK = "tijd_in";
    private final String STOP_WERK = "tijd_uit";

    // initialiseer logger
    public static Logger log = Logger.getLogger(ExcelUren.class.getName());

    public ExcelUren(String logDir, String logNaam) {
        super(logDir, logNaam);
    }



//        String logDir = ".";
//        String logNaam = "ExcelUren.log";
//        try {
//            MijnLog mijnlog = new MijnLog(logDir, logNaam, true);
//            log = mijnlog.getLog();
//            log.setLevel(Level.INFO);
//        } catch (Exception e) {
//            System.out.println(e.getMessage());
//        }


    /**
     * Zoek de rij met de opgegeven projectnaam
     * @param project
     * @return rijnummer
     */
    public int zoekProjectregel(String project) {

        Iterator<Row> rowIterator = getWerkblad().iterator();
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

    public float geefProjectDuur(String project) {

        int rij = zoekProjectregel(project);

        float totaal = 0;
        if (rij >= 0) {
            String waarde = leesCel(rij, Weekdagen.TOTAAL.get());
            if (!waarde.equals("")) {
                totaal = Float.parseFloat(leesCel(rij, Weekdagen.TOTAAL.get()));
            }
        }

        return totaal;

    }

    /**
     * Geef dagtotaal van dit werkblad
     * @return
     */
    public float geefDagtotaal() {

        Iterator<Row> rowIterator = getWerkblad().iterator();
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
            if (!waarde.equals("")) {
                totaal = Float.parseFloat(leesCel(rij, Weekdagen.TOTAAL.get()));
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
     * Lees alle xls-namen (CTSnn.xls) uit de opgegeven direcctory
     * @param dirXls
     * @return
     */
    public String[] leesXlsNamen(String dirXls) {
        File map = new File(dirXls);
        String[] files = map.list(new FilenameFilter() {
            @Override
            public boolean accept(File map, String fileName) {
                return Pattern.matches("cts\\d+\\.xls", fileName.toLowerCase());
            }
        });
        return files;
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
