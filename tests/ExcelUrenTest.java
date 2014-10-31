import junit.framework.TestCase;
import nl.ctammes.exceluren.ExcelUren;
import nl.ctammes.common.Diversen;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import nl.ctammes.exceluren.*;
import org.junit.Test;

import java.io.File;
import java.text.DateFormat;
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
public class ExcelUrenTest extends TestCase {
    static String dirXls = "/home/chris/IdeaProjects2/uren2013";

    static ExcelUren uren;

    @Override
    public void setUp() throws Exception {
        uren = new ExcelUren(dirXls, "CTS47.xls");
    }

    @Override
    public void tearDown() throws Exception {
        uren.sluitWerkboek();
    }

    @Test
    public void testUren() throws Exception {
        try {
            //Get iterator to all the rows in current sheet
            Iterator<Row> rowIterator = uren.getWerkblad().iterator();
            System.out.println(uren.getWerkblad().getSheetName());

            //Get iterator to all cells of current row
            Row row;
            Iterator<Cell> cellIterator;
            while (rowIterator.hasNext()) {
                row = rowIterator.next();
                cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    System.out.print(uren.celWaarde(cell));
                }
                System.out.println("");
            }

        } catch(Exception e) {
            System.out.println(e.getMessage());
        }

    }

    @Test
    public void testCel() throws Exception {
        HSSFRow row = uren.getRegel(51);
        Cell cell=row.getCell(0);
        System.out.println(uren.celWaarde(cell));

        cell=row.getCell(2);
        double start = Double.parseDouble(uren.celWaarde(cell));
        System.out.println(start + " - " + uren.tijdNaarTekst(start));

        row=uren.getRegel(52);
        cell=row.getCell(2);
        double einde = Double.parseDouble(uren.celWaarde(cell));
        System.out.println(einde + " - " + uren.tijdNaarTekst(einde));

        System.out.println(einde-start + " - " + uren.tijdNaarTekst(einde-start));
    }

    @Test
    public void testLeesOmschrijvingen() throws Exception {
        Iterator<Row> rowIterator = uren.getWerkblad().iterator();
        while(rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(0);
            System.out.printf("%-3s %s\n", row.getRowNum() + 1, uren.celWaarde(cell));
        }
    }

    @Test
    public void testWeekdagen() throws Exception {
        for (Weekdagen dag: Weekdagen.values()) {
            System.out.println(dag);
        }

        System.out.println(Weekdagen.MA.get());

        Weekdagen ma = Weekdagen.MA;
        System.out.println(Weekdagen.MA);

    }

    @Test
    public void testLeesWerktijden() throws Exception {
        List<Werkdag> tijden = uren.leesWerkTijden();
        for (Werkdag werkdag: tijden) {
            System.out.printf("%3d %3d %3d\n", werkdag.getDag(), werkdag.getTijd_in(), werkdag.getTijd_uit());
            System.out.printf("%3d %s %s %s\n", werkdag.getDag(), uren.tijdNaarTekst(werkdag.getTijd_in()), uren.tijdNaarTekst(werkdag.getTijd_uit()), uren.tijdNaarTekst(werkdag.getTijd_uit() - werkdag.getTijd_in()));
        }

    }

    @Test
    public void testLeesXlsnamen() throws Exception {

        String[] files = Diversen.leesFileNamen(dirXls, ExcelUren.URENMASK);
        Arrays.sort(files);

        System.out.println("Gevonden: " + files.length);
        for (String file: files) {
            System.out.println(file);
        }
    }

    @Test
    public void testTotaliseerProject() throws Exception {
        float totaal = uren.geefTaakDuur("Diversen ongeclassificeerd");
        assertEquals("totaliseer", 176, uren.nummerNaarMinuten(totaal));

        totaal = uren.geefTaakDuur("verlof");
        assertEquals("totaliseer", 0, uren.nummerNaarMinuten(totaal));
    }

    @Test
    public void testZoekTaakregel() throws Exception {
        int rij = uren.zoekTaakregel("HetHIS conversie");
        assertEquals("taakregel", 12, rij);

        rij = uren.zoekTaakregel("Diversen ongeclassificeerd");
        assertEquals("taakregel", 34, rij);

    }

    @Test
    public void testZoekProjectregel() throws Exception {
        int rij = uren.zoekProjectregel(uren.START_TEKST);
        assertEquals("projectregel", 5, rij);

        rij = uren.zoekProjectregel(uren.START1);
        assertEquals("projectregel", 36, rij);

    }

    @Test(expected = NullPointerException.class)
    public void testName() throws Exception {
        String tekst = "dit.is.een.test";
        assertEquals("dit", tekst.split("\\.")[0], "dit");
        assertEquals("dit", tekst.split("\\.").length , 4);
        System.out.println(tekst.split("\\.").length + " - " + tekst.split("\\.")[0]);

        tekst = "dit is een test";
        assertEquals("dit ...", tekst.split("\\.")[0], tekst);
        assertEquals("dit", tekst.split("\\.").length , 1);
        System.out.println(tekst.split("\\.").length + " - " + tekst.split("\\.")[0]);

        tekst = null;
        System.out.println(tekst.split("\\.").length + " - " + tekst.split("\\.")[0]);

    }

    @Test
    public void testAlleVerlof() throws Exception {
        String[] files = Diversen.leesFileNamen(dirXls, ExcelUren.URENMASK);
        Arrays.sort(files);

        int granttotal = 0;
        for (String xlsFile: files) {
            uren = new ExcelUren(dirXls, xlsFile);
            float totaal = uren.geefTaakDuur("verlof") / 60;
            float dagtotaal = uren.geefDagtotaal();
            granttotal += totaal;
            System.out.printf("file: %s, verlofuren: %2.2f, verlofdagen: %2.2f, uren gewerkt: %2.0f \n", xlsFile, totaal, (float) totaal / 60 / 9, dagtotaal);
            uren.sluitWerkboek();
        }
        System.out.printf("Totaal: verlofuren: %d, verlofdagen: %d", granttotal / 60, granttotal / 60 / 9);
    }

    @Test
    public void testDatumUitWeeknr() throws Exception {
        String[] dagen = uren.geefWeekDatums(23, 2012);
        assertEquals("begin", "04-06-2012", dagen[0]);
        assertEquals("einde", "10-06-2012", dagen[1]);
        System.out.println(dagen[0] + " - " + dagen[1]);

    }

    @Test
    public void testLognaam() {
        DateFormat df = new SimpleDateFormat("yyMM");
        System.out.println(df.format(new Date()));
    }

    @Test
    public void testDatumVanWeek() throws Exception {
        assertEquals("29-11-2013", uren.getDatumUitWeekDag(48, Calendar.FRIDAY, 2013));
    }

    @Test
    public void testJaarUitDirnaam() throws Exception {
        String dir = "/home/chris/IdeaProjects/uren2013";
        assertEquals(2013, uren.getJaarUitDirnaam(dir));
        dir = "/home/chris/IdeaProjects/uren";
        Calendar cal = Calendar.getInstance();
        int jaar = cal.get(Calendar.YEAR);
        assertEquals(jaar, uren.getJaarUitDirnaam(dir));
    }

    @Test
    public void testVerlof48() throws Exception {

        int granttotal = 0;
        String xlsFile = "CTS48.xls";
        uren = new ExcelUren("../../uren2013", xlsFile);
        float totaal = uren.geefTaakDuur("verlof") / 60;
        float dagtotaal = uren.geefDagtotaal();
        granttotal += totaal;
        System.out.printf("file: %s, verlofuren: %2.2f, verlofdagen: %2.2f, uren gewerkt: %2.0f \n", xlsFile, totaal, (float) totaal / 60 / 9, dagtotaal);
        uren.sluitWerkboek();
        System.out.printf("Totaal: verlofuren: %d, verlofdagen: %d\n", granttotal / 60, granttotal / 60 / 9);

        List<Verlofdag> verlofdagen =  uren.geefVerlofPerDag(48, 2013);
        for (Verlofdag dag : verlofdagen) {
            System.out.printf("%-10s %s  minuten: %6.1f, uren: %3.1f\n", dag.getDagnaam(), dag.getDatum(), dag.getMinuten(), dag.getMinuten()/60);
        }
    }

    @Test
    public void testProjectenMetTotaal() throws Exception {
        Map result = uren.projectenMetTotaal(Weekdagen.TOTAAL.get());
        System.out.println(result);

        result = uren.projectenMetTotaal(Weekdagen.VR.get());
        System.out.println(result);

    }

    @Test
    public void testUrenfileWeeknr() {
        Pattern pat = Pattern.compile("(.+)\\d{2}(\\.xls)", Pattern.CASE_INSENSITIVE);
        Matcher mat = pat.matcher(uren.getSheetFullName());
        while (mat.find()) {
            System.out.println(mat.group(1) + Diversen.getWeeknummer() + mat.group(2));
        }

        String urenfile = "/home/chris/Ideaprojects2/uren2013/Urenregistratie CT 37.xls";
        mat = pat.matcher(urenfile);
        while (mat.find()) {
            System.out.println(mat.group(1) + Diversen.getWeeknummer() + mat.group(2));
        }

    }

    @Test
    public void testMaakNieuwBestand() {
        try {
            File oud = new File("/home/chris/Ideaprojects2/uren2013/CTS47.xls");
            File nieuw = new File("/home/chris/Ideaprojects2/uren2013/CTS99.xls");
            FileUtils.copyFile(oud, nieuw);

            String path = "/home/chris/Ideaprojects2/uren2013";
            String file = "CTS99.xls";
            ExcelUren uren = new ExcelUren(path, file);
            uren.schrijfCel(2, 1, "Week: " + Diversen.getWeeknummer());

            for (int rij = uren.zoekProjectregel(uren.START_TEKST) - 1; rij < uren.zoekProjectregel(uren.STOP1); rij++) {
                uren.wisCellen(rij, Weekdagen.MA.get(), 5);
            }

            for (int rij = uren.zoekProjectregel(uren.START1); rij < uren.zoekProjectregel(uren.STOP_TEKST) - 1; rij++) {
                uren.wisCellen(rij, Weekdagen.MA.get(), 5);
            }

            uren.schrijfWerkboek();
            uren.sluitWerkboek();


        } catch (Exception e) {
            System.out.println(e.getMessage());
        }

    }
}

