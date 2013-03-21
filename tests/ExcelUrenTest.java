import junit.framework.Assert;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 20-3-13
 * Time: 12:13
 * To change this template use File | Settings | File Templates.
 */
public class ExcelUrenTest {
    static String dirXls = "../../uren2012";

    static Excel excel;

    @BeforeClass
    public static void setUp() throws Exception {
        excel = new Excel(dirXls + "/CTS47.xls");

    }

    @AfterClass
    public static void tearDown() throws Exception {

        excel.sluitWerkblad();
    }

    @Test
    public void testUren() throws Exception {
        try {
            //Get iterator to all the rows in current sheet
            Iterator<Row> rowIterator = excel.getWerkblad().iterator();
            System.out.println(excel.getWerkblad().getSheetName());

            //Get iterator to all cells of current row
            Row row;
            Iterator<Cell> cellIterator;
            while (rowIterator.hasNext()) {
                row = rowIterator.next();
                cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    System.out.print(excel.celWaarde(cell));
                }
                System.out.println("");
            }

        } catch(Exception e) {
            System.out.println(e.getMessage());
        }

    }

    @Test
    public void testCel() throws Exception {
        HSSFRow row = excel.getRegel(56);
        Cell cell=row.getCell(0);
        System.out.println(excel.celWaarde(cell));

        cell=row.getCell(2);
        double start = Double.parseDouble(excel.celWaarde(cell));
        System.out.println(start + " - " + excel.tijdNaarTekst(start));

        row=excel.getRegel(57);
        cell=row.getCell(2);
        double einde = Double.parseDouble(excel.celWaarde(cell));
        System.out.println(einde + " - " + excel.tijdNaarTekst(einde));

        System.out.println(einde-start + " - " + excel.tijdNaarTekst(einde-start));
    }

    @Test
    public void testLeesOmschrijvingen() throws Exception {
        Iterator<Row> rowIterator = excel.getWerkblad().iterator();
        while(rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(0);
            System.out.printf("%-3s %s\n", row.getRowNum() + 1, excel.celWaarde(cell));
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
        List<Werkdag> tijden = excel.leesWerkTijden();
        for (Werkdag werkdag: tijden) {
            System.out.printf("%3d %3d %3d\n", werkdag.getDag(), werkdag.getTijd_in(), werkdag.getTijd_uit());
            System.out.printf("%3d %s %s\n", werkdag.getDag(), excel.tijdNaarTekst(werkdag.getTijd_in()), excel.tijdNaarTekst(werkdag.getTijd_uit()));
        }

    }

    @Test
    public void testLeesXlsnamen() throws Exception {
        ExcelUrenView uren = new ExcelUrenView();

        String[] files = uren.leesXlsNamen(dirXls);
        Arrays.sort(files);

        System.out.println("Gevonden: " + files.length);
        for (String file: files) {
            System.out.println(file);
        }
    }

    @Test
    public void testTotaliseerProject() throws Exception {
        int totaal = excel.totaliseerDuur("Diversen ongeclassificeerd");
        Assert.assertEquals("totaliseer", totaal, 175);

        totaal = excel.totaliseerDuur("verlof");
        Assert.assertEquals("totaliseer", totaal, 540);
    }

    @Test
    public void testZoekProjectregel() throws Exception {
        int rij = excel.zoekProjectregel("HetHIS naconversie");
        Assert.assertEquals("projectregel", rij, 17);

        rij = excel.zoekProjectregel("Diversen ongeclassificeerd");
        Assert.assertEquals("projectregel", rij, 39);

    }

    @Test(expected = NullPointerException.class)
    public void testName() throws Exception {
        String tekst = "dit.is.een.test";
        Assert.assertEquals("dit", tekst.split("\\.")[0], "dit");
        Assert.assertEquals("dit", tekst.split("\\.").length , 4);
        System.out.println(tekst.split("\\.").length + " - " + tekst.split("\\.")[0]);

        tekst = "dit is een test";
        Assert.assertEquals("dit ...", tekst.split("\\.")[0], tekst);
        Assert.assertEquals("dit", tekst.split("\\.").length , 1);
        System.out.println(tekst.split("\\.").length + " - " + tekst.split("\\.")[0]);

        tekst = null;
        System.out.println(tekst.split("\\.").length + " - " + tekst.split("\\.")[0]);

    }

    @Test
    public void testAlleVerlof() throws Exception {
        ExcelUrenView uren = new ExcelUrenView();

        String[] files = uren.leesXlsNamen(dirXls);
        Arrays.sort(files);

        int granttotal = 0;
        for (String xlsFile: files) {
            excel = new Excel(dirXls + "/" + xlsFile);
            int totaal = excel.totaliseerDuur("verlof");
            float dagtotaal = excel.totaliseerDagtotaal() * 24;
            granttotal += totaal;
            System.out.printf("file: %s, verlofuren: %4d, verlofdagen: %2.2f, uren gewerkt: %2.0f \n", xlsFile, totaal, (float) totaal / 60 / 9, dagtotaal);
            excel.sluitWerkblad();
        }
        System.out.printf("Totaal: verlofuren: %d, verlofdagen: %d", granttotal / 60, granttotal / 60 / 9);
    }

    @Test
    public void testDatumUitWeeknr() throws Exception {
        String[] dagen = excel.geefWeekDatums(23, 2012);
        Assert.assertEquals("begin", dagen[0], "04-06-2012");
        Assert.assertEquals("einde", dagen[1], "10-06-2012");
        System.out.println(dagen[0] + " - " + dagen[1]);

    }


}
