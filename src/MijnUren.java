import nl.ctammes.common.Diversen;
import nl.ctammes.common.MijnIni;
import nl.ctammes.common.MijnLog;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 21-3-13
 * Time: 8:59
 * To change this template use File | Settings | File Templates.
 */
public class MijnUren {

    protected JPanel mainPanel;

    private JTextField txtExcelDir;
    private JCheckBox chkVerlof;
    private JCheckBox chkProject;
    private JComboBox cmbProjecten;
    private JButton btnStart;
    private JList lstResultaat;
    private JButton btnInlezen;
    private JCheckBox chkTijdInUit;

    private ExcelUren uren;

    private static int weekNr;
    private static MijnIni ini = null;
    private static String inifile = "MijnUren.ini";

    private static String dirXls = "/home/chris/IdeaProjects/uren2012";
    private String[] files = null;

    public MijnUren() {

        txtExcelDir.setText(dirXls);

        btnInlezen.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                // lees en sorteer alle xls-bestanden
                dirXls = txtExcelDir.getText();
                ini.schrijf("Algemeen", "dirxls", dirXls);
                files = Diversen.leesFileNamen(dirXls, ExcelUren.URENMASK);
                Arrays.sort(files);

                if (files.length > 0) {
                    // per bestand projecten inlezen in gesorteerde lijst (TreeSet) zonder duplicaten (Set)
                    Set<String> projecten = new TreeSet<String>();
                    for (String xlsFile: files) {
                        uren = new ExcelUren(dirXls, xlsFile);
                        if (uren.getWeeknrUitFilenaam(xlsFile) <= weekNr) {
                            // nieuwe projecten toevoegen - geen duplicaten (want Set)
                            projecten.addAll(uren.leesProjecten());
                        }
                        uren.sluitWerkblad();
                    }
                    if (projecten.size() > 0) {
                        cmbProjecten.removeAll();
                        for (String project: projecten) {
                            //TODO wat is een ComboBoxModel ??
                            cmbProjecten.addItem(project);
                        }
                    }
                }
            }
        });
        btnStart.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                String project = "";
                if (chkVerlof.isSelected()) {
                    project = "verlof";
                } else if (chkTijdInUit.isSelected()) {
                    project = "tijdinuit";
                } else if (chkProject.isSelected() && cmbProjecten.getSelectedIndex()>=0) {
                    project = cmbProjecten.getSelectedItem().toString();

                }

                // verwerk de files
                int grandtotal = 0;
                int grandsaldo = 0;
                DefaultListModel listModel = new DefaultListModel();
                if (!project.equals("") && !project.equals("tijdinuit")) {
                    String tekst = String.format("%s\n", project);
                    listModel.addElement(tekst);
                }
                for (String xlsFile: files) {
                    uren = new ExcelUren(dirXls, xlsFile);
                    if (uren.getWeeknrUitFilenaam(xlsFile) <= weekNr) {
                        float totaal = uren.geefTaakDuur(project);
                        float dagtotaal = uren.geefDagtotaal();
                        if (project.equals("verlof") || (!project.equals("verlof") && totaal > 0)) {
                            grandtotal += totaal;
                            String tekst = String.format("file: %s, minuten: %6.1f, uren: %3.1f, dagen: %4.2f, uren gewerkt: %2.0f \n", xlsFile, (float) totaal, (float) totaal / 60, (float) totaal / 60 / uren.URENPERDAG, dagtotaal);
                            listModel.addElement(tekst);
                        }
                        if (project.equals("tijdinuit")) {
                            String tekst = String.format("file: %s\n", xlsFile);
                            listModel.addElement(tekst);
                            List<Werkdag> tijden = uren.leesWerkTijden();
                            int gewerkt = 0;
                            int dagen = 0;
                            for (Werkdag werkdag: tijden) {
                                int in = werkdag.getTijd_in();
                                int uit = werkdag.getTijd_uit();
                                if (in > 0 && uit > 0) {
                                    dagen++;
                                    gewerkt += uit - in;
                                    tekst = String.format("\tdag: %s, in: %s, uit: %s, uren gewerkt: %s\n", werkdag.getDagnaamKort(), uren.tijdNaarTekst(in), uren.tijdNaarTekst(uit), uren.tijdNaarTekst(uit - in));
                                    listModel.addElement(tekst);
                                }
                            }
                            String urenGewerkt = uren.tijdNaarTekst(gewerkt);
                            // Corrigeer evt. gewerkte dagen
                            dagen = (dagen > uren.DAGENPERWEEK) ? uren.DAGENPERWEEK : dagen;
                            float saldo = gewerkt - (dagen * uren.URENPERDAG * 60);
                            grandsaldo +=  saldo;
                            String sign = (saldo<0) ? "-" : "";
                            tekst = String.format("\tTotaal: uren: %s, saldo: %s%s", urenGewerkt, sign, uren.tijdNaarTekst(Math.abs(saldo)));
                            listModel.addElement(tekst);
                        }
                    }
                    uren.sluitWerkblad();
                }
                String tekst = "";
                if (grandtotal > 0) {
                    tekst = String.format("Totaal: uren: %d, dagen: %d", grandtotal / 60, grandtotal / 60 / 9);
                } else {
                    String sign = (grandsaldo<0) ? "-" : "";

                    tekst = String.format("Saldo: uren: %s%s", sign, uren.tijdNaarTekst(Math.abs(grandsaldo / 60)));
                }
                listModel.addElement(tekst);

                lstResultaat.setModel(listModel);
            }
        });
        chkProject.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                cmbProjecten.setEnabled(chkProject.isEnabled());
            }
        });
    }

    public static void main(String[] args) {

        // initialiseer logger
        Logger log = Logger.getLogger(MijnUren.class.getName());

        String logDir = ".";
        String logNaam = "MijnUren.log";
        try {
            MijnLog mijnlog = new MijnLog(logDir, logNaam, true);
            log = mijnlog.getLog();
            log.setLevel(Level.INFO);
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }

        // inifile lezen of initieel vullen
        if (new File(inifile).exists()) {
            ini = new MijnIni(inifile);
            dirXls = ini.lees("Algemeen", "dirxls");
        } else {
            ini = new MijnIni(inifile);
            ini.schrijf("Algemeen", "dirxls", dirXls);
            log.info("Inifile " + inifile + " aangemaakt en gevuld");
        }

        // Bepaal huidige weeknummer
        Calendar cal = Calendar.getInstance();
        weekNr = cal.get(Calendar.WEEK_OF_YEAR);

        JFrame frame = new JFrame("MijnUren");
        frame.setContentPane(new MijnUren().mainPanel);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setLocation(200,200);
//        frame.setSize();
        frame.pack();
        frame.setVisible(true);

    }


}
