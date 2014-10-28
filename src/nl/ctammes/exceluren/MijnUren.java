package nl.ctammes.exceluren;

import nl.ctammes.common.Diversen;
import nl.ctammes.common.MijnIni;
import nl.ctammes.common.MijnLog;
import nl.ctammes.exceluren.ExcelUren;

import javax.swing.*;
import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.*;
import java.util.List;
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
    private JButton btnCopyText;

    private ExcelUren uren;

    private static int weekNr;
    private static MijnIni ini = null;
    private static String inifile = "MijnUren.ini";

    private static String dirXls = "/home/chris/IdeaProjects/uren";
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
                    Calendar cal = Calendar.getInstance();
                    int jaar = cal.get(Calendar.YEAR);              // dit jaar
                    int dirjaar = ExcelUren.getJaarUitDirnaam(dirXls);   // jaar waarvan de werkbladen gelezen worden
                    // per bestand projecten inlezen in gesorteerde lijst (TreeSet) zonder duplicaten (Set)
                    Set<String> projecten = new TreeSet<String>();
                    for (String xlsFile: files) {
                        uren = new ExcelUren(dirXls, xlsFile);
                        if (dirjaar < jaar || uren.getWeeknrUitFilenaam(xlsFile) <= weekNr) {
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
                Calendar cal = Calendar.getInstance();
                int jaar = cal.get(Calendar.YEAR);                  // dit jaar
                int dirjaar = ExcelUren.getJaarUitDirnaam(dirXls);  // jaar waarvan de werkbladen gelezen worden
                DefaultListModel listModel = new DefaultListModel();
                if (!project.equals("") && !project.equals("tijdinuit")) {
                    String tekst = String.format("%s\n", project);
                    listModel.addElement(tekst);
                }
                for (String xlsFile: files) {
                    uren = new ExcelUren(dirXls, xlsFile);
                    if (dirjaar < jaar || uren.getWeeknrUitFilenaam(xlsFile) <= weekNr) {
                        float totaal = uren.geefTaakDuur(project);
                        float dagtotaal = uren.geefDagtotaal();
                        if (project.equals("verlof") || (!project.equals("verlof") && totaal > 0)) {
                            grandtotal += totaal;
                            String tekst = String.format("file: %s, minuten: %6.1f, uren: %3.1f, dagen: %4.2f, uren gewerkt: %2.0f \n"
                                    , xlsFile, (float) totaal, (float) totaal / 60, (float) totaal / 60 / uren.URENPERDAG, dagtotaal);
                            listModel.addElement(tekst);

                            // verlof tonen per dag
                            List<Verlofdag> verlofdagen =  uren.geefVerlofPerDag(uren.getWeeknrUitFilenaam(xlsFile), dirjaar);
                            for (Verlofdag dag : verlofdagen) {
                                tekst = String.format("%-10s %s  minuten: %6.1f, uren: %3.1f\n", dag.getDagnaam(), dag.getDatum(), dag.getMinuten(), dag.getMinuten()/60);
                                listModel.addElement(tekst);
                            }
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
        btnCopyText.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                StringBuffer tekst = new StringBuffer("");
                ListModel model = lstResultaat.getModel();
                for (int i = 0; i < model.getSize(); i++) {
                    tekst.append(model.getElementAt(i).toString());
                }
                tekstNaarKlembord(tekst.toString());
            }
        });
    }

    private void tekstNaarKlembord(String tekst) {
        if (tekst.length() > 0) {
            StringSelection stringSelection = new StringSelection(tekst);
            Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
            clipboard.setContents( stringSelection, stringSelection);
        }
    }

    public static void main(String[] args) {

        // initialiseer logger
        Logger log = Logger.getLogger(MijnUren.class.getName());

        // Bepaal huidige weeknummer en jaar
        // Wijzigt als je een andere directory kiest
        Calendar cal = Calendar.getInstance();
        weekNr = cal.get(Calendar.WEEK_OF_YEAR);
        int jaar = cal.get(Calendar.YEAR);

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
            String dir = ini.lees("Algemeen", "dirxls");
            if (dir != null) {
                dirXls = dir;
            } else {
                dirXls = dirXls.concat(String.valueOf(jaar));
                if (new File(dirXls).exists()) {
                    ini.schrijf("Algemeen", "dbxls", dirXls);
                }
            }
        } else {
            ini = new MijnIni(inifile);
            dirXls = dirXls.concat(String.valueOf(jaar));
            ini.schrijf("Algemeen", "dirxls", dirXls);
            log.info("Inifile " + inifile + " aangemaakt en gevuld");
        }

        JFrame frame = new JFrame("MijnUren");
        frame.setContentPane(new MijnUren().mainPanel);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setLocation(200,200);
//        frame.setSize();
        frame.pack();
        frame.setVisible(true);

    }


}