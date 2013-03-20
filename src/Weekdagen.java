/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 20-3-13
 * Time: 16:43
 * To change this template use File | Settings | File Templates.
 */
public enum Weekdagen {
    // dagen komen overeen met kolomnummers in urensheet
    MA(2), DI(3), WO(4), DO(5), VR(6), ZA(7), ZO(8), TOTAAL(9);

    private int dagnr;

    private Weekdagen(int dagnr) {
        this.dagnr = dagnr;
    }

    public int get() {
        return dagnr;

    }

}
