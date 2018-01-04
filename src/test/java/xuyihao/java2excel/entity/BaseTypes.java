package xuyihao.java2excel.entity;

/**
 * Created by xuyh at 2018/1/4 11:06.
 */
public class BaseTypes {
    private int a = 12;
    private float b = 12.2F;
    private byte c = 1;
    private double d = 12.33D;
    private boolean e = true;
    private char f = 'a';
    private long g = 12654685465416L;
    private short h = 22;

    public int getA() {
        return a;
    }

    public void setA(int a) {
        this.a = a;
    }

    public float getB() {
        return b;
    }

    public void setB(float b) {
        this.b = b;
    }

    public byte getC() {
        return c;
    }

    public void setC(byte c) {
        this.c = c;
    }

    public double getD() {
        return d;
    }

    public void setD(double d) {
        this.d = d;
    }

    public boolean isE() {
        return e;
    }

    public void setE(boolean e) {
        this.e = e;
    }

    public char getF() {
        return f;
    }

    public void setF(char f) {
        this.f = f;
    }

    public long getG() {
        return g;
    }

    public void setG(long g) {
        this.g = g;
    }

    public short getH() {
        return h;
    }

    public void setH(short h) {
        this.h = h;
    }

    @Override
    public String toString() {
        return "BaseTypes{" +
                "a=" + a +
                ", b=" + b +
                ", c=" + c +
                ", d=" + d +
                ", e=" + e +
                ", f=" + f +
                ", g=" + g +
                ", h=" + h +
                '}';
    }
}
