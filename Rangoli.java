import java.awt.*;
import java.awt.event.*;
public class Rangoli extends Frame implements AdjustmentListener
{
        Scrollbar r,g,b;
        Label rl,gl,bl;
        Canvas cvas;
        public Rangoli()
        {
                setBackground(Color.cyan);
                Panel p=new Panel();
                r=new Scrollbar(Scrollbar.HORIZONTAL,0,0,0,255);
                g=new Scrollbar(Scrollbar.HORIZONTAL,0,0,0,255);
                b=new Scrollbar(Scrollbar.HORIZONTAL,0,0,0,255);
                r.setUnitIncrement(5);
                g.setUnitIncrement(5);
                b.setUnitIncrement(5);
                r.setBlockIncrement(15);
                r.setBlockIncrement(15);
                r.setBlockIncrement(15);
                r.addAdjustmentListener(this);
                g.addAdjustmentListener(this);
                b.addAdjustmentListener(this);
                rl=new Label("RED");
                bl=new Label("BLUE");
                gl=new Label("GREEN");
                p.setLayout(new GridLayout(3,2,10,8));
                p.add(rl);
                p.add(r);
                p.add(gl);
                p.add(g);
                p.add(bl);
                p.add(b);
                cvas=new Canvas();
                add(cvas,"Center");
                add(p,"South");
                setTitle("Rangoli");
                setSize(300,300);
                setVisible(true);
        }

        public void adjustmentValueChanged(AdjustmentEvent e)
        {
                int rv=r.getValue();
                int gv=g.getValue();
                int bv=b.getValue();

                rl.setText("RED : "+rv);
                gl.setText("GREEN : "+gv);
                bl.setText("BLUE : "+bv);

                cvas.setBackground(new Color(rv,gv,bv));
        }

        public static void main(String s[])
        {
                new Rangoli();
        }
}