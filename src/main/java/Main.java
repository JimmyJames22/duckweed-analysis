import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;

import javax.imageio.ImageIO;
import javax.swing.*;
import java.awt.*;
import java.awt.Color;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.awt.image.BufferStrategy;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.Scanner;

public class Main implements MouseListener, KeyListener {

    public static void main(String[] args) {
        Main m = new Main();
    }

    public JFrame frame;
    public JPanel panel;
    public Canvas canvas;
    public BufferStrategy bs;

    public boolean control = false;
    public boolean scaling1 = true;
    public boolean scaling2 = true;

    public double p1x = -10;
    public double p1y = -10;
    public double p2x = -10;
    public double p2y = -10;

    public double pixPerCm;

    public int leftBound = 180;
    public int rightBound = 830;

    ArrayList<Integer> red;
    ArrayList<Integer> blue;
    ArrayList<Integer> green;

    public ArrayList<Page> pages;
    public Page page;

    public String[] treatmentVals;

    public int numReps;
    public int treatmentNum;
    public String pageName;

    public int repNum;

    public BufferedImage image;
    public BufferedImage blurredImage;
    public BufferedImage imageControl;

    public double aspectRatio;
    public double imageScale;
    public int blur = 5;


    public int WIDTH = 1000;
    public int HEIGHT = 1000;

    public int argb;

    public boolean done = false;

    public boolean box = false;

    public boolean continu = true;

    public Main(){
        graphicsSetup();
        setUpExcel();
        run();
    }

    public void run(){
        dataSetup();
        setScale();
//        run();
    }

    public void setScale(){
        render();
        if(!scaling1 && !scaling2){
            pixPerCm = Math.sqrt(Math.pow(p1x - p2x, 2) + Math.pow(p1y - p2y, 2));
            System.out.println("Loading image data");
            loadImage();
            crunch();
//            done = true;
        } else {
            System.out.println("Click on scaling points");
        }
    }

    public void setUpExcel(){
        try {
            System.out.println("Enter page name");
            Scanner pageScanner = new Scanner(System.in);
            String page = pageScanner.nextLine().toUpperCase();

            System.out.println("Enter num reps");
            Scanner repScanner = new Scanner(System.in);
            int rep = repScanner.nextInt();

            System.out.println("Enter treatments separated by commas");
            Scanner treatmentScanner = new Scanner(System.in);
            String[] treatmentArr = treatmentScanner.nextLine().replace(" ", "").split(",");

            pages.add(new Page(page, rep, treatmentArr));

            pageName = page;
            treatmentVals = treatmentArr;
            numReps = rep;

            // connect to excel doc
            FileInputStream inputStream = new FileInputStream(new File("/users/james/IdeaProjects/DYODataCollection Edge Finding Fill/james.xls"));
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet worksheet = workbook.createSheet(pageName);

            // create bold font style for headers
            Font boldFont = workbook.createFont();
            boldFont.setBold(true);
            CellStyle boldStyle = workbook.createCellStyle();
            boldStyle.setFont(boldFont);

            // new row and cells
            Row row0 = worksheet.createRow((short) 2);
            Cell table1Label = row0.createCell((short) 0);
            table1Label.setCellValue("Pixel Count");
            table1Label.setCellStyle(boldStyle);
            Row row1 = worksheet.createRow((short) 3);
            for(int x=0; x<treatmentVals.length; x++){
                Cell treatment = row1.createCell((short) (x+1));
                treatment.setCellValue(treatmentVals[x]);
                treatment.setCellStyle(boldStyle);
            }

            for(int x=0; x<numReps; x++){
                Row row = worksheet.createRow((short) (x+4));
                Cell cell = row.createCell((short) 0);
                cell.setCellValue("Rep. " + (x+1));
                cell.setCellStyle(boldStyle);
            }

            Row row12 = worksheet.createRow((short) numReps+6);
            Cell table2Label = row12.createCell(0);
            table2Label.setCellValue("Average RGB");
            table2Label.setCellStyle(boldStyle);
            Row row13 = worksheet.createRow((short) numReps+6);

            for(int x=0; x<treatmentVals.length; x++){
                Cell r = row13.createCell((short) (((x+1)*3)-2));
                r.setCellValue(treatmentVals[x] + " R");
                r.setCellStyle(boldStyle);

                Cell g = row13.createCell((short) (((x+1)*3)-1));
                g.setCellValue(treatmentVals[x] + " G");
                g.setCellStyle(boldStyle);

                Cell b = row13.createCell((short) ((x+1)*3));
                b.setCellValue(treatmentVals[x] + " B");
                b.setCellStyle(boldStyle);
            }

            for(int x=0; x<numReps; x++){
                Row row = worksheet.createRow((short) (x+numReps+7));
                Cell cell = row.createCell((short) 0);
                cell.setCellValue("Rep. " + (x+1));
                cell.setCellStyle(boldStyle);
            }

            FileOutputStream fos = new FileOutputStream(new File("/users/james/IdeaProjects/DYODataCollection Edge Finding Fill/james.xls"));
            workbook.write(fos);
            fos.flush();
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
            setUpExcel();
            return;
        }

        System.out.println("Excel file formatted");
        System.out.println();

    }

    public void crunch(){
        double redVal = 0;
        double greenVal = 0;
        double blueVal = 0;

        double count = red.size();

        for(int x=0; x<count; x++){
            redVal += red.get(x);
            blueVal += blue.get(x);
            greenVal += green.get(x);
        }

        redVal /= count;
        greenVal /= count;
        blueVal /= count;

        count /= (Math.pow(pixPerCm, 2));

        System.out.println("Number of cm^2: " + count);
        System.out.println("r: " + redVal + " g: " + greenVal + " b: " + blueVal);
        System.out.println();

        try {
            FileInputStream inputStream = new FileInputStream(new File("/users/james/IdeaProjects/DYODataCollection Edge Finding Fill/james.xls"));
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet worksheet = workbook.getSheet(pageName);

            Cell pixelCell = worksheet.getRow(repNum+3).createCell(treatmentNum);
            pixelCell.setCellValue(count);

            Cell rCell = worksheet.getRow(repNum+numReps+6).createCell((treatmentNum*3)-2);
            rCell.setCellValue(redVal);
            Cell gCell = worksheet.getRow(repNum+numReps+6).createCell((treatmentNum*3)-1);
            gCell.setCellValue(greenVal);
            Cell bCell = worksheet.getRow(repNum+numReps+6).createCell(treatmentNum*3);
            bCell.setCellValue(blueVal);

            inputStream.close();

            FileOutputStream outputStream = new FileOutputStream("/users/james/IdeaProjects/DYODataCollection Edge Finding Fill/james.xls");
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
//        System.out.println("Press 'n' to continue");
        run();
    }

    public void loadImage(){
        Color color;
        Color prevColor;

        double ave;
        double avePrev;

        double dev;
        double devPrev;

        double dist;

        int r;
        int g;
        int b;
        int rPrev;
        int gPrev;
        int bPrev;

        for(int x = (int) (leftBound*imageScale)+(3*blur); x<rightBound*imageScale; x+=blur){
            for(int y=(3*blur); y<image.getHeight(); y+=blur) {
                color = new Color(blurredImage.getRGB(x-blur, y-blur));
                if(color.getRGB() == argb){
                    continue;
                }

                prevColor = new Color(blurredImage.getRGB(x-(3*blur), y-(3*blur)));

                r = color.getRed();
                g = color.getGreen();
                b = color.getBlue();
                rPrev = prevColor.getRed();
                gPrev = prevColor.getGreen();
                bPrev = prevColor.getBlue();

                dist = Math.sqrt((Math.pow(r - rPrev, 2) + Math.pow(g - gPrev, 2)) + Math.pow(b - bPrev, 2));

                ave = (double)(r+g+b)/3;
                avePrev = (double)(rPrev+gPrev+bPrev)/3;

                dev = Math.abs(r - ave)+Math.abs(g - ave)+Math.abs(b - ave);
                devPrev = Math.abs(rPrev - avePrev)+Math.abs(gPrev - avePrev)+Math.abs(bPrev - avePrev);

                if(g>b || r>b) {
                    if ((dist > 20 && (((ave < 80 && avePrev < 80) || (ave > 120 && avePrev < 80) || (avePrev > 120 && ave < 80)) && (dev > 30 || devPrev > 30) && (Math.abs(dev - devPrev) > 0)))) {
                        for (int x1 = x - blur; x1 <= x; x1++) {
                            for (int y1 = y - blur; y1 <= y; y1++) {
                                red.add(new Color(imageControl.getRGB(x1, y1)).getRed());
                                green.add(new Color(imageControl.getRGB(x1, y1)).getGreen());
                                blue.add(new Color(imageControl.getRGB(x1, y1)).getBlue());
                                image.setRGB(x1, y1, new Color(0, 0, 255).getRGB());
                            }
                        }
                    } else if (ave < 60 && dev < 80) {
                        for (int x1 = x - blur; x1 <= x; x1++) {
                            for (int y1 = y - blur; y1 <= y; y1++) {
                                red.add(new Color(imageControl.getRGB(x1, y1)).getRed());
                                green.add(new Color(imageControl.getRGB(x1, y1)).getGreen());
                                blue.add(new Color(imageControl.getRGB(x1, y1)).getBlue());
                                image.setRGB(x1, y1, new Color(238, 0, 255).getRGB());
                            }
                        }
                    } else if (r > g && (ave < 100 || dev > 72) && g>b) {
                        for (int x1 = x - blur; x1 <= x; x1++) {
                            for (int y1 = y - blur; y1 <= y; y1++) {
                                red.add(new Color(imageControl.getRGB(x1, y1)).getRed());
                                green.add(new Color(imageControl.getRGB(x1, y1)).getGreen());
                                blue.add(new Color(imageControl.getRGB(x1, y1)).getBlue());
                                image.setRGB(x1, y1, new Color(238, 255, 0).getRGB());
                            }
                        }
                    } else if(g > r && (g-b > 70)){
                        for (int x1 = x - blur; x1 <= x; x1++) {
                            for (int y1 = y - blur; y1 <= y; y1++) {
                                red.add(new Color(imageControl.getRGB(x1, y1)).getRed());
                                green.add(new Color(imageControl.getRGB(x1, y1)).getGreen());
                                blue.add(new Color(imageControl.getRGB(x1, y1)).getBlue());
                                image.setRGB(x1, y1, new Color(238, 0, 0).getRGB());
                            }
                        }
                    }
                }
            }
        }
        System.out.println("Done taking data");
        render();
    }

    public void dataSetup(){
        System.out.println("Enter filename");
        Scanner fileNameIn = new Scanner(System.in);
        String fileName = fileNameIn.nextLine();

        frame.setTitle(fileName);

        File imageFile = new File(fileName);
        try {
            image = ImageIO.read(imageFile);
            blurredImage = ImageIO.read(imageFile);
            imageControl = ImageIO.read(imageFile);
        } catch(IOException e){
            e.printStackTrace();
            dataSetup();
            return;
        }

        int aveR;
        int aveG;
        int aveB;
        int num;

        for(int x=blur; x<image.getWidth(); x+=blur){
            for(int y=blur; y<image.getHeight(); y+=blur){
                aveR = 0;
                aveG = 0;
                aveB = 0;
                num = 0;

                for(int x1=x-blur; x1<=x; x1++){
                    for(int y1=y-blur; y1<=y; y1++){
                        Color color = new Color(image.getRGB(x1, y1));
                        aveR += color.getRed();
                        aveG += color.getGreen();
                        aveB += color.getBlue();
                        num ++;
                    }
                }

                aveR /= num;
                aveG /= num;
                aveB /= num;

                int newColor = new Color(aveR, aveG, aveB).getRGB();

                for(int x1=x-blur; x1<=x; x1++){
                    for(int y1=y-blur; y1<=y; y1++){
                        image.setRGB(x1, y1, newColor);
                        blurredImage.setRGB(x1, y1, newColor);
                    }
                }
            }
        }

        aspectRatio = (double)(image.getHeight())/(double)(image.getWidth());

        red = new ArrayList<>();
        green = new ArrayList<>();
        blue = new ArrayList<>();

        Graphics2D g = (Graphics2D) bs.getDrawGraphics();
        g.clearRect(0, 0, WIDTH, HEIGHT);
        g.dispose();
        bs.show();


        System.out.println("Choose treatment");
        System.out.println("--------------------");
        for(int x=0; x<treatmentVals.length; x++){
            System.out.println((x+1) + ". " + treatmentVals[x]);
        }
        Scanner treatmentScan = new Scanner(System.in);
        treatmentNum = treatmentScan.nextInt();
        System.out.println();

        System.out.println("Enter run (int 1 - " + numReps + ")");
        Scanner repScan = new Scanner(System.in);
        repNum = repScan.nextInt();

        if(!scaling1 && !scaling2) {
            System.out.println("Reset scaling?");
            Scanner scaleScanner = new Scanner(System.in);
            if (scaleScanner.nextLine().toLowerCase().contains("y")) {
                scaling1 = true;
                scaling2 = true;
            }
        }

        imageScale = image.getWidth()/(double)(1000);
    }

    public void graphicsSetup() {
        frame = new JFrame("Hales Class Grapher");
        frame.setSize(WIDTH, HEIGHT);
        panel = new JPanel();
        panel.setSize(new Dimension(WIDTH, HEIGHT));

        canvas = new Canvas();
        canvas.setBounds(0, 0, WIDTH, HEIGHT);
        canvas.addMouseListener(this);
//        canvas.addKeyListener(this);

        panel.add(canvas);
        frame.add(panel);

        frame.setDefaultCloseOperation(3);
        frame.pack();
        frame.setResizable(true);
        frame.setVisible(true);

        canvas.createBufferStrategy(1);
        bs = canvas.getBufferStrategy();
        canvas.requestFocus();

        argb = new Color(255, 0, 150).getRGB();
        pages = new ArrayList<Page>();
    }

    public void render(){
        Graphics2D g = (Graphics2D) bs.getDrawGraphics();
        g.setColor(Color.red);
        if(control){
            g.drawImage(imageControl, 0, 0, 1000, (int) (1000*aspectRatio),null);
        } else {
            g.drawImage(image, 0, 0, 1000, (int) (1000*aspectRatio), null);
        }
        g.drawRect(leftBound, 0, rightBound - leftBound, (int) (image.getHeight()/imageScale));
        g.drawLine((int) (p1x/imageScale), (int) (p1y/imageScale), (int) (p2x/imageScale), (int) (p2y/imageScale));
        g.dispose();
        bs.show();
    }

    @Override
    public void keyTyped(KeyEvent e) {

    }

    /**
     * Invoked when a key has been pressed.
     * See the class description for {@link KeyEvent} for a definition of
     * a key pressed event.
     *
     * @param e the event to be processed
     */
    @Override
    public void keyPressed(KeyEvent e) {
        if(e.getKeyChar() == 'c' && control == false){
            control = true;
            aspectRatio = (double)(imageControl.getHeight())/(double)(imageControl.getWidth());
            render();
        }
        if(e.getKeyChar() == 'b'){
            box = true;
            render();
        }
    }

    /**
     * Invoked when a key has been released.
     * See the class description for {@link KeyEvent} for a definition of
     * a key released event.
     *
     * @param e the event to be processed
     */
    @Override
    public void keyReleased(KeyEvent e) {
        if(done){
            if(e.getKeyChar() == 'n'){
                done = false;
                run();
            }
        }
        if(e.getKeyChar() == 'c' && control == true){
            control = false;
            aspectRatio = (double)(image.getHeight())/(double)(image.getWidth());
            render();
        }
        if(e.getKeyChar() == 'b'){
            box = false;
            render();
        }
    }

    /**
     * Invoked when the mouse button has been clicked (pressed
     * and released) on a component.
     *
     * @param e the event to be processed
     */
    @Override
    public void mouseClicked(MouseEvent e) {
        try {
            if (control) {
                System.out.println("x: " + e.getX() + ", y: " + e.getY() + ", color: " + new Color(imageControl.getRGB((int) (e.getX()*imageScale), (int) (e.getY()*imageScale))));
            } else {
                System.out.println("x: " + e.getX() + ", y: " + e.getY() + ", color: " + new Color(image.getRGB(e.getX(), e.getY())));
            }

            if(scaling2 && !scaling1){
                p2x = e.getX()*imageScale;
                p2y = e.getY()*imageScale;
                scaling2 = false;
                setScale();
            } else if(scaling1){
                p1x = e.getX()*imageScale;
                p1y = e.getY()*imageScale;
                p2x = e.getX()*imageScale;
                p2y = e.getY()*imageScale;
                scaling1 = false;
            }
        } catch(ArrayIndexOutOfBoundsException ex){
            ex.printStackTrace();
        }
    }

    /**
     * Invoked when a mouse button has been pressed on a component.
     *
     * @param e the event to be processed
     */
    @Override
    public void mousePressed(MouseEvent e) {

    }

    /**
     * Invoked when a mouse button has been released on a component.
     *
     * @param e the event to be processed
     */
    @Override
    public void mouseReleased(MouseEvent e) {

    }

    /**
     * Invoked when the mouse enters a component.
     *
     * @param e the event to be processed
     */
    @Override
    public void mouseEntered(MouseEvent e) {

    }

    /**
     * Invoked when the mouse exits a component.
     *
     * @param e the event to be processed
     */
    @Override
    public void mouseExited(MouseEvent e) {

    }
}
