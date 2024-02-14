package examenmetodos;
import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.InputEvent;
import java.util.Scanner;

public class Mouse {
    private Robot robot;
    private int distanciaY;
    private int distanciaX;
    public Mouse(){
        distanciaY=20;
        distanciaX=200;
    } 
    public void moverMouse(int x,int y){
        robot.mouseMove(x, y);
    }
    public void darClick(){
        robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
        robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }
    public void seleccionarGrafica(int n){
        int nuevaDistanciaY=distanciaY*n;
        moverMouse(distanciaX,distanciaY);
        robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
        moverMouse(distanciaX,distanciaY);
        robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        
    }
    
    public void setDistanciaY(int d){
        distanciaY=d;
    }
    public int getDistanciaY(){
        return distanciaY;
    }
    public void setDistanciaX(int d){
        distanciaX=d;
    }
    public int getDistanciaX(){
        return distanciaX;
    }
    
}
