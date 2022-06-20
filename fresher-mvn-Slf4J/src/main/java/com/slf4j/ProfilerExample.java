package com.slf4j;

import org.slf4j.profiler.Profiler;
import org.slf4j.profiler.TimeInstrument;

public class ProfilerExample {
    public void demoMethod1(){
        double sum = 0;
        for(int i=0; i< 1000; i++){
            sum = sum+(Math.pow(i, 2));
        }
        System.out.println("Sum of squares of the numbers from 1 to 10000: "+sum);
    }
    public void demoMethod2(){
        int sum = 0;
        for(int i=0; i< 10000; i++){
            sum = sum+i;
        }
        System.out.println("Sum of the numbers from 1 to 10000: "+sum);
    }
    public static void main(String[] args) {
        ProfilerExample obj = new ProfilerExample();

        //Creating a profiler
        Profiler profiler = new Profiler("Sample");

        //Starting a child stop watch and stopping the previous one.
        profiler.start("Task 1");
        obj.demoMethod1();

        //Starting another child stop watch and stopping the previous one.
        profiler.start("Task 2");
        obj.demoMethod2();

        //Stopping the current child watch and the global watch.
        TimeInstrument tm = profiler.stop();

        //printing the contents of the time instrument
        tm.print();
    }
}
