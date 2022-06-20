package com.slf4j;

import org.junit.Test;
import org.slf4j.profiler.Profiler;

public class ProfilerTest {
    public void doA(){

    }

    public void doB(){

    }

    public void doOther(){

    }


    @Test
    public void testFunc01(){
        Profiler profiler = new Profiler("BASIC");
        profiler.start("A");
        doA();

        profiler.start("B");
        doB();

        profiler.start("OTHER");
        doOther();
    }
}
