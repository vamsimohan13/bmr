/*
 * Copyright 2015 vamsi.mohan.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package mnm.buildmyreport;

import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 *
 * @author vamsi.mohan
 */
public class mynt {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
//        // TODO code application logic here
//        System.out.println(invertBits(40));
//        System.out.println(invert(40));
//    }
//        
//        static int invertBits(int n){
//        int pow2=1;int bit=0;
//            int newnum=0;
//            while(n>0) {
//              bit = (n & 1);
//              if(bit==0)
//                  newnum+= pow2;
//              n=n>>1;
//              pow2*=2;
//          }
//          return newnum;
//        }
//
//    private static int invert(int i) {
//        return ~i;
//    }
           //Calculating square root of even numbers from 1 to N       
    int min = 1;
    int max = 10000000;

    List<Integer> sourceList = new ArrayList<Integer>();
    for (int i = min; i < max; i++) {
        sourceList.add(i);
    }

    List<Double> result = new LinkedList<Double>();


    //Collections approach
    long t0 = System.nanoTime();
    long elapsed = 0;
    for (Integer i : sourceList) {
        if(i % 2 == 0){
            result.add( doSomeCalculate(i));
        }
    }
    elapsed = System.nanoTime() - t0;       
    System.out.printf("Collections: Elapsed time:\t %d ns \t(%f seconds)%n", elapsed, elapsed / Math.pow(10, 9));


    //Stream approach
    Stream<Integer> stream = sourceList.stream();       
    t0 = System.nanoTime();
    result = stream.filter(i -> i%2 == 0).map(i -> doSomeCalculate(i))
            .collect(Collectors.toList());
    elapsed = System.nanoTime() - t0;       
    System.out.printf("Streams: Elapsed time:\t\t %d ns \t(%f seconds)%n", elapsed, elapsed / Math.pow(10, 9));


    //Parallel stream approach
    stream = sourceList.stream().parallel();        
    t0 = System.nanoTime();
    result = stream.filter(i -> i%2 == 0).map(i ->  doSomeCalculate(i))
            .collect(Collectors.toList());
    elapsed = System.nanoTime() - t0;       
    System.out.printf("Parallel streams: Elapsed time:\t %d ns \t(%f seconds)%n", elapsed, elapsed / Math.pow(10, 9));      
}

static double doSomeCalculate(int input) {
    for(int i=0; i<100000; i++){
        Math.sqrt(i+input);
    }
    return Math.sqrt(input);
}
    }
    

