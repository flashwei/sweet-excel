package com.github.read.template;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.regex.Pattern;


public class Test {

    @org.junit.Test
    public void test() {
        String pattern ="\\$\\{[A-Za-z]+(\\.[A-Za-z]+)+\\}";
        System.out.println(Pattern.matches(pattern,"${user.name.name}"));

        List<String> list = new ArrayList<>();

    }

    @org.junit.Test
    public void test2(){
        String str =  "${user.name.test}";
        StringBuilder stringBuilder = new StringBuilder();
        char[] valueChar = str.toCharArray();
        for (int i = 0; i <valueChar.length ; i++) {
            if(Objects.equals(valueChar[i],'$') || Objects.equals(valueChar[i],'{') || Objects.equals(valueChar[i],'}')){
                continue;
            }
            stringBuilder.append(valueChar[i]);
        }
        System.out.println(stringBuilder.toString());
    }
}
