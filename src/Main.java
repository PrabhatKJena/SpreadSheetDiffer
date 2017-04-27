import javax.swing.*;
import java.awt.*;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;

public class Main {

    public static void main(String[] args) {
        ArrayList<String> list1 = new ArrayList<>();
        list1.add("1.0.2");
        list1.add("2.0.2");
        list1.add("3.0.1");

        ArrayList<String> list2 = new ArrayList<>();
        list2.add("1.0.2");
        list2.add("4.0.0");

        list1.removeAll(list2);
        list1.addAll(list2);
        System.out.println(list1);
        list1.sort(new Comparator<String>() {
            @Override
            public int compare(String left, String right) {
                String majorVersionLeft = left.split("\\.")[0];
                String majorVersionRight = right.split("\\.")[0];
                return majorVersionRight.compareTo(majorVersionLeft);
            }
        });
        System.out.println(list1);
    }
}

//2.0.1 1.0.1   2.0.1