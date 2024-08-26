package excel_data_driven_test;

import java.io.FileNotFoundException;
import java.util.ArrayList;

public class TestSample {

	public static void main(String[] args) throws FileNotFoundException {
		
		Excel_DataDrivenTest d = new Excel_DataDrivenTest();
		
		ArrayList data = d.getData("Addprofile");
		
		System.out.println(data.get(0));
		System.out.println(data.get(1));
		System.out.println(data.get(2));
		System.out.println(data.get(3));
		
	}

}
