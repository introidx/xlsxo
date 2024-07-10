package com.introidx.xlsxo;

import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.util.ArrayList;
import java.util.List;

@SpringBootApplication
public class XlsxoApplication {

	public static void main(String[] args) {
		List<Student> students = new ArrayList<>();
		Student student1 = new Student(1L, "John Doe", "john@gmail.com");
		Student student2 = new Student(2L, "Jane Doe", "jobh2@gmail.com");

		students.add(student1);
		students.add(student2);

		try {
			byte[] bytes = XlsxGenerator.generate(students, "Students");
			System.out.println("Excel file generated successfully");
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
