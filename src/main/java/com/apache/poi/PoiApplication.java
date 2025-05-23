package com.apache.poi;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.domain.EntityScan;

@SpringBootApplication
@EntityScan(basePackages = "com.apache.poi.entity")
public class PoiApplication {

	public static void main(String[] args) {
		SpringApplication.run(PoiApplication.class, args);
	}

}
