package com.app;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.ComponentScan;

/**
 * @author
 * @create 2018-02-03 16:20
 **/
@SpringBootApplication
@ComponentScan(basePackages = {"com.lance.controller"})
public class App {
    public static void main(String[] args) {
        SpringApplication.run(App.class, args);
    }
}
