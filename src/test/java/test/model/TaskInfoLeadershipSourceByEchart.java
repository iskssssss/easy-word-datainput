package test.model;

import lombok.Data;

import java.util.LinkedHashMap;

@Data
public class TaskInfoLeadershipSourceByEchart {

    private String month;

    private LinkedHashMap<String,Integer> data = new LinkedHashMap<>();

}
