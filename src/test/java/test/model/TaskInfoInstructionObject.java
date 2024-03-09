package test.model;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
public class TaskInfoInstructionObject {

    private String name;

    private Integer count = 0;

    private List<TaskInfoLeadershipRate> data = new ArrayList<>();

}
