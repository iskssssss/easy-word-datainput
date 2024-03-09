package test.model;

import lombok.Data;

import java.io.Serializable;

@Data

public class TaskInfoLeadershipRate implements Serializable {

    private String key;

    private Integer value = 0;

    private Integer rate = 0;
}
