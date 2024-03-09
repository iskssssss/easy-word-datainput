package test.model;

import lombok.Data;

import java.io.Serializable;

@Data
public class TaskInfoLeadershipCountVo implements Serializable {


    private Integer count;

    private Integer completed;

    private Integer uncompleted;

    private String completionRate;

    private Integer plainTask;

    private Integer encryptedTask;

}
